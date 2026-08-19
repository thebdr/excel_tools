using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Projection
{
    /// <summary>
    /// Propagates logical row or column identity strictly through resolved alignment links. Ambiguous or
    /// uncovered child indices remain unmapped, making accidental physical-position fallback impossible.
    /// </summary>
    public static class LogicalAxisIdentityProjector
    {
        private const string RootReasonCode = "identity.root-physical-anchor";
        private const string ParentConfidenceReasonCode = "identity.inherited-parent-confidence";
        private const string AmbiguousReasonCode = "identity.ambiguous-block-no-propagation";
        private const string MissingParentReasonCode = "identity.parent-identity-unresolved";
        private const string UncoveredReasonCode = "identity.physical-index-uncovered";

        public static LogicalAxisIdentityProjection CreateRoot(
            string projectId,
            string logicalSheetId,
            string rootRevisionId,
            LogicalAxisKind axis,
            int physicalCount)
        {
            Require(projectId, nameof(projectId));
            Require(logicalSheetId, nameof(logicalSheetId));
            Require(rootRevisionId, nameof(rootRevisionId));
            if (physicalCount < 0) throw new ArgumentOutOfRangeException(nameof(physicalCount));

            var assignments = new Dictionary<int, LogicalAxisIdentityAssignment>();
            var rootReasons = new[]
            {
                new AlignmentReason(
                    RootReasonCode,
                    "The logical identity is anchored to a physical index in the selected root revision."),
            };
            for (int index = 0; index < physicalCount; index++)
            {
                string logicalId = StableId.Compute(
                    "logical-axis-identity-root-v1",
                    projectId,
                    logicalSheetId,
                    AxisCode(axis),
                    rootRevisionId,
                    Number(index));
                assignments.Add(index, CreateAssignment(
                    logicalId,
                    rootRevisionId,
                    index,
                    LogicalAxisIdentitySource.Root,
                    rootRevisionId,
                    index,
                    null,
                    AlignmentConfidence.Confirmed,
                    AxisAlignmentStrategy.None,
                    rootReasons));
            }

            string projectionId = ComputeProjectionId(
                projectId,
                logicalSheetId,
                rootRevisionId,
                "",
                axis,
                physicalCount,
                AlignmentConfidence.Confirmed,
                assignments,
                Array.Empty<LogicalAxisStructuralEvent>(),
                Array.Empty<LogicalAxisIdentityAmbiguity>());
            return new LogicalAxisIdentityProjection(
                projectionId,
                projectId,
                logicalSheetId,
                rootRevisionId,
                "",
                "",
                axis,
                physicalCount,
                AlignmentConfidence.Confirmed,
                assignments,
                Array.Empty<LogicalAxisStructuralEvent>(),
                Array.Empty<LogicalAxisIdentityAmbiguity>());
        }

        public static LogicalAxisIdentityProjection Propagate(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            int childPhysicalCount,
            AxisAlignmentResult alignment)
        {
            if (parent == null) throw new ArgumentNullException(nameof(parent));
            Require(childRevisionId, nameof(childRevisionId));
            if (string.Equals(parent.RevisionId, childRevisionId, StringComparison.Ordinal))
                throw new ArgumentException("Parent and child revision IDs must differ.", nameof(childRevisionId));
            if (childPhysicalCount < 0) throw new ArgumentOutOfRangeException(nameof(childPhysicalCount));
            if (alignment == null) throw new ArgumentNullException(nameof(alignment));

            ValidateAlignmentBounds(parent.PhysicalCount, childPhysicalCount, alignment);

            var assignments = new Dictionary<int, LogicalAxisIdentityAssignment>();
            var events = new List<LogicalAxisStructuralEvent>();
            var ambiguities = new List<LogicalAxisIdentityAmbiguity>();
            var coveredBefore = new HashSet<int>();
            var coveredAfter = new HashSet<int>();

            foreach (AxisAlignmentAmbiguity block in OrderAmbiguities(alignment.Ambiguities))
            {
                MarkCovered(block.BeforeIndices, coveredBefore);
                MarkCovered(block.AfterIndices, coveredAfter);
                var reasons = AppendReason(
                    block.Reasons,
                    new AlignmentReason(
                        AmbiguousReasonCode,
                        "No parent-child logical identity was created for this ambiguous alignment block."));
                ambiguities.Add(CreateAmbiguity(
                    parent,
                    childRevisionId,
                    LogicalAxisAmbiguityKind.AlignmentBlock,
                    block.BeforeIndices,
                    block.AfterIndices,
                    block.EvidenceToken,
                    alignment.Strategy,
                    reasons));
            }

            foreach (AxisAlignmentLink link in OrderLinks(alignment.Links))
            {
                if (link.BeforeIndex.HasValue) coveredBefore.Add(link.BeforeIndex.Value);
                if (link.AfterIndex.HasValue) coveredAfter.Add(link.AfterIndex.Value);

                switch (link.Kind)
                {
                    case AxisAlignmentKind.Matched:
                        ProjectMatch(parent, childRevisionId, link, assignments, events, ambiguities);
                        break;
                    case AxisAlignmentKind.OnlyAfter:
                        ProjectInsertion(parent, childRevisionId, link, assignments, events);
                        break;
                    case AxisAlignmentKind.OnlyBefore:
                        ProjectDeletion(parent, childRevisionId, link, events, ambiguities);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(alignment), "Unknown alignment-link kind.");
                }
            }

            for (int index = 0; index < parent.PhysicalCount; index++)
            {
                if (coveredBefore.Contains(index)) continue;
                ambiguities.Add(CreateUncoveredAmbiguity(parent, childRevisionId, index, true, alignment.Strategy));
            }
            for (int index = 0; index < childPhysicalCount; index++)
            {
                if (coveredAfter.Contains(index)) continue;
                ambiguities.Add(CreateUncoveredAmbiguity(parent, childRevisionId, index, false, alignment.Strategy));
            }

            events = events
                .OrderBy(item => item.AfterIndex ?? int.MaxValue)
                .ThenBy(item => item.BeforeIndex ?? int.MaxValue)
                .ThenBy(item => item.Kind)
                .ThenBy(item => item.EventId, StringComparer.Ordinal)
                .ToList();
            ambiguities = ambiguities
                .OrderBy(item => item.AfterIndices.Count == 0 ? int.MaxValue : item.AfterIndices[0])
                .ThenBy(item => item.BeforeIndices.Count == 0 ? int.MaxValue : item.BeforeIndices[0])
                .ThenBy(item => item.AmbiguityId, StringComparer.Ordinal)
                .ToList();

            AlignmentConfidence confidence = AggregateConfidence(assignments.Values, events, ambiguities);
            string projectionId = ComputeProjectionId(
                parent.ProjectId,
                parent.LogicalSheetId,
                childRevisionId,
                parent.ProjectionId,
                parent.Axis,
                childPhysicalCount,
                confidence,
                assignments,
                events,
                ambiguities);
            return new LogicalAxisIdentityProjection(
                projectionId,
                parent.ProjectId,
                parent.LogicalSheetId,
                childRevisionId,
                parent.RevisionId,
                parent.ProjectionId,
                parent.Axis,
                childPhysicalCount,
                confidence,
                assignments,
                events,
                ambiguities);
        }

        private static void ProjectMatch(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            AxisAlignmentLink link,
            IDictionary<int, LogicalAxisIdentityAssignment> assignments,
            ICollection<LogicalAxisStructuralEvent> events,
            ICollection<LogicalAxisIdentityAmbiguity> ambiguities)
        {
            int before = link.BeforeIndex.Value;
            int after = link.AfterIndex.Value;
            if (!parent.AssignmentsByPhysicalIndex.TryGetValue(before, out LogicalAxisIdentityAssignment inherited))
            {
                var reasons = AppendReason(
                    link.Reasons,
                    new AlignmentReason(
                        MissingParentReasonCode,
                        "The alignment link is resolved, but its parent physical index has no logical identity."));
                ambiguities.Add(CreateAmbiguity(
                    parent,
                    childRevisionId,
                    LogicalAxisAmbiguityKind.MissingParentIdentity,
                    new[] { before },
                    new[] { after },
                    "",
                    link.Strategy,
                    reasons));
                if (link.Reordered)
                    events.Add(CreateEvent(
                        parent,
                        childRevisionId,
                        LogicalAxisStructuralEventKind.Reorder,
                        "",
                        before,
                        after,
                        AlignmentConfidence.None,
                        link.Strategy,
                        reasons));
                return;
            }

            AlignmentConfidence confidence = Minimum(inherited.Confidence, link.Confidence);
            IReadOnlyList<AlignmentReason> assignmentReasons = link.Reasons;
            if (confidence != link.Confidence)
                assignmentReasons = AppendReason(
                    assignmentReasons,
                    new AlignmentReason(
                        ParentConfidenceReasonCode,
                        "Identity confidence remains limited by the parent assignment."));
            assignments.Add(after, CreateAssignment(
                inherited.LogicalIdentityId,
                childRevisionId,
                after,
                LogicalAxisIdentitySource.Inherited,
                inherited.OriginRevisionId,
                inherited.OriginPhysicalIndex,
                before,
                confidence,
                link.Strategy,
                assignmentReasons));

            if (link.Reordered)
                events.Add(CreateEvent(
                    parent,
                    childRevisionId,
                    LogicalAxisStructuralEventKind.Reorder,
                    inherited.LogicalIdentityId,
                    before,
                    after,
                    confidence,
                    link.Strategy,
                    link.Reasons));
        }

        private static void ProjectInsertion(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            AxisAlignmentLink link,
            IDictionary<int, LogicalAxisIdentityAssignment> assignments,
            ICollection<LogicalAxisStructuralEvent> events)
        {
            int after = link.AfterIndex.Value;
            string logicalId = StableId.Compute(
                "logical-axis-identity-insertion-v1",
                parent.ProjectId,
                parent.LogicalSheetId,
                AxisCode(parent.Axis),
                childRevisionId,
                Number(after));
            assignments.Add(after, CreateAssignment(
                logicalId,
                childRevisionId,
                after,
                LogicalAxisIdentitySource.Inserted,
                childRevisionId,
                after,
                null,
                link.Confidence,
                link.Strategy,
                link.Reasons));
            events.Add(CreateEvent(
                parent,
                childRevisionId,
                LogicalAxisStructuralEventKind.Insertion,
                logicalId,
                null,
                after,
                link.Confidence,
                link.Strategy,
                link.Reasons));
        }

        private static void ProjectDeletion(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            AxisAlignmentLink link,
            ICollection<LogicalAxisStructuralEvent> events,
            ICollection<LogicalAxisIdentityAmbiguity> ambiguities)
        {
            int before = link.BeforeIndex.Value;
            if (parent.AssignmentsByPhysicalIndex.TryGetValue(before, out LogicalAxisIdentityAssignment deleted))
            {
                events.Add(CreateEvent(
                    parent,
                    childRevisionId,
                    LogicalAxisStructuralEventKind.Deletion,
                    deleted.LogicalIdentityId,
                    before,
                    null,
                    Minimum(deleted.Confidence, link.Confidence),
                    link.Strategy,
                    link.Reasons));
                return;
            }

            var reasons = AppendReason(
                link.Reasons,
                new AlignmentReason(
                    MissingParentReasonCode,
                    "The deleted parent physical index already had unresolved logical identity."));
            events.Add(CreateEvent(
                parent,
                childRevisionId,
                LogicalAxisStructuralEventKind.Deletion,
                "",
                before,
                null,
                AlignmentConfidence.None,
                link.Strategy,
                reasons));
            ambiguities.Add(CreateAmbiguity(
                parent,
                childRevisionId,
                LogicalAxisAmbiguityKind.MissingParentIdentity,
                new[] { before },
                Array.Empty<int>(),
                "",
                link.Strategy,
                reasons));
        }

        private static LogicalAxisIdentityAmbiguity CreateUncoveredAmbiguity(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            int physicalIndex,
            bool before,
            AxisAlignmentStrategy strategy)
        {
            var reasons = new[]
            {
                new AlignmentReason(
                    UncoveredReasonCode,
                    "The alignment result did not cover this physical axis index; no identity was inferred."),
            };
            return CreateAmbiguity(
                parent,
                childRevisionId,
                LogicalAxisAmbiguityKind.UncoveredPhysicalIndex,
                before ? new[] { physicalIndex } : Array.Empty<int>(),
                before ? Array.Empty<int>() : new[] { physicalIndex },
                "",
                strategy,
                reasons);
        }

        private static LogicalAxisIdentityAmbiguity CreateAmbiguity(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            LogicalAxisAmbiguityKind kind,
            IEnumerable<int> beforeIndices,
            IEnumerable<int> afterIndices,
            string evidenceToken,
            AxisAlignmentStrategy strategy,
            IEnumerable<AlignmentReason> reasons)
        {
            int[] before = (beforeIndices ?? Array.Empty<int>()).OrderBy(value => value).ToArray();
            int[] after = (afterIndices ?? Array.Empty<int>()).OrderBy(value => value).ToArray();
            var known = new Dictionary<int, string>();
            foreach (int index in before)
                if (parent.IdentityByPhysicalIndex.TryGetValue(index, out string identity)) known[index] = identity;

            var components = new List<string>
            {
                parent.ProjectId,
                parent.LogicalSheetId,
                AxisCode(parent.Axis),
                parent.RevisionId,
                childRevisionId,
                Number(kind),
                evidenceToken ?? "",
                Number(strategy),
            };
            AddIndices(components, "before", before);
            AddIndices(components, "after", after);
            foreach (KeyValuePair<int, string> pair in known.OrderBy(item => item.Key))
            {
                components.Add("known");
                components.Add(Number(pair.Key));
                components.Add(pair.Value);
            }
            AddReasonCodes(components, reasons);
            string ambiguityId = StableId.Compute("logical-axis-ambiguity-v1", components.ToArray());
            return new LogicalAxisIdentityAmbiguity(
                ambiguityId,
                kind,
                before,
                after,
                known,
                evidenceToken,
                AlignmentConfidence.None,
                strategy,
                reasons);
        }

        private static LogicalAxisIdentityAssignment CreateAssignment(
            string logicalIdentityId,
            string revisionId,
            int physicalIndex,
            LogicalAxisIdentitySource source,
            string originRevisionId,
            int originPhysicalIndex,
            int? parentPhysicalIndex,
            AlignmentConfidence confidence,
            AxisAlignmentStrategy strategy,
            IEnumerable<AlignmentReason> reasons)
        {
            var components = new List<string>
            {
                logicalIdentityId,
                revisionId,
                Number(physicalIndex),
                Number(source),
                originRevisionId,
                Number(originPhysicalIndex),
                parentPhysicalIndex.HasValue ? Number(parentPhysicalIndex.Value) : null,
                Number(confidence),
                Number(strategy),
            };
            AddReasonCodes(components, reasons);
            string assignmentId = StableId.Compute("logical-axis-assignment-v1", components.ToArray());
            return new LogicalAxisIdentityAssignment(
                assignmentId,
                logicalIdentityId,
                revisionId,
                physicalIndex,
                source,
                originRevisionId,
                originPhysicalIndex,
                parentPhysicalIndex,
                confidence,
                strategy,
                reasons);
        }

        private static LogicalAxisStructuralEvent CreateEvent(
            LogicalAxisIdentityProjection parent,
            string childRevisionId,
            LogicalAxisStructuralEventKind kind,
            string logicalIdentityId,
            int? beforeIndex,
            int? afterIndex,
            AlignmentConfidence confidence,
            AxisAlignmentStrategy strategy,
            IEnumerable<AlignmentReason> reasons)
        {
            var components = new List<string>
            {
                parent.ProjectId,
                parent.LogicalSheetId,
                AxisCode(parent.Axis),
                parent.RevisionId,
                childRevisionId,
                Number(kind),
                logicalIdentityId ?? "",
                beforeIndex.HasValue ? Number(beforeIndex.Value) : null,
                afterIndex.HasValue ? Number(afterIndex.Value) : null,
                Number(confidence),
                Number(strategy),
            };
            AddReasonCodes(components, reasons);
            string eventId = StableId.Compute("logical-axis-structural-event-v1", components.ToArray());
            return new LogicalAxisStructuralEvent(
                eventId,
                kind,
                parent.RevisionId,
                childRevisionId,
                logicalIdentityId,
                beforeIndex,
                afterIndex,
                confidence,
                strategy,
                reasons);
        }

        private static string ComputeProjectionId(
            string projectId,
            string logicalSheetId,
            string revisionId,
            string parentProjectionId,
            LogicalAxisKind axis,
            int physicalCount,
            AlignmentConfidence confidence,
            IDictionary<int, LogicalAxisIdentityAssignment> assignments,
            IEnumerable<LogicalAxisStructuralEvent> events,
            IEnumerable<LogicalAxisIdentityAmbiguity> ambiguities)
        {
            var components = new List<string>
            {
                projectId,
                logicalSheetId,
                revisionId,
                parentProjectionId ?? "",
                AxisCode(axis),
                Number(physicalCount),
                Number(confidence),
            };
            foreach (KeyValuePair<int, LogicalAxisIdentityAssignment> pair in assignments.OrderBy(item => item.Key))
            {
                components.Add("assignment");
                components.Add(Number(pair.Key));
                components.Add(pair.Value.AssignmentId);
            }
            foreach (LogicalAxisStructuralEvent item in events.OrderBy(item => item.EventId, StringComparer.Ordinal))
            {
                components.Add("event");
                components.Add(item.EventId);
            }
            foreach (LogicalAxisIdentityAmbiguity item in ambiguities.OrderBy(item => item.AmbiguityId, StringComparer.Ordinal))
            {
                components.Add("ambiguity");
                components.Add(item.AmbiguityId);
            }
            return StableId.Compute("logical-axis-projection-v1", components.ToArray());
        }

        private static AlignmentConfidence AggregateConfidence(
            IEnumerable<LogicalAxisIdentityAssignment> assignments,
            IEnumerable<LogicalAxisStructuralEvent> events,
            IReadOnlyCollection<LogicalAxisIdentityAmbiguity> ambiguities)
        {
            if (ambiguities.Count > 0) return AlignmentConfidence.None;
            AlignmentConfidence result = AlignmentConfidence.Confirmed;
            bool any = false;
            foreach (LogicalAxisIdentityAssignment assignment in assignments)
            {
                result = Minimum(result, assignment.Confidence);
                any = true;
            }
            foreach (LogicalAxisStructuralEvent item in events)
            {
                result = Minimum(result, item.Confidence);
                any = true;
            }
            return any ? result : AlignmentConfidence.Confirmed;
        }

        private static void ValidateAlignmentBounds(
            int parentPhysicalCount,
            int childPhysicalCount,
            AxisAlignmentResult alignment)
        {
            var before = new HashSet<int>();
            var after = new HashSet<int>();
            foreach (AxisAlignmentLink link in alignment.Links)
            {
                ValidateAndClaim(link.BeforeIndex, parentPhysicalCount, before, "before");
                ValidateAndClaim(link.AfterIndex, childPhysicalCount, after, "after");
            }
            foreach (AxisAlignmentAmbiguity ambiguity in alignment.Ambiguities)
            {
                foreach (int index in ambiguity.BeforeIndices)
                    ValidateAndClaim(index, parentPhysicalCount, before, "before");
                foreach (int index in ambiguity.AfterIndices)
                    ValidateAndClaim(index, childPhysicalCount, after, "after");
            }
        }

        private static void ValidateAndClaim(
            int? index,
            int count,
            ISet<int> claimed,
            string side)
        {
            if (!index.HasValue) return;
            if (index.Value < 0 || index.Value >= count)
                throw new ArgumentOutOfRangeException(
                    nameof(index),
                    "The alignment contains an out-of-range " + side + " physical index.");
            if (!claimed.Add(index.Value))
                throw new ArgumentException(
                    "The alignment assigns the same " + side + " physical index more than once.");
        }

        private static IEnumerable<AxisAlignmentLink> OrderLinks(IEnumerable<AxisAlignmentLink> links) =>
            links.OrderBy(item => item.AfterIndex ?? int.MaxValue)
                .ThenBy(item => item.BeforeIndex ?? int.MaxValue)
                .ThenBy(item => item.Kind);

        private static IEnumerable<AxisAlignmentAmbiguity> OrderAmbiguities(
            IEnumerable<AxisAlignmentAmbiguity> ambiguities) =>
            ambiguities.OrderBy(item => item.AfterIndices.Count == 0 ? int.MaxValue : item.AfterIndices[0])
                .ThenBy(item => item.BeforeIndices.Count == 0 ? int.MaxValue : item.BeforeIndices[0])
                .ThenBy(item => item.EvidenceToken, StringComparer.Ordinal);

        private static void MarkCovered(IEnumerable<int> indices, ISet<int> covered)
        {
            foreach (int index in indices) covered.Add(index);
        }

        private static IReadOnlyList<AlignmentReason> AppendReason(
            IEnumerable<AlignmentReason> reasons,
            AlignmentReason additional)
        {
            var result = new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()) { additional };
            return result;
        }

        private static void AddReasonCodes(ICollection<string> components, IEnumerable<AlignmentReason> reasons)
        {
            foreach (AlignmentReason reason in (reasons ?? Array.Empty<AlignmentReason>())
                         .OrderBy(item => item.Code, StringComparer.Ordinal))
            {
                components.Add("reason");
                components.Add(reason.Code);
                components.Add(reason.ScoreImpact.ToString("R", CultureInfo.InvariantCulture));
            }
        }

        private static void AddIndices(ICollection<string> components, string label, IEnumerable<int> indices)
        {
            foreach (int index in indices)
            {
                components.Add(label);
                components.Add(Number(index));
            }
        }

        private static AlignmentConfidence Minimum(AlignmentConfidence left, AlignmentConfidence right) =>
            (AlignmentConfidence)Math.Min((int)left, (int)right);

        private static string AxisCode(LogicalAxisKind axis)
        {
            switch (axis)
            {
                case LogicalAxisKind.Row: return "row";
                case LogicalAxisKind.Column: return "column";
                default: throw new ArgumentOutOfRangeException(nameof(axis));
            }
        }

        private static string Number<T>(T value) where T : struct =>
            Convert.ToInt32(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);

        private static void Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A stable identity component is required.", name);
        }
    }
}
