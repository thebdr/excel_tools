using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Revision;

namespace ExcelPowerTools.Core.Comparison.Lineage
{
    public enum LineageConfidenceLevel
    {
        None = 0,
        Weak = 1,
        Moderate = 2,
        Strong = 3,
        Certain = 4,
    }

    public enum LineageReasonCode
    {
        ExplicitRevisionToken = 0,
        WeakFilenameToken = 1,
        AmbiguousFilenameToken = 2,
        TerminalNumberAmbiguity = 3,
        ExactDuplicateContent = 4,
        ModifiedTimestamp = 5,
        CreationTimestamp = 6,
        ConflictingTimestamp = 7,
        DeterministicTieBreak = 8,
        FirstRevisionHasNoParent = 9,
        FilenameParentProposal = 10,
        WeakTimestampParentProposal = 11,
        AmbiguousParent = 12,
        NoParentEvidence = 13,
        ExplicitMasterSelection = 14,
        ExplicitBaselineSelection = 15,
        SelectionRequired = 16,
    }

    /// <summary>A calibrated evidence strength. Any proposal below Strong remains confirmation-required.</summary>
    public sealed class LineageConfidence
    {
        internal LineageConfidence(double score)
        {
            if (double.IsNaN(score) || double.IsInfinity(score) || score < 0 || score > 1)
                throw new ArgumentOutOfRangeException(nameof(score));
            Score = score;
            Level = score >= 0.99
                ? LineageConfidenceLevel.Certain
                : score >= 0.85
                    ? LineageConfidenceLevel.Strong
                    : score >= 0.50
                        ? LineageConfidenceLevel.Moderate
                        : score > 0
                            ? LineageConfidenceLevel.Weak
                            : LineageConfidenceLevel.None;
        }

        public double Score { get; }
        public LineageConfidenceLevel Level { get; }
        public bool RequiresConfirmation => Level < LineageConfidenceLevel.Strong;
    }

    /// <summary>Structured, non-localized evidence behind one inference decision or refusal to decide.</summary>
    public sealed class LineageReason
    {
        internal LineageReason(
            LineageReasonCode code,
            double confidence,
            string message,
            IEnumerable<string> relatedRevisionIds = null)
        {
            Code = code;
            Confidence = new LineageConfidence(confidence);
            Message = message ?? "";
            RelatedRevisionIds = new ReadOnlyCollection<string>(
                new List<string>(relatedRevisionIds ?? Array.Empty<string>()));
        }

        public LineageReasonCode Code { get; }
        public LineageConfidence Confidence { get; }
        public string Message { get; }
        public IReadOnlyList<string> RelatedRevisionIds { get; }
    }

    /// <summary>
    /// Immutable occurrence evidence supplied to lineage inference. FileName is reduced to a basename so this
    /// portable evidence object cannot accidentally retain a local absolute path.
    /// </summary>
    public sealed class LineageRevisionEvidence
    {
        public LineageRevisionEvidence(
            string revisionId,
            string fileName,
            string byteContentId,
            DateTimeOffset? createdUtc = null,
            DateTimeOffset? modifiedUtc = null,
            string description = "")
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("A source filename is required.", nameof(fileName));
            if (!StableId.IsSha256(byteContentId))
                throw new ArgumentException("A byte-content SHA-256 is required.", nameof(byteContentId));

            string baseName = Path.GetFileName(fileName);
            if (string.IsNullOrWhiteSpace(baseName))
                throw new ArgumentException("A source filename is required.", nameof(fileName));

            RevisionId = revisionId;
            FileName = baseName;
            ByteContentId = byteContentId.ToUpperInvariant();
            CreatedUtc = createdUtc;
            ModifiedUtc = modifiedUtc;
            Description = description ?? "";
        }

        public string RevisionId { get; }
        public string FileName { get; }
        public string ByteContentId { get; }
        public DateTimeOffset? CreatedUtc { get; }
        public DateTimeOffset? ModifiedUtc { get; }
        public string Description { get; }
    }

    /// <summary>Master and trusted baseline are user assertions; the inference engine never invents them.</summary>
    public sealed class LineageExplicitSelections
    {
        public LineageExplicitSelections(string masterRevisionId = null, string trustedBaselineRevisionId = null)
        {
            ValidateOptional(masterRevisionId, nameof(masterRevisionId));
            ValidateOptional(trustedBaselineRevisionId, nameof(trustedBaselineRevisionId));
            MasterRevisionId = masterRevisionId;
            TrustedBaselineRevisionId = trustedBaselineRevisionId;
        }

        public string MasterRevisionId { get; }
        public string TrustedBaselineRevisionId { get; }

        private static void ValidateOptional(string value, string name)
        {
            if (value != null && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An explicit selection must be null or a non-blank revision ID.", name);
        }
    }

    public sealed class LineageDuplicateGroup
    {
        internal LineageDuplicateGroup(string byteContentId, IEnumerable<string> revisionIds)
        {
            ByteContentId = byteContentId;
            RevisionIds = new ReadOnlyCollection<string>(new List<string>(revisionIds));
            GroupId = StableId.Compute(
                "lineage-exact-duplicate-group-v1",
                new[] { byteContentId }.Concat(RevisionIds).ToArray());
        }

        public string GroupId { get; }
        public string ByteContentId { get; }
        public IReadOnlyList<string> RevisionIds { get; }
    }

    public sealed class LineageRevisionProposal
    {
        internal LineageRevisionProposal(
            LineageRevisionEvidence evidence,
            RevisionLabelInferenceResult labelInference,
            int proposedOrdinal,
            string proposedParentRevisionId,
            double orderConfidence,
            double parentConfidence,
            IEnumerable<string> exactDuplicateRevisionIds,
            IEnumerable<LineageReason> orderReasons,
            IEnumerable<LineageReason> parentReasons)
        {
            Evidence = evidence;
            LabelInference = labelInference;
            ProposedOrdinal = proposedOrdinal;
            ProposedParentRevisionId = proposedParentRevisionId;
            OrderConfidence = new LineageConfidence(orderConfidence);
            ParentConfidence = new LineageConfidence(parentConfidence);
            ExactDuplicateRevisionIds = new ReadOnlyCollection<string>(
                new List<string>(exactDuplicateRevisionIds ?? Array.Empty<string>()));
            OrderReasons = new ReadOnlyCollection<LineageReason>(
                new List<LineageReason>(orderReasons ?? Array.Empty<LineageReason>()));
            ParentReasons = new ReadOnlyCollection<LineageReason>(
                new List<LineageReason>(parentReasons ?? Array.Empty<LineageReason>()));
        }

        public string RevisionId => Evidence.RevisionId;
        public LineageRevisionEvidence Evidence { get; }
        public RevisionLabelInferenceResult LabelInference { get; }
        /// <summary>One-based proposed display order.</summary>
        public int ProposedOrdinal { get; }
        public string ProposedParentRevisionId { get; }
        public LineageConfidence OrderConfidence { get; }
        public LineageConfidence ParentConfidence { get; }
        public IReadOnlyList<string> ExactDuplicateRevisionIds { get; }
        public IReadOnlyList<LineageReason> OrderReasons { get; }
        public IReadOnlyList<LineageReason> ParentReasons { get; }
    }

    public sealed class LineageInferenceResult
    {
        private readonly IReadOnlyDictionary<string, LineageRevisionProposal> _byRevisionId;

        internal LineageInferenceResult(
            string proposalGenerationId,
            IEnumerable<LineageRevisionProposal> proposals,
            IEnumerable<LineageDuplicateGroup> duplicateGroups,
            string explicitMasterRevisionId,
            string explicitTrustedBaselineRevisionId,
            IEnumerable<LineageReason> reasons)
        {
            ProposalGenerationId = proposalGenerationId;
            Proposals = new ReadOnlyCollection<LineageRevisionProposal>(new List<LineageRevisionProposal>(proposals));
            DuplicateGroups = new ReadOnlyCollection<LineageDuplicateGroup>(new List<LineageDuplicateGroup>(duplicateGroups));
            ExplicitMasterRevisionId = explicitMasterRevisionId;
            ExplicitTrustedBaselineRevisionId = explicitTrustedBaselineRevisionId;
            Reasons = new ReadOnlyCollection<LineageReason>(new List<LineageReason>(reasons));
            _byRevisionId = new ReadOnlyDictionary<string, LineageRevisionProposal>(
                Proposals.ToDictionary(proposal => proposal.RevisionId, StringComparer.Ordinal));
        }

        public string ProposalGenerationId { get; }
        public IReadOnlyList<LineageRevisionProposal> Proposals { get; }
        public IReadOnlyList<LineageDuplicateGroup> DuplicateGroups { get; }
        public string ExplicitMasterRevisionId { get; }
        public string ExplicitTrustedBaselineRevisionId { get; }
        public IReadOnlyList<LineageReason> Reasons { get; }

        public LineageRevisionProposal GetProposal(string revisionId)
        {
            if (revisionId == null) throw new ArgumentNullException(nameof(revisionId));
            if (!_byRevisionId.TryGetValue(revisionId, out LineageRevisionProposal proposal))
                throw new KeyNotFoundException("Unknown revision occurrence ID: " + revisionId);
            return proposal;
        }
    }

    /// <summary>
    /// Conservative multi-revision lineage inference. Strong filename tokens constrain relative order;
    /// timestamps only break otherwise-unresolved order and produce weak parent proposals. Exact hashes prove
    /// content identity but never chronology or ancestry. A deterministic tie-break makes proposals reproducible
    /// without disguising uncertainty as evidence.
    /// </summary>
    public static class MultiRevisionLineageInference
    {
        public static LineageInferenceResult Infer(
            IEnumerable<LineageRevisionEvidence> revisions,
            LineageExplicitSelections selections = null)
        {
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));
            List<LineageRevisionEvidence> evidence = revisions.ToList();
            if (evidence.Count == 0)
                throw new ArgumentException("At least one revision occurrence is required.", nameof(revisions));

            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (LineageRevisionEvidence item in evidence)
            {
                if (item == null) throw new ArgumentException("Revision evidence cannot contain null.", nameof(revisions));
                if (!ids.Add(item.RevisionId))
                    throw new ArgumentException("Duplicate revision occurrence ID: " + item.RevisionId, nameof(revisions));
            }

            selections = selections ?? new LineageExplicitSelections();
            ValidateSelection(selections.MasterRevisionId, ids, "master");
            ValidateSelection(selections.TrustedBaselineRevisionId, ids, "trusted baseline");

            List<WorkItem> work = evidence
                .Select(item => new WorkItem(item, RevisionLabelInference.Infer(item.FileName)))
                .OrderBy(item => item.Evidence.RevisionId, StringComparer.Ordinal)
                .ToList();
            List<LineageDuplicateGroup> duplicateGroups = BuildDuplicateGroups(work);
            HashSet<string> timestampConflicts = FindTimestampConflicts(work);
            List<WorkItem> ordered = StableTopologicalOrder(work);
            List<LineageRevisionProposal> proposals = BuildProposals(ordered, timestampConflicts);
            List<LineageReason> reasons = BuildGlobalReasons(selections);
            string proposalGenerationId = ComputeProposalGenerationId(evidence, selections);

            return new LineageInferenceResult(
                proposalGenerationId,
                proposals,
                duplicateGroups,
                selections.MasterRevisionId,
                selections.TrustedBaselineRevisionId,
                reasons);
        }

        private static List<LineageDuplicateGroup> BuildDuplicateGroups(List<WorkItem> work)
        {
            var groups = new List<LineageDuplicateGroup>();
            foreach (IGrouping<string, WorkItem> hashGroup in work
                         .GroupBy(item => item.Evidence.ByteContentId, StringComparer.Ordinal)
                         .OrderBy(group => group.Key, StringComparer.Ordinal))
            {
                List<WorkItem> members = hashGroup.OrderBy(
                    item => item.Evidence.RevisionId,
                    StringComparer.Ordinal).ToList();
                if (members.Count < 2) continue;
                string[] revisionIds = members.Select(item => item.Evidence.RevisionId).ToArray();
                groups.Add(new LineageDuplicateGroup(hashGroup.Key, revisionIds));
                foreach (WorkItem member in members)
                {
                    member.DuplicateRevisionIds = revisionIds
                        .Where(id => !string.Equals(id, member.Evidence.RevisionId, StringComparison.Ordinal))
                        .ToArray();
                }
            }
            return groups;
        }

        private static HashSet<string> FindTimestampConflicts(List<WorkItem> work)
        {
            var conflicts = new HashSet<string>(StringComparer.Ordinal);
            for (int leftIndex = 0; leftIndex < work.Count; leftIndex++)
            {
                WorkItem left = work[leftIndex];
                if (left.StrongToken == null || !left.EffectiveTimestamp.HasValue) continue;
                for (int rightIndex = leftIndex + 1; rightIndex < work.Count; rightIndex++)
                {
                    WorkItem right = work[rightIndex];
                    if (right.StrongToken == null || !right.EffectiveTimestamp.HasValue) continue;
                    int tokenOrder = left.StrongToken.Number.CompareTo(right.StrongToken.Number);
                    int timeOrder = left.EffectiveTimestamp.Value.CompareTo(right.EffectiveTimestamp.Value);
                    if (tokenOrder != 0 && timeOrder != 0 && Math.Sign(tokenOrder) != Math.Sign(timeOrder))
                    {
                        conflicts.Add(left.Evidence.RevisionId);
                        conflicts.Add(right.Evidence.RevisionId);
                        left.ConflictingRevisionIds.Add(right.Evidence.RevisionId);
                        right.ConflictingRevisionIds.Add(left.Evidence.RevisionId);
                    }
                }
            }
            return conflicts;
        }

        private static List<WorkItem> StableTopologicalOrder(List<WorkItem> work)
        {
            var edges = work.ToDictionary(
                item => item.Evidence.RevisionId,
                item => new HashSet<string>(StringComparer.Ordinal),
                StringComparer.Ordinal);
            var indegree = work.ToDictionary(
                item => item.Evidence.RevisionId,
                item => 0,
                StringComparer.Ordinal);

            for (int leftIndex = 0; leftIndex < work.Count; leftIndex++)
            {
                WorkItem left = work[leftIndex];
                if (left.StrongToken == null) continue;
                for (int rightIndex = leftIndex + 1; rightIndex < work.Count; rightIndex++)
                {
                    WorkItem right = work[rightIndex];
                    if (right.StrongToken == null) continue;
                    int comparison = left.StrongToken.Number.CompareTo(right.StrongToken.Number);
                    if (comparison == 0) continue;
                    WorkItem before = comparison < 0 ? left : right;
                    WorkItem after = comparison < 0 ? right : left;
                    if (edges[before.Evidence.RevisionId].Add(after.Evidence.RevisionId))
                        indegree[after.Evidence.RevisionId]++;
                }
            }

            var remaining = new Dictionary<string, WorkItem>(
                work.ToDictionary(item => item.Evidence.RevisionId, StringComparer.Ordinal),
                StringComparer.Ordinal);
            var result = new List<WorkItem>(work.Count);
            while (remaining.Count > 0)
            {
                WorkItem next = remaining.Values
                    .Where(item => indegree[item.Evidence.RevisionId] == 0)
                    .OrderBy(item => item, WorkItemTieBreaker.Instance)
                    .FirstOrDefault();
                if (next == null)
                    throw new InvalidOperationException("Strong revision-token constraints unexpectedly formed a cycle.");

                result.Add(next);
                remaining.Remove(next.Evidence.RevisionId);
                foreach (string childId in edges[next.Evidence.RevisionId]) indegree[childId]--;
            }
            return result;
        }

        private static List<LineageRevisionProposal> BuildProposals(
            List<WorkItem> ordered,
            HashSet<string> timestampConflicts)
        {
            var proposals = new List<LineageRevisionProposal>(ordered.Count);
            for (int index = 0; index < ordered.Count; index++)
            {
                WorkItem item = ordered[index];
                var orderReasons = BuildOrderReasons(item);
                double orderConfidence = OrderConfidence(item, ordered, timestampConflicts);
                ParentProposal parent = InferParent(item, index, ordered);
                proposals.Add(new LineageRevisionProposal(
                    item.Evidence,
                    item.LabelInference,
                    index + 1,
                    parent.ParentRevisionId,
                    orderConfidence,
                    parent.Confidence,
                    item.DuplicateRevisionIds,
                    orderReasons,
                    parent.Reasons));
            }
            return proposals;
        }

        private static List<LineageReason> BuildOrderReasons(WorkItem item)
        {
            var reasons = new List<LineageReason>();
            if (item.StrongToken != null)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.ExplicitRevisionToken,
                    item.StrongToken.Confidence,
                    "Filename token '" + item.StrongToken.RawToken + "' explicitly proposes revision " +
                    item.StrongToken.Number + "."));
            }
            else if (item.LabelInference.Candidates.Count > 0)
            {
                bool terminalAmbiguity = item.LabelInference.Candidates.Count > 1 &&
                    item.LabelInference.Candidates.All(candidate =>
                        candidate.Kind == RevisionTokenKind.TrailingWholeNumber ||
                        candidate.Kind == RevisionTokenKind.TrailingSplitDecimal);
                string alternatives = string.Join(
                    ", ",
                    item.LabelInference.Candidates.Select(candidate =>
                        candidate.Number + " (" + candidate.Kind + ")"));
                reasons.Add(new LineageReason(
                    terminalAmbiguity
                        ? LineageReasonCode.TerminalNumberAmbiguity
                        : item.LabelInference.Candidates.Count == 1
                            ? LineageReasonCode.WeakFilenameToken
                            : LineageReasonCode.AmbiguousFilenameToken,
                    item.LabelInference.Candidates.Max(candidate => candidate.Confidence),
                    "Filename evidence remains unselected: " + alternatives + "."));
            }

            if (item.Evidence.ModifiedUtc.HasValue)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.ModifiedTimestamp,
                    0.25,
                    "Filesystem modified time " + Utc(item.Evidence.ModifiedUtc.Value) +
                    " is weak ordering evidence only."));
            }
            if (item.Evidence.CreatedUtc.HasValue)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.CreationTimestamp,
                    0.15,
                    "Filesystem creation time " + Utc(item.Evidence.CreatedUtc.Value) +
                    " is weak ordering evidence only."));
            }
            if (item.DuplicateRevisionIds.Count > 0)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.ExactDuplicateContent,
                    1.0,
                    "The exact byte hash is shared by another occurrence. This proves content identity, not " +
                    "chronology or parentage.",
                    item.DuplicateRevisionIds));
            }
            if (item.ConflictingRevisionIds.Count > 0)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.ConflictingTimestamp,
                    0.95,
                    "Weak timestamps conflict with explicit revision-token order; the explicit token constraint wins.",
                    item.ConflictingRevisionIds.OrderBy(id => id, StringComparer.Ordinal)));
            }
            if (item.StrongToken == null && !item.EffectiveTimestamp.HasValue)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.DeterministicTieBreak,
                    0,
                    "No reliable order evidence exists; filename and revision ID provide only a reproducible tie-break."));
            }
            return reasons;
        }

        private static double OrderConfidence(
            WorkItem item,
            List<WorkItem> all,
            HashSet<string> timestampConflicts)
        {
            double confidence = item.StrongToken != null
                ? item.StrongToken.Confidence
                : item.EffectiveTimestamp.HasValue ? 0.25 : 0;

            if (item.StrongToken != null)
            {
                int sameNumber = all.Count(candidate => candidate.StrongToken != null &&
                    candidate.StrongToken.Number.Equals(item.StrongToken.Number));
                if (sameNumber > 1) confidence = Math.Min(confidence, 0.35);
            }
            if (item.DuplicateRevisionIds.Count > 0) confidence = Math.Min(confidence, 0.25);
            if (timestampConflicts.Contains(item.Evidence.RevisionId)) confidence = Math.Min(confidence, 0.70);
            return confidence;
        }

        private static ParentProposal InferParent(WorkItem child, int index, List<WorkItem> ordered)
        {
            var reasons = new List<LineageReason>();
            if (index == 0)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.FirstRevisionHasNoParent,
                    1,
                    "The first proposed occurrence has no inferred parent."));
                return new ParentProposal(null, 0, reasons);
            }
            if (child.DuplicateRevisionIds.Count > 0)
            {
                reasons.Add(new LineageReason(
                    LineageReasonCode.ExactDuplicateContent,
                    1,
                    "An exact duplicate is a separate occurrence of the same state; no parent is fabricated.",
                    child.DuplicateRevisionIds));
                return new ParentProposal(null, 0, reasons);
            }

            if (child.StrongToken != null)
            {
                List<WorkItem> lower = ordered.Take(index)
                    .Where(candidate => candidate.StrongToken != null &&
                        candidate.StrongToken.Number.CompareTo(child.StrongToken.Number) < 0)
                    .ToList();
                if (lower.Count > 0)
                {
                    RevisionNumber nearestNumber = lower
                        .Select(candidate => candidate.StrongToken.Number)
                        .OrderByDescending(number => number)
                        .First();
                    List<WorkItem> nearest = lower
                        .Where(candidate => candidate.StrongToken.Number.Equals(nearestNumber))
                        .ToList();
                    if (nearest.Count == 1)
                    {
                        WorkItem parent = nearest[0];
                        reasons.Add(new LineageReason(
                            LineageReasonCode.FilenameParentProposal,
                            0.65,
                            "The unique nearest lower explicit revision token proposes a parent. Filename " +
                            "sequence does not prove ancestry and requires confirmation.",
                            new[] { parent.Evidence.RevisionId }));
                        return new ParentProposal(parent.Evidence.RevisionId, 0.65, reasons);
                    }

                    reasons.Add(new LineageReason(
                        LineageReasonCode.AmbiguousParent,
                        0,
                        "Several occurrences share the nearest lower revision token; no parent was selected.",
                        nearest.Select(candidate => candidate.Evidence.RevisionId)
                            .OrderBy(id => id, StringComparer.Ordinal)));
                    return new ParentProposal(null, 0, reasons);
                }
            }
            else
            {
                WorkItem previous = ordered[index - 1];
                if (previous.EffectiveTimestamp.HasValue && child.EffectiveTimestamp.HasValue &&
                    previous.EffectiveTimestamp.Value <= child.EffectiveTimestamp.Value &&
                    !string.Equals(
                        previous.Evidence.ByteContentId,
                        child.Evidence.ByteContentId,
                        StringComparison.Ordinal))
                {
                    reasons.Add(new LineageReason(
                        LineageReasonCode.WeakTimestampParentProposal,
                        0.20,
                        "The immediately preceding weak timestamp proposes a parent. This is adjacency, not " +
                        "proof of derivation, and requires confirmation.",
                        new[] { previous.Evidence.RevisionId }));
                    return new ParentProposal(previous.Evidence.RevisionId, 0.20, reasons);
                }
            }

            reasons.Add(new LineageReason(
                LineageReasonCode.NoParentEvidence,
                0,
                "No conservative parent proposal could be made from the available evidence."));
            return new ParentProposal(null, 0, reasons);
        }

        private static List<LineageReason> BuildGlobalReasons(LineageExplicitSelections selections)
        {
            var reasons = new List<LineageReason>();
            reasons.Add(selections.MasterRevisionId == null
                ? new LineageReason(
                    LineageReasonCode.SelectionRequired,
                    0,
                    "Master is not inferred automatically; an explicit user selection is required.")
                : new LineageReason(
                    LineageReasonCode.ExplicitMasterSelection,
                    1,
                    "Master was explicitly selected.",
                    new[] { selections.MasterRevisionId }));
            reasons.Add(selections.TrustedBaselineRevisionId == null
                ? new LineageReason(
                    LineageReasonCode.SelectionRequired,
                    0,
                    "Trusted baseline is not inferred automatically; an explicit user selection is required.")
                : new LineageReason(
                    LineageReasonCode.ExplicitBaselineSelection,
                    1,
                    "Trusted baseline was explicitly selected.",
                    new[] { selections.TrustedBaselineRevisionId }));
            return reasons;
        }

        private static string ComputeProposalGenerationId(
            IEnumerable<LineageRevisionEvidence> evidence,
            LineageExplicitSelections selections)
        {
            var components = new List<string>();
            foreach (LineageRevisionEvidence item in evidence.OrderBy(
                         value => value.RevisionId,
                         StringComparer.Ordinal))
            {
                components.Add(item.RevisionId);
                components.Add(item.FileName);
                components.Add(item.ByteContentId);
                components.Add(DateComponent(item.CreatedUtc));
                components.Add(DateComponent(item.ModifiedUtc));
                components.Add(item.Description);
            }
            components.Add(selections.MasterRevisionId);
            components.Add(selections.TrustedBaselineRevisionId);
            return StableId.Compute("lineage-inference-proposal-generation-v1", components.ToArray());
        }

        private static string DateComponent(DateTimeOffset? value) =>
            value.HasValue ? value.Value.UtcDateTime.Ticks.ToString(CultureInfo.InvariantCulture) : null;

        private static string Utc(DateTimeOffset value) =>
            value.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ", CultureInfo.InvariantCulture);

        private static void ValidateSelection(string revisionId, HashSet<string> ids, string role)
        {
            if (revisionId != null && !ids.Contains(revisionId))
                throw new ArgumentException("The explicit " + role + " selection is not an imported revision: " + revisionId);
        }

        private sealed class WorkItem
        {
            internal WorkItem(LineageRevisionEvidence evidence, RevisionLabelInferenceResult labelInference)
            {
                Evidence = evidence;
                LabelInference = labelInference;
                DuplicateRevisionIds = Array.Empty<string>();
                ConflictingRevisionIds = new HashSet<string>(StringComparer.Ordinal);
            }

            internal LineageRevisionEvidence Evidence { get; }
            internal RevisionLabelInferenceResult LabelInference { get; }
            internal RevisionTokenCandidate StrongToken => LabelInference.StrongCandidate;
            internal DateTimeOffset? EffectiveTimestamp => Evidence.ModifiedUtc ?? Evidence.CreatedUtc;
            internal IReadOnlyList<string> DuplicateRevisionIds { get; set; }
            internal HashSet<string> ConflictingRevisionIds { get; }
        }

        private sealed class WorkItemTieBreaker : IComparer<WorkItem>
        {
            internal static readonly WorkItemTieBreaker Instance = new WorkItemTieBreaker();

            public int Compare(WorkItem left, WorkItem right)
            {
                if (ReferenceEquals(left, right)) return 0;
                if (left == null) return -1;
                if (right == null) return 1;
                if (left.EffectiveTimestamp.HasValue && right.EffectiveTimestamp.HasValue)
                {
                    int timestamp = left.EffectiveTimestamp.Value.CompareTo(right.EffectiveTimestamp.Value);
                    if (timestamp != 0) return timestamp;
                }
                else if (left.EffectiveTimestamp.HasValue != right.EffectiveTimestamp.HasValue)
                {
                    return left.EffectiveTimestamp.HasValue ? -1 : 1;
                }

                int insensitive = StringComparer.OrdinalIgnoreCase.Compare(left.Evidence.FileName, right.Evidence.FileName);
                if (insensitive != 0) return insensitive;
                int fileName = StringComparer.Ordinal.Compare(left.Evidence.FileName, right.Evidence.FileName);
                if (fileName != 0) return fileName;
                return StringComparer.Ordinal.Compare(left.Evidence.RevisionId, right.Evidence.RevisionId);
            }
        }

        private sealed class ParentProposal
        {
            internal ParentProposal(string parentRevisionId, double confidence, IEnumerable<LineageReason> reasons)
            {
                ParentRevisionId = parentRevisionId;
                Confidence = confidence;
                Reasons = reasons;
            }

            internal string ParentRevisionId { get; }
            internal double Confidence { get; }
            internal IEnumerable<LineageReason> Reasons { get; }
        }
    }
}
