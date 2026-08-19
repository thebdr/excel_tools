using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;

namespace ExcelPowerTools.Core.Comparison.Projection
{
    public enum LogicalAxisKind
    {
        Row,
        Column,
    }

    public enum LogicalAxisIdentitySource
    {
        Root,
        Inherited,
        Inserted,
    }

    public enum LogicalAxisStructuralEventKind
    {
        Insertion,
        Deletion,
        Reorder,
    }

    public enum LogicalAxisAmbiguityKind
    {
        AlignmentBlock,
        MissingParentIdentity,
        UncoveredPhysicalIndex,
    }

    /// <summary>One logical identity assigned to one zero-based physical axis index in a revision.</summary>
    public sealed class LogicalAxisIdentityAssignment
    {
        internal LogicalAxisIdentityAssignment(
            string assignmentId,
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
            AssignmentId = Require(assignmentId, nameof(assignmentId));
            LogicalIdentityId = Require(logicalIdentityId, nameof(logicalIdentityId));
            RevisionId = Require(revisionId, nameof(revisionId));
            if (physicalIndex < 0) throw new ArgumentOutOfRangeException(nameof(physicalIndex));
            if (originPhysicalIndex < 0) throw new ArgumentOutOfRangeException(nameof(originPhysicalIndex));
            if (parentPhysicalIndex < 0) throw new ArgumentOutOfRangeException(nameof(parentPhysicalIndex));

            PhysicalIndex = physicalIndex;
            Source = source;
            OriginRevisionId = Require(originRevisionId, nameof(originRevisionId));
            OriginPhysicalIndex = originPhysicalIndex;
            ParentPhysicalIndex = parentPhysicalIndex;
            Confidence = confidence;
            Strategy = strategy;
            Reasons = Freeze(reasons);
        }

        public string AssignmentId { get; }
        public string LogicalIdentityId { get; }
        public string RevisionId { get; }
        public int PhysicalIndex { get; }
        public LogicalAxisIdentitySource Source { get; }
        public string OriginRevisionId { get; }
        public int OriginPhysicalIndex { get; }
        public int? ParentPhysicalIndex { get; }
        public AlignmentConfidence Confidence { get; }
        public AxisAlignmentStrategy Strategy { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A stable identity component is required.", name);
            return value;
        }

        private static IReadOnlyList<AlignmentReason> Freeze(IEnumerable<AlignmentReason> reasons) =>
            new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
    }

    /// <summary>An explicit structural effect of one parent-to-child alignment edge.</summary>
    public sealed class LogicalAxisStructuralEvent
    {
        internal LogicalAxisStructuralEvent(
            string eventId,
            LogicalAxisStructuralEventKind kind,
            string parentRevisionId,
            string childRevisionId,
            string logicalIdentityId,
            int? beforeIndex,
            int? afterIndex,
            AlignmentConfidence confidence,
            AxisAlignmentStrategy strategy,
            IEnumerable<AlignmentReason> reasons)
        {
            EventId = Require(eventId, nameof(eventId));
            Kind = kind;
            ParentRevisionId = Require(parentRevisionId, nameof(parentRevisionId));
            ChildRevisionId = Require(childRevisionId, nameof(childRevisionId));
            LogicalIdentityId = logicalIdentityId ?? "";
            if (beforeIndex < 0) throw new ArgumentOutOfRangeException(nameof(beforeIndex));
            if (afterIndex < 0) throw new ArgumentOutOfRangeException(nameof(afterIndex));
            BeforeIndex = beforeIndex;
            AfterIndex = afterIndex;
            Confidence = confidence;
            Strategy = strategy;
            Reasons = Freeze(reasons);
        }

        public string EventId { get; }
        public LogicalAxisStructuralEventKind Kind { get; }
        public string ParentRevisionId { get; }
        public string ChildRevisionId { get; }
        /// <summary>Empty only when the parent physical index already had unresolved identity.</summary>
        public string LogicalIdentityId { get; }
        public int? BeforeIndex { get; }
        public int? AfterIndex { get; }
        public AlignmentConfidence Confidence { get; }
        public AxisAlignmentStrategy Strategy { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A value is required.", name);
            return value;
        }

        private static IReadOnlyList<AlignmentReason> Freeze(IEnumerable<AlignmentReason> reasons) =>
            new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
    }

    /// <summary>
    /// An unresolved physical block. After indices in this block deliberately have no logical assignment.
    /// Known parent candidates are retained as evidence without pairing them to any child index.
    /// </summary>
    public sealed class LogicalAxisIdentityAmbiguity
    {
        internal LogicalAxisIdentityAmbiguity(
            string ambiguityId,
            LogicalAxisAmbiguityKind kind,
            IEnumerable<int> beforeIndices,
            IEnumerable<int> afterIndices,
            IDictionary<int, string> knownParentIdentities,
            string evidenceToken,
            AlignmentConfidence confidence,
            AxisAlignmentStrategy strategy,
            IEnumerable<AlignmentReason> reasons)
        {
            if (string.IsNullOrWhiteSpace(ambiguityId))
                throw new ArgumentException("An ambiguity ID is required.", nameof(ambiguityId));
            AmbiguityId = ambiguityId;
            Kind = kind;
            BeforeIndices = FreezeIndices(beforeIndices);
            AfterIndices = FreezeIndices(afterIndices);
            if (BeforeIndices.Count == 0 && AfterIndices.Count == 0)
                throw new ArgumentException("An ambiguity must identify at least one physical index.");
            KnownParentIdentities = new ReadOnlyDictionary<int, string>(
                new Dictionary<int, string>(knownParentIdentities ?? new Dictionary<int, string>()));
            EvidenceToken = evidenceToken ?? "";
            Confidence = confidence;
            Strategy = strategy;
            Reasons = new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
        }

        public string AmbiguityId { get; }
        public LogicalAxisAmbiguityKind Kind { get; }
        public IReadOnlyList<int> BeforeIndices { get; }
        public IReadOnlyList<int> AfterIndices { get; }
        public IReadOnlyDictionary<int, string> KnownParentIdentities { get; }
        public string EvidenceToken { get; }
        public AlignmentConfidence Confidence { get; }
        public AxisAlignmentStrategy Strategy { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }

        private static IReadOnlyList<int> FreezeIndices(IEnumerable<int> values)
        {
            var copy = new List<int>(values ?? Array.Empty<int>());
            copy.Sort();
            return new ReadOnlyCollection<int>(copy);
        }
    }

    /// <summary>
    /// Immutable logical-axis state for one revision. Missing physical-index assignments are intentional and
    /// are explained by <see cref="Ambiguities"/>; consumers must not fall back to physical identity.
    /// </summary>
    public sealed class LogicalAxisIdentityProjection
    {
        internal LogicalAxisIdentityProjection(
            string projectionId,
            string projectId,
            string logicalSheetId,
            string revisionId,
            string parentRevisionId,
            string parentProjectionId,
            LogicalAxisKind axis,
            int physicalCount,
            AlignmentConfidence confidence,
            IDictionary<int, LogicalAxisIdentityAssignment> assignments,
            IEnumerable<LogicalAxisStructuralEvent> structuralEvents,
            IEnumerable<LogicalAxisIdentityAmbiguity> ambiguities)
        {
            ProjectionId = Require(projectionId, nameof(projectionId));
            ProjectId = Require(projectId, nameof(projectId));
            LogicalSheetId = Require(logicalSheetId, nameof(logicalSheetId));
            RevisionId = Require(revisionId, nameof(revisionId));
            ParentRevisionId = parentRevisionId ?? "";
            ParentProjectionId = parentProjectionId ?? "";
            Axis = axis;
            if (physicalCount < 0) throw new ArgumentOutOfRangeException(nameof(physicalCount));
            PhysicalCount = physicalCount;
            Confidence = confidence;

            var assignmentCopy = new Dictionary<int, LogicalAxisIdentityAssignment>();
            var identityCopy = new Dictionary<int, string>();
            var seenIdentities = new HashSet<string>(StringComparer.Ordinal);
            foreach (KeyValuePair<int, LogicalAxisIdentityAssignment> pair in
                     (assignments ?? new Dictionary<int, LogicalAxisIdentityAssignment>()).OrderBy(item => item.Key))
            {
                if (pair.Key < 0 || pair.Key >= physicalCount)
                    throw new ArgumentOutOfRangeException(nameof(assignments));
                if (pair.Value == null || pair.Value.PhysicalIndex != pair.Key)
                    throw new ArgumentException("Assignment keys must match their physical indices.", nameof(assignments));
                if (!string.Equals(pair.Value.RevisionId, revisionId, StringComparison.Ordinal))
                    throw new ArgumentException("Assignments must belong to the projection revision.", nameof(assignments));
                if (!seenIdentities.Add(pair.Value.LogicalIdentityId))
                    throw new ArgumentException("A logical axis identity cannot be assigned twice.", nameof(assignments));
                assignmentCopy.Add(pair.Key, pair.Value);
                identityCopy.Add(pair.Key, pair.Value.LogicalIdentityId);
            }

            AssignmentsByPhysicalIndex = new ReadOnlyDictionary<int, LogicalAxisIdentityAssignment>(assignmentCopy);
            IdentityByPhysicalIndex = new ReadOnlyDictionary<int, string>(identityCopy);
            StructuralEvents = new ReadOnlyCollection<LogicalAxisStructuralEvent>(
                new List<LogicalAxisStructuralEvent>(structuralEvents ?? Array.Empty<LogicalAxisStructuralEvent>()));
            Ambiguities = new ReadOnlyCollection<LogicalAxisIdentityAmbiguity>(
                new List<LogicalAxisIdentityAmbiguity>(ambiguities ?? Array.Empty<LogicalAxisIdentityAmbiguity>()));

            var unmapped = new List<int>();
            for (int index = 0; index < physicalCount; index++)
                if (!assignmentCopy.ContainsKey(index)) unmapped.Add(index);
            UnmappedPhysicalIndices = new ReadOnlyCollection<int>(unmapped);
        }

        public string ProjectionId { get; }
        public string ProjectId { get; }
        public string LogicalSheetId { get; }
        public string RevisionId { get; }
        public string ParentRevisionId { get; }
        public string ParentProjectionId { get; }
        public LogicalAxisKind Axis { get; }
        public int PhysicalCount { get; }
        public AlignmentConfidence Confidence { get; }
        public IReadOnlyDictionary<int, LogicalAxisIdentityAssignment> AssignmentsByPhysicalIndex { get; }
        public IReadOnlyDictionary<int, string> IdentityByPhysicalIndex { get; }
        public IReadOnlyList<LogicalAxisStructuralEvent> StructuralEvents { get; }
        public IReadOnlyList<LogicalAxisIdentityAmbiguity> Ambiguities { get; }
        public IReadOnlyList<int> UnmappedPhysicalIndices { get; }
        public bool IsFullyResolved => UnmappedPhysicalIndices.Count == 0 && Ambiguities.Count == 0;

        public bool TryGetLogicalIdentity(int physicalIndex, out string logicalIdentityId) =>
            IdentityByPhysicalIndex.TryGetValue(physicalIndex, out logicalIdentityId);

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A value is required.", name);
            return value;
        }
    }
}
