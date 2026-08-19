using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>
    /// Immutable append-only reviewer decision. It references immutable evidence and validates its selection scope
    /// before it can be persisted. The constructor parameter names intentionally match the public properties so
    /// constructor-aware JSON serializers can round-trip the contract without public setters.
    /// </summary>
    public sealed class DecisionEvent
    {
        public DecisionEvent(
            string eventId,
            string projectId,
            string generationId,
            string mappingGenerationId,
            ReviewDecisionStatus priorDecision,
            ReviewDecisionStatus newDecision,
            DateTimeOffset createdUtc,
            DecisionSelectionScope selectionScope,
            string runId = "",
            string subjectId = "",
            string groupId = "",
            IReadOnlyList<string> affectedSubjectIds = null,
            string selectedObservationId = "",
            IReadOnlyList<string> supportingRevisionIds = null,
            IReadOnlyList<string> supportingContentHashes = null,
            ReviewerSnapshot reviewerSnapshot = null,
            string addInBuild = "",
            string note = "",
            string supersededEventId = "",
            IReadOnlyList<DecisionLocationContext> locations = null,
            string requestId = "")
        {
            EventId = Require(eventId, nameof(eventId));
            ProjectId = Require(projectId, nameof(projectId));
            GenerationId = Require(generationId, nameof(generationId));
            MappingGenerationId = Require(mappingGenerationId, nameof(mappingGenerationId));
            if (!Enum.IsDefined(typeof(ReviewDecisionStatus), priorDecision))
                throw new ArgumentOutOfRangeException(nameof(priorDecision));
            if (!Enum.IsDefined(typeof(ReviewDecisionStatus), newDecision))
                throw new ArgumentOutOfRangeException(nameof(newDecision));
            if (!Enum.IsDefined(typeof(DecisionSelectionScope), selectionScope))
                throw new ArgumentOutOfRangeException(nameof(selectionScope));
            if (createdUtc == default(DateTimeOffset))
                throw new ArgumentException("A decision event requires a creation timestamp.", nameof(createdUtc));
            if (createdUtc.Offset != TimeSpan.Zero)
                throw new ArgumentException("CreatedUtc must have a zero UTC offset.", nameof(createdUtc));

            var affected = CopyIdentities(affectedSubjectIds, nameof(affectedSubjectIds));
            SubjectId = subjectId ?? "";
            GroupId = groupId ?? "";
            ValidateSelectionScope(selectionScope, SubjectId, GroupId, affected);

            SelectedObservationId = selectedObservationId ?? "";
            bool selectsObservation = newDecision == ReviewDecisionStatus.UseMaster ||
                                      newDecision == ReviewDecisionStatus.UseCandidate;
            if (selectsObservation && string.IsNullOrWhiteSpace(SelectedObservationId))
            {
                throw new ArgumentException(
                    "A decision that selects a revision requires the selected observation ID.",
                    nameof(selectedObservationId));
            }
            if (selectsObservation && selectionScope != DecisionSelectionScope.SingleSubject)
            {
                throw new ArgumentException(
                    "A decision that selects one observation must use single-subject scope.",
                    nameof(selectionScope));
            }
            if (!selectsObservation && SelectedObservationId.Length != 0)
            {
                throw new ArgumentException(
                    "Only a use-master or use-candidate decision may select an observation.",
                    nameof(selectedObservationId));
            }
            ValidateOptionalIdentity(SelectedObservationId, nameof(selectedObservationId));
            ValidateOptionalIdentity(runId, nameof(runId));
            ValidateOptionalIdentity(supersededEventId, nameof(supersededEventId));
            ValidateOptionalIdentity(requestId, nameof(requestId));

            RunId = runId ?? "";
            RequestId = requestId ?? "";
            PriorDecision = priorDecision;
            NewDecision = newDecision;
            CreatedUtc = createdUtc;
            SelectionScope = selectionScope;
            AffectedSubjectIds = new ReadOnlyCollection<string>(affected);
            SupportingRevisionIds = new ReadOnlyCollection<string>(
                CopyIdentities(supportingRevisionIds, nameof(supportingRevisionIds)));
            SupportingContentHashes = new ReadOnlyCollection<string>(
                CopyIdentities(supportingContentHashes, nameof(supportingContentHashes)));
            ReviewerSnapshot = reviewerSnapshot ?? ReviewerSnapshot.Empty;
            AddInBuild = addInBuild ?? "";
            Note = note ?? "";
            SupersededEventId = supersededEventId ?? "";
            Locations = new ReadOnlyCollection<DecisionLocationContext>(CopyLocations(locations));
        }

        public string EventId { get; }
        /// <summary>
        /// Deterministic identity of the immutable request that produced this event. Empty is retained only for
        /// backward-compatible construction of legacy in-memory events; durable replay records require it.
        /// </summary>
        public string RequestId { get; }
        public string ProjectId { get; }
        public string RunId { get; }
        /// <summary>Immutable evidence generation reviewed by this event.</summary>
        public string EvidenceGenerationId => GenerationId;
        /// <summary>Backward-compatible name for <see cref="EvidenceGenerationId"/>.</summary>
        public string GenerationId { get; }
        public string MappingGenerationId { get; }
        public string SubjectId { get; }
        public string GroupId { get; }
        public IReadOnlyList<string> AffectedSubjectIds { get; }
        public ReviewDecisionStatus PriorDecision { get; }
        public ReviewDecisionStatus NewDecision { get; }
        public string SelectedObservationId { get; }
        public IReadOnlyList<string> SupportingRevisionIds { get; }
        public IReadOnlyList<string> SupportingContentHashes { get; }
        public ReviewerSnapshot ReviewerSnapshot { get; }
        public DateTimeOffset CreatedUtc { get; }
        public string AddInBuild { get; }
        public string Note { get; }
        public DecisionSelectionScope SelectionScope { get; }
        public string SupersededEventId { get; }
        public IReadOnlyList<DecisionLocationContext> Locations { get; }

        private static void ValidateSelectionScope(
            DecisionSelectionScope scope,
            string subjectId,
            string groupId,
            List<string> affected)
        {
            switch (scope)
            {
                case DecisionSelectionScope.SingleSubject:
                    Require(subjectId, nameof(subjectId));
                    if (!string.IsNullOrEmpty(groupId))
                        throw new ArgumentException("A single-subject decision cannot also target a group.", nameof(groupId));
                    if (affected.Count == 0)
                        affected.Add(subjectId);
                    if (affected.Count != 1 || !string.Equals(affected[0], subjectId, StringComparison.Ordinal))
                        throw new ArgumentException(
                            "A single-subject decision must affect exactly its SubjectId.",
                            nameof(affected));
                    break;

                case DecisionSelectionScope.MultipleSubjects:
                    if (!string.IsNullOrEmpty(subjectId) || !string.IsNullOrEmpty(groupId))
                        throw new ArgumentException("A multi-subject decision uses only AffectedSubjectIds.");
                    if (affected.Count < 2)
                        throw new ArgumentException("A multi-subject decision requires at least two subjects.", nameof(affected));
                    break;

                case DecisionSelectionScope.Group:
                    Require(groupId, nameof(groupId));
                    if (!string.IsNullOrEmpty(subjectId))
                        throw new ArgumentException("A group decision cannot also target one SubjectId.", nameof(subjectId));
                    if (affected.Count == 0)
                        throw new ArgumentException("A group decision must freeze at least one affected subject.", nameof(affected));
                    break;
            }
        }

        private static List<string> CopyIdentities(IEnumerable<string> values, string parameterName)
        {
            var result = new List<string>();
            if (values == null) return result;

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string value in values)
            {
                Require(value, parameterName);
                if (!seen.Add(value))
                    throw new ArgumentException("A durable identity collection cannot contain duplicates.", parameterName);
                result.Add(value);
            }
            return result;
        }

        private static List<DecisionLocationContext> CopyLocations(
            IEnumerable<DecisionLocationContext> locations)
        {
            var result = new List<DecisionLocationContext>();
            if (locations == null) return result;
            foreach (DecisionLocationContext location in locations)
            {
                if (location == null)
                    throw new ArgumentException("Decision locations cannot contain null entries.", nameof(locations));
                result.Add(location);
            }
            return result;
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable identity component is required.", name);
            return value;
        }

        private static void ValidateOptionalIdentity(string value, string name)
        {
            if (value != null && value.Length != 0 && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An optional identity must be empty or non-blank.", name);
        }
    }

    public sealed class ReviewerSnapshot
    {
        private static readonly ReviewerSnapshot EmptyInstance = new ReviewerSnapshot();

        public ReviewerSnapshot(
            string confirmedDisplayName = "",
            string officeDisplayName = "",
            IReadOnlyList<string> identitySourceLabels = null)
        {
            ConfirmedDisplayName = confirmedDisplayName ?? "";
            OfficeDisplayName = officeDisplayName ?? "";
            IdentitySourceLabels = new ReadOnlyCollection<string>(
                CopyLabels(identitySourceLabels));
        }

        public static ReviewerSnapshot Empty => EmptyInstance;

        public string ConfirmedDisplayName { get; }
        public string OfficeDisplayName { get; }
        public IReadOnlyList<string> IdentitySourceLabels { get; }

        private static List<string> CopyLabels(IEnumerable<string> labels)
        {
            var result = new List<string>();
            if (labels == null) return result;
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string label in labels)
            {
                if (string.IsNullOrWhiteSpace(label))
                    throw new ArgumentException("Reviewer identity-source labels cannot be blank.", nameof(labels));
                if (!seen.Add(label))
                    throw new ArgumentException("Reviewer identity-source labels cannot be duplicated.", nameof(labels));
                result.Add(label);
            }
            return result;
        }
    }

    public sealed class DecisionLocationContext
    {
        public DecisionLocationContext(
            string role,
            string revisionId,
            string logicalSheetId = "",
            string sheetName = "",
            string address = "",
            string sourceLabel = "",
            int physicalRow = 0,
            int physicalColumn = 0)
        {
            if (string.IsNullOrWhiteSpace(role))
                throw new ArgumentException("A decision-location role is required.", nameof(role));
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A decision-location revision ID is required.", nameof(revisionId));
            if (physicalRow < 0)
                throw new ArgumentOutOfRangeException(nameof(physicalRow));
            if (physicalColumn < 0)
                throw new ArgumentOutOfRangeException(nameof(physicalColumn));
            Role = role;
            RevisionId = revisionId;
            LogicalSheetId = logicalSheetId ?? "";
            SheetName = sheetName ?? "";
            Address = address ?? "";
            SourceLabel = sourceLabel ?? "";
            PhysicalRow = physicalRow;
            PhysicalColumn = physicalColumn;
        }

        public string Role { get; }
        public string RevisionId { get; }
        public string LogicalSheetId { get; }
        public string SheetName { get; }
        public string Address { get; }
        public string SourceLabel { get; }
        public int PhysicalRow { get; }
        public int PhysicalColumn { get; }
    }
}
