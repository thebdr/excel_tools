using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>
    /// Immutable optimistic-concurrency request. All affected subject IDs are captured before application;
    /// group application additionally verifies that this frozen set still exactly matches the named group.
    /// </summary>
    public sealed class ReviewDecisionRequest
    {
        private ReviewDecisionRequest(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            DecisionSelectionScope selectionScope,
            string subjectId,
            string groupId,
            IEnumerable<string> affectedSubjectIds,
            ReviewDecisionStatus newDecision,
            string selectedObservationId,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string runId,
            string addInBuild,
            string note)
        {
            ProjectId = Require(projectId, nameof(projectId));
            BaseGenerationId = Require(baseGenerationId, nameof(baseGenerationId));
            MappingGenerationId = Require(mappingGenerationId, nameof(mappingGenerationId));
            ExpectedDecisionGenerationId = Require(
                expectedDecisionGenerationId,
                nameof(expectedDecisionGenerationId));
            if (!Enum.IsDefined(typeof(DecisionSelectionScope), selectionScope))
                throw new ArgumentOutOfRangeException(nameof(selectionScope));
            if (!Enum.IsDefined(typeof(ReviewDecisionStatus), newDecision))
                throw new ArgumentOutOfRangeException(nameof(newDecision));
            if (createdUtc == default(DateTimeOffset))
                throw new ArgumentException("A decision request requires a creation timestamp.", nameof(createdUtc));
            if (createdUtc.Offset != TimeSpan.Zero)
                throw new ArgumentException("CreatedUtc must have a zero UTC offset.", nameof(createdUtc));

            SelectionScope = selectionScope;
            SubjectId = subjectId ?? "";
            GroupId = groupId ?? "";
            AffectedSubjectIds = new ReadOnlyCollection<string>(
                FreezeAffectedSubjects(affectedSubjectIds));
            ValidateScope();

            SelectedObservationId = selectedObservationId ?? "";
            ValidateOptionalIdentity(SelectedObservationId, nameof(selectedObservationId));
            ValidateSelectionSemantics(newDecision, selectionScope, SelectedObservationId);
            ValidateOptionalIdentity(runId, nameof(runId));

            NewDecision = newDecision;
            Reviewer = reviewer ?? ReviewerSnapshot.Empty;
            CreatedUtc = createdUtc;
            RunId = runId ?? "";
            AddInBuild = addInBuild ?? "";
            Note = note ?? "";
            RequestId = ComputeRequestId();
        }

        public string RequestId { get; }
        public string ProjectId { get; }
        public string BaseGenerationId { get; }
        public string MappingGenerationId { get; }
        public string ExpectedDecisionGenerationId { get; }
        public DecisionSelectionScope SelectionScope { get; }
        public string SubjectId { get; }
        public string GroupId { get; }
        public IReadOnlyList<string> AffectedSubjectIds { get; }
        public ReviewDecisionStatus NewDecision { get; }
        public string SelectedObservationId { get; }
        public ReviewerSnapshot Reviewer { get; }
        public DateTimeOffset CreatedUtc { get; }
        public string RunId { get; }
        public string AddInBuild { get; }
        public string Note { get; }

        public static ReviewDecisionRequest ForSubject(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            string subjectId,
            ReviewDecisionStatus newDecision,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string selectedObservationId = "",
            string runId = "",
            string addInBuild = "",
            string note = "") =>
            new ReviewDecisionRequest(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                expectedDecisionGenerationId,
                DecisionSelectionScope.SingleSubject,
                Require(subjectId, nameof(subjectId)),
                "",
                new[] { subjectId },
                newDecision,
                selectedObservationId,
                reviewer,
                createdUtc,
                runId,
                addInBuild,
                note);

        public static ReviewDecisionRequest ForSubjects(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            IEnumerable<string> subjectIds,
            ReviewDecisionStatus newDecision,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string selectedObservationId = "",
            string runId = "",
            string addInBuild = "",
            string note = "") =>
            new ReviewDecisionRequest(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                expectedDecisionGenerationId,
                DecisionSelectionScope.MultipleSubjects,
                "",
                "",
                subjectIds,
                newDecision,
                selectedObservationId,
                reviewer,
                createdUtc,
                runId,
                addInBuild,
                note);

        public static ReviewDecisionRequest ForGroup(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            string groupId,
            IEnumerable<string> affectedSubjectIds,
            ReviewDecisionStatus newDecision,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string selectedObservationId = "",
            string runId = "",
            string addInBuild = "",
            string note = "") =>
            new ReviewDecisionRequest(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                expectedDecisionGenerationId,
                DecisionSelectionScope.Group,
                "",
                Require(groupId, nameof(groupId)),
                affectedSubjectIds,
                newDecision,
                selectedObservationId,
                reviewer,
                createdUtc,
                runId,
                addInBuild,
                note);

        public static ReviewDecisionRequest ClearSubject(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            string subjectId,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string runId = "",
            string addInBuild = "",
            string note = "") =>
            ForSubject(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                expectedDecisionGenerationId,
                subjectId,
                ReviewDecisionStatus.Unreviewed,
                reviewer,
                createdUtc,
                runId: runId,
                addInBuild: addInBuild,
                note: note);

        public static ReviewDecisionRequest ClearSubjects(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            IEnumerable<string> subjectIds,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string runId = "",
            string addInBuild = "",
            string note = "") =>
            ForSubjects(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                expectedDecisionGenerationId,
                subjectIds,
                ReviewDecisionStatus.Unreviewed,
                reviewer,
                createdUtc,
                runId: runId,
                addInBuild: addInBuild,
                note: note);

        public static ReviewDecisionRequest ClearGroup(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string expectedDecisionGenerationId,
            string groupId,
            IEnumerable<string> affectedSubjectIds,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string runId = "",
            string addInBuild = "",
            string note = "") =>
            ForGroup(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                expectedDecisionGenerationId,
                groupId,
                affectedSubjectIds,
                ReviewDecisionStatus.Unreviewed,
                reviewer,
                createdUtc,
                runId: runId,
                addInBuild: addInBuild,
                note: note);

        private string ComputeRequestId()
        {
            var components = new List<string>
            {
                ProjectId,
                BaseGenerationId,
                MappingGenerationId,
                ExpectedDecisionGenerationId,
                ((int)SelectionScope).ToString(CultureInfo.InvariantCulture),
                SubjectId,
                GroupId,
                ((int)NewDecision).ToString(CultureInfo.InvariantCulture),
                SelectedObservationId,
                RunId,
                AddInBuild,
                CreatedUtc.ToString("O", CultureInfo.InvariantCulture),
                Reviewer.ConfirmedDisplayName,
                Reviewer.OfficeDisplayName,
                Note,
            };
            components.AddRange(AffectedSubjectIds.Select(item => "affected:" + item));
            components.AddRange(Reviewer.IdentitySourceLabels.Select(item => "reviewer-source:" + item));
            return StableId.Compute("review-decision-request-v1", components.ToArray());
        }

        private void ValidateScope()
        {
            switch (SelectionScope)
            {
                case DecisionSelectionScope.SingleSubject:
                    Require(SubjectId, nameof(SubjectId));
                    if (GroupId.Length != 0 || AffectedSubjectIds.Count != 1 ||
                        !string.Equals(AffectedSubjectIds[0], SubjectId, StringComparison.Ordinal))
                        throw new ArgumentException("A single-subject request must freeze exactly its SubjectId.");
                    break;
                case DecisionSelectionScope.MultipleSubjects:
                    if (SubjectId.Length != 0 || GroupId.Length != 0 || AffectedSubjectIds.Count < 2)
                        throw new ArgumentException("A multi-subject request requires at least two affected subjects.");
                    break;
                case DecisionSelectionScope.Group:
                    Require(GroupId, nameof(GroupId));
                    if (SubjectId.Length != 0 || AffectedSubjectIds.Count == 0)
                        throw new ArgumentException("A group request must freeze at least one affected subject.");
                    break;
            }
        }

        private static List<string> FreezeAffectedSubjects(IEnumerable<string> subjectIds)
        {
            if (subjectIds == null) throw new ArgumentNullException(nameof(subjectIds));
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string subjectId in subjectIds)
            {
                string required = Require(subjectId, nameof(subjectIds));
                if (!seen.Add(required))
                    throw new ArgumentException(
                        "A decision request cannot repeat an affected subject ID.",
                        nameof(subjectIds));
                result.Add(required);
            }
            result.Sort(StringComparer.Ordinal);
            return result;
        }

        private static void ValidateSelectionSemantics(
            ReviewDecisionStatus decision,
            DecisionSelectionScope scope,
            string selectedObservationId)
        {
            bool selectsObservation = decision == ReviewDecisionStatus.UseMaster ||
                                      decision == ReviewDecisionStatus.UseCandidate;
            if (selectsObservation && string.IsNullOrWhiteSpace(selectedObservationId))
                throw new ArgumentException(
                    "Use-master and use-candidate requests require a selected observation ID.",
                    nameof(selectedObservationId));
            if (!selectsObservation && selectedObservationId.Length != 0)
                throw new ArgumentException(
                    "Only use-master and use-candidate requests may select an observation.",
                    nameof(selectedObservationId));
            if (selectsObservation && scope != DecisionSelectionScope.SingleSubject)
                throw new ArgumentException(
                    "A batch cannot use one scalar observation ID for several subjects; submit explicit single-subject selections.",
                    nameof(scope));
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable identity is required.", name);
            return value;
        }

        private static void ValidateOptionalIdentity(string value, string name)
        {
            if (value != null && value.Length != 0 && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An optional identity must be empty or non-blank.", name);
        }
    }
}
