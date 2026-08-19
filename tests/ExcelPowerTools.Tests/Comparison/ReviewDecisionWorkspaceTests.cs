using System;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class ReviewDecisionWorkspaceTests
    {
        [Fact]
        public void Projection_orders_physical_cells_and_retains_exact_choices_groups_and_state()
        {
            ReviewDecisionGeneration root = CreateGeneration();
            ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                root,
                new[] { "subject-formula" },
                ReviewDecisionStatus.UseCandidate,
                ReviewerSnapshot.Empty,
                Utc(10),
                selectedObservationId: ObservationId(
                    "subject-formula",
                    "r2",
                    ObservationExactType.Formula,
                    'b'));
            ReviewDecisionGeneration decided = ReviewDecisionEngine.Apply(root, request);

            ReviewDecisionWorkspaceProjection projection =
                ReviewDecisionWorkspaceProjection.Build(decided);

            Assert.Equal(decided.GenerationId, projection.GenerationId);
            Assert.Equal(new[] { "subject-formula", "subject-value" },
                projection.Rows.Select(item => item.SubjectId));
            ReviewDecisionWorkspaceRow formula = projection.Rows[0];
            Assert.Equal("Data!B2", formula.Location);
            Assert.Equal(ReviewDecisionStatus.UseCandidate, formula.Status);
            Assert.Equal(
                ObservationId("subject-formula", "r2", ObservationExactType.Formula, 'b'),
                formula.SelectedObservationId);
            Assert.Equal("=A2+3", formula.MasterContent);
            Assert.Equal("=A2+2", formula.LatestCandidateContent);
            Assert.Equal("r2.xlsx", formula.LatestCandidateRevisionLabel);
            Assert.Equal("r3", formula.Master.RevisionId);
            Assert.Equal(new[] { "r1", "r2" }, formula.Candidates.Select(item => item.RevisionId));
            Assert.Equal(new[] { "group-a" }, formula.GroupIds);
            Assert.True(formula.CanSelectObservation);
        }

        [Fact]
        public void Request_builder_supports_homogeneous_multi_selection_for_non_observation_statuses()
        {
            ReviewDecisionGeneration root = CreateGeneration();

            ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                root,
                new[] { "subject-value", "subject-formula" },
                ReviewDecisionStatus.Equivalent,
                new ReviewerSnapshot("Reviewer"),
                Utc(11),
                note: "Reviewed together");
            ReviewDecisionGeneration next = ReviewDecisionEngine.Apply(root, request);

            Assert.Equal(DecisionSelectionScope.MultipleSubjects, request.SelectionScope);
            Assert.Equal(new[] { "subject-formula", "subject-value" }, request.AffectedSubjectIds);
            Assert.All(next.SubjectStates.Values,
                item => Assert.Equal(ReviewDecisionStatus.Equivalent, item.Status));
            Assert.Equal("Reviewed together", Assert.Single(next.Events).Note);
        }

        [Fact]
        public void Request_builder_keeps_observation_selection_single_subject_and_exact()
        {
            ReviewDecisionGeneration root = CreateGeneration();

            ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                root,
                new[] { "subject-formula" },
                ReviewDecisionStatus.UseMaster,
                ReviewerSnapshot.Empty,
                Utc(12),
                selectedObservationId: ObservationId(
                    "subject-formula",
                    "r3",
                    ObservationExactType.Formula,
                    'c'));

            ReviewDecisionGeneration next = ReviewDecisionEngine.Apply(root, request);
            Assert.Equal(ReviewDecisionStatus.UseMaster,
                next.GetSubjectState("subject-formula").Status);
            Assert.Equal(
                ObservationId("subject-formula", "r3", ObservationExactType.Formula, 'c'),
                next.GetSubjectState("subject-formula").SelectedObservationId);

            Assert.Throws<ArgumentException>(() =>
                ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                    root,
                    new[] { "subject-formula", "subject-value" },
                    ReviewDecisionStatus.UseCandidate,
                    ReviewerSnapshot.Empty,
                    Utc(13),
                    selectedObservationId: ObservationId(
                        "subject-formula",
                        "r2",
                        ObservationExactType.Formula,
                        'b')));
        }

        [Fact]
        public void Group_request_freezes_authoritative_membership_and_rejects_unknown_subjects()
        {
            ReviewDecisionGeneration root = CreateGeneration();

            ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForGroup(
                root,
                "group-a",
                ReviewDecisionStatus.NeedsInvestigation,
                ReviewerSnapshot.Empty,
                Utc(14));
            ReviewDecisionGeneration next = ReviewDecisionEngine.Apply(root, request);

            Assert.Equal(DecisionSelectionScope.Group, request.SelectionScope);
            Assert.Equal(new[] { "subject-formula", "subject-value" }, request.AffectedSubjectIds);
            Assert.All(next.SubjectStates.Values,
                item => Assert.Equal(ReviewDecisionStatus.NeedsInvestigation, item.Status));
            Assert.Throws<System.Collections.Generic.KeyNotFoundException>(() =>
                ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                    root,
                    new[] { "not-in-catalog" },
                    ReviewDecisionStatus.Equivalent,
                    ReviewerSnapshot.Empty,
                    Utc(15)));
        }

        private static ReviewDecisionGeneration CreateGeneration()
        {
            var evidence = new ReviewEvidenceGeneration(
                "project-a",
                "mapping-a",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("r1", 0, Hash('1')),
                    new ReviewEvidenceRevisionIdentity("r2", 1, Hash('2')),
                    new ReviewEvidenceRevisionIdentity("r3", 2, Hash('3')),
                });
            var subjects = new[]
            {
                new ReviewDecisionSubject(
                    "subject-value",
                    "sheet:data/row:2/column:3/aspect:value",
                    ComparisonAspect.Value,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable),
                new ReviewDecisionSubject(
                    "subject-formula",
                    "sheet:data/row:2/column:2/aspect:formula",
                    ComparisonAspect.Formula,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable),
            };
            var observations = new[]
            {
                Formula("subject-formula", "r1", "A2+1", "r1.xlsx", "B2", 'a'),
                Formula("subject-formula", "r2", "A2+2", "r2.xlsx", "B2", 'b'),
                Formula("subject-formula", "r3", "A2+3", "r3.xlsx", "B2", 'c'),
                Value("subject-value", "r1", "one", "r1.xlsx", "C2", 'd'),
                Value("subject-value", "r2", "two", "r2.xlsx", "C2", 'e'),
                Value("subject-value", "r3", "three", "r3.xlsx", "C2", 'f'),
            };
            var catalog = new ReviewDecisionCatalog(
                evidence,
                "r3",
                subjects,
                observations,
                new[]
                {
                    new ReviewDecisionGroup(
                        "group-a",
                        new[] { "subject-value", "subject-formula" }),
                });
            return ReviewDecisionEngine.Create(catalog);
        }

        private static Observation Formula(
            string subject,
            string revision,
            string formula,
            string source,
            string address,
            char hash) =>
            new Observation(
                ObservationId(subject, revision, ObservationExactType.Formula, hash),
                subject,
                revision,
                ObservationExactType.Formula,
                Hash(hash),
                formula: formula,
                provenance: Provenance(source, address, 2));

        private static Observation Value(
            string subject,
            string revision,
            string value,
            string source,
            string address,
            char hash) =>
            new Observation(
                ObservationId(subject, revision, ObservationExactType.String, hash),
                subject,
                revision,
                ObservationExactType.String,
                Hash(hash),
                value: new CanonicalScalar(ObservationExactType.String, value),
                provenance: Provenance(source, address, 3));

        private static ObservationProvenance Provenance(
            string source,
            string address,
            int column) =>
            new ObservationProvenance(
                source,
                "Data",
                address,
                2,
                column,
                "logical-data");

        private static DateTimeOffset Utc(int minute) =>
            new DateTimeOffset(2026, 7, 28, 8, minute, 0, TimeSpan.Zero);

        private static string Hash(char value) =>
            new string(value, 64);

        private static string ObservationId(
            string subject,
            string revision,
            ObservationExactType exactType,
            char hash) =>
            ComparisonIds.Observation(subject, revision, exactType, Hash(hash));
    }
}
