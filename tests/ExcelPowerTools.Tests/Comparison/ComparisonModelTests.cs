using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class ComparisonModelTests
    {
        [Fact]
        public void Project_manifest_keeps_distinct_revision_occurrences_with_shared_content()
        {
            string byteContentId = StableId.ForBytes(Encoding.UTF8.GetBytes("same artifact bytes"));
            var project = new ComparisonProject { ProjectId = "project-1" };
            project.Revisions.Add(new Revision
            {
                RevisionId = "revision-email-a",
                ByteContentId = byteContentId,
                Label = "emailed copy"
            });
            project.Revisions.Add(new Revision
            {
                RevisionId = "revision-folder-b",
                ByteContentId = byteContentId,
                Label = "folder copy"
            });

            Assert.Equal(2, project.Revisions.Count);
            Assert.NotEqual(project.Revisions[0].RevisionId, project.Revisions[1].RevisionId);
            Assert.Equal(project.Revisions[0].ByteContentId, project.Revisions[1].ByteContentId);
        }

        [Fact]
        public void Evidence_decision_and_merge_state_are_independent()
        {
            var observation = new Observation(
                "observation-1",
                "subject-1",
                "revision-1",
                ObservationExactType.Formula,
                "hash-before-decision",
                formula: "=A1+B1");
            var decision = new DecisionEvent(
                "decision-1",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.UseCandidate,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.SingleSubject,
                subjectId: "subject-1",
                selectedObservationId: observation.ObservationId);
            var merge = new MergeRun
            {
                MergeId = "merge-1",
                State = MergeRunState.Prepared
            };
            merge.SelectedDecisionIds.Add(decision.EventId);
            merge.State = MergeRunState.AuditCommitted;

            Assert.Equal("hash-before-decision", observation.ContentHash);
            Assert.Equal(ReviewDecisionStatus.UseCandidate, decision.NewDecision);
            Assert.Equal(MergeRunState.AuditCommitted, merge.State);
        }

        [Fact]
        public void Durable_contract_collections_are_ready_for_population()
        {
            Assert.NotNull(new Revision().Coverage.UnsupportedFeatures);
            Assert.NotNull(new Transition().EquivalenceTags);
            var decision = new DecisionEvent(
                "decision-1",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.SingleSubject,
                subjectId: "subject-1");
            Assert.NotNull(decision.AffectedSubjectIds);
            Assert.NotNull(new MergeRun().PreconditionResults);
        }

        [Fact]
        public void Observation_constructor_enforces_type_and_cache_invariants()
        {
            var number = new Observation(
                "observation-1",
                "subject-1",
                "revision-1",
                ObservationExactType.Number,
                "content-hash",
                value: new CanonicalScalar(ObservationExactType.Number, "1.25"));

            Assert.Equal("1.25", number.Value.CanonicalText);
            Assert.Null(number.CachedResult);
            Assert.Throws<ArgumentException>(() => new Observation(
                "observation-2",
                "subject-1",
                "revision-1",
                ObservationExactType.Number,
                "content-hash",
                value: new CanonicalScalar(ObservationExactType.String, "1.25")));
            Assert.Throws<ArgumentException>(() => new Observation(
                "observation-3",
                "subject-1",
                "revision-1",
                ObservationExactType.Formula,
                "content-hash",
                formula: "=A1",
                cachedResult: new CanonicalScalar(ObservationExactType.Number, "2")));

            var isoDate = new CanonicalScalar(
                ObservationExactType.IsoDateTime,
                "2026-07-16T14:15:16.123");
            Assert.Equal(ExcelDateSystem.NotApplicable, isoDate.DateSystem);
            Assert.Throws<ArgumentOutOfRangeException>(() => new CanonicalScalar(
                ObservationExactType.IsoDateTime,
                "2026-07-16T14:15:16.123",
                ExcelDateSystem.Date1900));
        }

        [Fact]
        public void Decision_event_freezes_membership_and_validates_scope()
        {
            var affected = new List<string> { "subject-a", "subject-b" };
            var decision = new DecisionEvent(
                "decision-1",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.MultipleSubjects,
                affectedSubjectIds: affected);

            affected.Add("subject-c");
            Assert.Equal(new[] { "subject-a", "subject-b" }, decision.AffectedSubjectIds);
            Assert.Throws<NotSupportedException>(() =>
                ((IList<string>)decision.AffectedSubjectIds).Add("subject-c"));
            Assert.Throws<ArgumentException>(() => new DecisionEvent(
                "decision-2",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.MultipleSubjects,
                affectedSubjectIds: new[] { "only-one" }));
        }

        [Fact]
        public void Decision_event_rejects_observation_selection_outside_single_subject_selection_statuses()
        {
            Assert.Throws<ArgumentException>(() => new DecisionEvent(
                "decision-multi-selection",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.UseCandidate,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.MultipleSubjects,
                affectedSubjectIds: new[] { "subject-a", "subject-b" },
                selectedObservationId: "observation-a"));

            Assert.Throws<ArgumentException>(() => new DecisionEvent(
                "decision-non-selection",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.Equivalent,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.SingleSubject,
                subjectId: "subject-a",
                selectedObservationId: "observation-a"));
        }

        [Fact]
        public void Constructor_aware_json_rejects_persisted_multi_subject_observation_selection()
        {
            var valid = new DecisionEvent(
                "decision-1",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.Equivalent,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.MultipleSubjects,
                affectedSubjectIds: new[] { "subject-a", "subject-b" });
            string persisted = JsonSerializer.Serialize(valid)
                .Replace("\"NewDecision\":3", "\"NewDecision\":2")
                .Replace(
                    "\"SelectedObservationId\":\"\"",
                    "\"SelectedObservationId\":\"observation-a\"");

            Assert.Contains("\"NewDecision\":2", persisted, StringComparison.Ordinal);
            Assert.Contains("\"SelectedObservationId\":\"observation-a\"", persisted, StringComparison.Ordinal);
            Assert.Throws<ArgumentException>(() =>
                JsonSerializer.Deserialize<DecisionEvent>(persisted));
        }

        [Fact]
        public void Legacy_decision_location_json_defaults_new_provenance_fields()
        {
            const string persisted =
                "{\"Role\":\"selected\",\"RevisionId\":\"revision-1\"," +
                "\"LogicalSheetId\":\"sheet-1\",\"SheetName\":\"Data\",\"Address\":\"A1\"}";

            DecisionLocationContext restored =
                JsonSerializer.Deserialize<DecisionLocationContext>(persisted)!;

            Assert.Equal("selected", restored.Role);
            Assert.Equal("revision-1", restored.RevisionId);
            Assert.Equal("sheet-1", restored.LogicalSheetId);
            Assert.Equal("Data", restored.SheetName);
            Assert.Equal("A1", restored.Address);
            Assert.Empty(restored.SourceLabel);
            Assert.Equal(0, restored.PhysicalRow);
            Assert.Equal(0, restored.PhysicalColumn);
        }

        [Fact]
        public void Revision_content_hash_aliases_cannot_diverge()
        {
            string first = StableId.ForBytes(Encoding.UTF8.GetBytes("first"));
            string second = StableId.ForBytes(Encoding.UTF8.GetBytes("second"));
            var revision = new Revision { Sha256 = first.ToLowerInvariant() };

            Assert.Equal(first, revision.ByteContentId);
            Assert.Equal(first, revision.Sha256);
            revision.ByteContentId = first;
            Assert.Throws<InvalidOperationException>(() => revision.ByteContentId = second);
            Assert.Throws<InvalidOperationException>(() => revision.Sha256 = "");
            Assert.Throws<ArgumentException>(() => new Revision { ByteContentId = "not-a-sha256" });
        }

        [Fact]
        public void Immutable_evidence_contracts_round_trip_with_constructor_aware_json()
        {
            var observation = new Observation(
                "observation-1",
                "subject-1",
                "revision-1",
                ObservationExactType.Formula,
                "content-hash",
                formula: "=A1",
                cachedResult: new CanonicalScalar(ObservationExactType.Number, "2"),
                cachedResultTrust: CachedResultTrust.Unknown,
                provenance: new ObservationProvenance("source", "Data", "A1", 1, 1));
            var decision = new DecisionEvent(
                "decision-1",
                "project-1",
                "evidence-generation-1",
                "mapping-generation-1",
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.UseCandidate,
                new DateTimeOffset(2026, 7, 16, 10, 0, 0, TimeSpan.Zero),
                DecisionSelectionScope.SingleSubject,
                subjectId: "subject-1",
                selectedObservationId: "observation-1",
                reviewerSnapshot: new ReviewerSnapshot("Reviewer", "Office", new[] { "confirmed" }),
                locations: new[]
                {
                    new DecisionLocationContext(
                        "selected",
                        "revision-1",
                        sheetName: "Data",
                        address: "A1",
                        sourceLabel: "source",
                        physicalRow: 1,
                        physicalColumn: 1),
                });

            Observation restoredObservation = JsonSerializer.Deserialize<Observation>(
                JsonSerializer.Serialize(observation))!;
            DecisionEvent restoredDecision = JsonSerializer.Deserialize<DecisionEvent>(
                JsonSerializer.Serialize(decision))!;

            Assert.Equal(observation.Formula, restoredObservation.Formula);
            Assert.Equal(observation.CachedResult.CanonicalText, restoredObservation.CachedResult.CanonicalText);
            Assert.Equal(decision.AffectedSubjectIds, restoredDecision.AffectedSubjectIds);
            Assert.Equal("Reviewer", restoredDecision.ReviewerSnapshot.ConfirmedDisplayName);
            DecisionLocationContext restoredLocation = Assert.Single(restoredDecision.Locations);
            Assert.Equal("source", restoredLocation.SourceLabel);
            Assert.Equal(1, restoredLocation.PhysicalRow);
            Assert.Equal(1, restoredLocation.PhysicalColumn);
        }
    }
}
