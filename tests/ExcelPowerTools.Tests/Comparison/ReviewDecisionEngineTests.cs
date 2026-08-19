using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewDecisionEngineTests
    {
        private static readonly DateTimeOffset When = new DateTimeOffset(
            2026, 7, 17, 8, 9, 10, TimeSpan.Zero);
        private static readonly ReviewerSnapshot Reviewer = new ReviewerSnapshot(
            "Confirmed Reviewer",
            "Office Reviewer",
            new[] { "confirmed-dialog", "office-profile" });

        [Fact]
        public void Catalog_root_request_event_and_generation_ids_are_order_independent_and_deterministic()
        {
            ReviewDecisionCatalog leftCatalog = Catalog();
            ReviewDecisionCatalog rightCatalog = Catalog(reverseInputs: true);

            Assert.Equal(leftCatalog.CatalogId, rightCatalog.CatalogId);
            Assert.True(StableId.IsSha256(leftCatalog.CatalogId));
            ReviewDecisionGeneration leftRoot = ReviewDecisionEngine.Create(leftCatalog);
            ReviewDecisionGeneration rightRoot = ReviewDecisionEngine.Create(rightCatalog);
            Assert.Equal(leftRoot.GenerationId, rightRoot.GenerationId);
            Assert.Empty(leftRoot.PriorGenerationId);

            ReviewDecisionRequest leftRequest = ReviewDecisionRequest.ForSubjects(
                "project-1",
                "evidence-1",
                "mapping-1",
                leftRoot.GenerationId,
                new[] { "subject-2", "subject-1" },
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When,
                runId: "run-1",
                addInBuild: "1.2.3",
                note: "Same result");
            ReviewDecisionRequest rightRequest = ReviewDecisionRequest.ForSubjects(
                "project-1",
                "evidence-1",
                "mapping-1",
                rightRoot.GenerationId,
                new[] { "subject-1", "subject-2" },
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When,
                runId: "run-1",
                addInBuild: "1.2.3",
                note: "Same result");

            Assert.Equal(leftRequest.RequestId, rightRequest.RequestId);
            ReviewDecisionGeneration left = ReviewDecisionEngine.Apply(leftRoot, leftRequest);
            ReviewDecisionGeneration right = ReviewDecisionEngine.Apply(rightRoot, rightRequest);
            Assert.Equal(left.GenerationId, right.GenerationId);
            Assert.Equal(left.Events[0].EventId, right.Events[0].EventId);
            Assert.True(StableId.IsSha256(left.Events[0].EventId));
            Assert.True(StableId.IsSha256(left.GenerationId));
        }

        [Fact]
        public void Single_subject_replace_and_clear_append_superseding_events_without_mutating_prior_generations()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionGeneration candidate = ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(
                    root,
                    "subject-1",
                    ReviewDecisionStatus.UseCandidate,
                    "observation-1-candidate",
                    When,
                    "Candidate is correct"));

            DecisionEvent candidateEvent = Assert.Single(candidate.Events);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, candidateEvent.PriorDecision);
            Assert.Equal(ReviewDecisionStatus.UseCandidate, candidateEvent.NewDecision);
            Assert.Equal("observation-1-candidate", candidateEvent.SelectedObservationId);
            Assert.Equal(new[] { "candidate" }, candidateEvent.SupportingRevisionIds);
            Assert.Equal(new[] { "hash-1-candidate" }, candidateEvent.SupportingContentHashes);
            Assert.Empty(candidateEvent.SupersededEventId);
            Assert.Equal(DecisionSelectionScope.SingleSubject, candidateEvent.SelectionScope);
            Assert.Equal("subject-1", candidateEvent.SubjectId);
            Assert.Equal("run-1", candidateEvent.RunId);
            Assert.Equal("1.2.3", candidateEvent.AddInBuild);
            Assert.Equal("Candidate is correct", candidateEvent.Note);
            Assert.Same(Reviewer, candidateEvent.ReviewerSnapshot);
            Assert.Collection(
                candidateEvent.Locations,
                selected =>
                {
                    Assert.Equal(ReviewDecisionLocationRoles.Selected, selected.Role);
                    Assert.Equal("candidate", selected.RevisionId);
                    Assert.Equal("candidate", selected.SourceLabel);
                    Assert.Equal("Sheet-subject-1", selected.SheetName);
                    Assert.Equal("C11", selected.Address);
                    Assert.Equal(11, selected.PhysicalRow);
                    Assert.Equal(3, selected.PhysicalColumn);
                },
                masterLocation =>
                {
                    Assert.Equal(ReviewDecisionLocationRoles.Master, masterLocation.Role);
                    Assert.Equal("master", masterLocation.RevisionId);
                    Assert.Equal("master", masterLocation.SourceLabel);
                    Assert.Equal("Sheet-subject-1", masterLocation.SheetName);
                    Assert.Equal("C10", masterLocation.Address);
                    Assert.Equal(10, masterLocation.PhysicalRow);
                    Assert.Equal(3, masterLocation.PhysicalColumn);
                });
            Assert.True(candidate.TryGetActiveDecision("subject-1", out DecisionEvent activeCandidate));
            Assert.Same(candidateEvent, activeCandidate);

            ReviewDecisionGeneration master = ReviewDecisionEngine.Apply(
                candidate,
                SubjectRequest(
                    candidate,
                    "subject-1",
                    ReviewDecisionStatus.UseMaster,
                    "observation-1-master",
                    When.AddMinutes(1),
                    "Master wins"));
            DecisionEvent masterEvent = master.Events[1];
            Assert.Equal(candidateEvent.EventId, masterEvent.SupersededEventId);
            Assert.Equal(ReviewDecisionStatus.UseCandidate, masterEvent.PriorDecision);
            Assert.Equal(ReviewDecisionStatus.UseMaster, masterEvent.NewDecision);
            Assert.Equal(candidate.GenerationId, master.PriorGenerationId);
            Assert.Equal(
                new[] { ReviewDecisionLocationRoles.Selected, ReviewDecisionLocationRoles.Master },
                masterEvent.Locations.Select(item => item.Role));
            Assert.All(masterEvent.Locations, item => Assert.Equal("master", item.RevisionId));

            ReviewDecisionGeneration cleared = ReviewDecisionEngine.Apply(
                master,
                ReviewDecisionRequest.ClearSubject(
                    master.ProjectId,
                    master.BaseGenerationId,
                    master.MappingGenerationId,
                    master.GenerationId,
                    "subject-1",
                    Reviewer,
                    When.AddMinutes(2),
                    note: "Clear for another reviewer"));
            DecisionEvent clearEvent = cleared.Events[2];
            Assert.Equal(masterEvent.EventId, clearEvent.SupersededEventId);
            Assert.Equal(ReviewDecisionStatus.UseMaster, clearEvent.PriorDecision);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, clearEvent.NewDecision);
            Assert.False(cleared.TryGetActiveDecision("subject-1", out _));
            Assert.True(cleared.TryGetEffectiveEvent("subject-1", out DecisionEvent effectiveClear));
            Assert.Same(clearEvent, effectiveClear);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, cleared.GetSubjectState("subject-1").Status);
            Assert.Equal(clearEvent.EventId, cleared.GetSubjectState("subject-1").EffectiveEventId);

            Assert.Empty(root.Events);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, root.GetSubjectState("subject-1").Status);
            Assert.Single(candidate.Events);
            Assert.Equal(ReviewDecisionStatus.UseCandidate, candidate.GetSubjectState("subject-1").Status);
            Assert.Equal(2, master.Events.Count);
            Assert.Equal(3, cleared.Events.Count);
        }

        [Fact]
        public void Multi_and_group_decisions_freeze_membership_and_can_replace_and_clear_a_homogeneous_batch()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            var mutableSubjects = new List<string> { "subject-2", "subject-1" };
            ReviewDecisionRequest multiple = ReviewDecisionRequest.ForSubjects(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                mutableSubjects,
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When);
            mutableSubjects.Clear();
            mutableSubjects.Add("subject-3");

            ReviewDecisionGeneration equivalent = ReviewDecisionEngine.Apply(root, multiple);
            DecisionEvent multipleEvent = Assert.Single(equivalent.Events);
            Assert.Equal(DecisionSelectionScope.MultipleSubjects, multipleEvent.SelectionScope);
            Assert.Empty(multipleEvent.Locations);
            Assert.Equal(new[] { "subject-1", "subject-2" }, multiple.AffectedSubjectIds);
            Assert.Equal(new[] { "subject-1", "subject-2" }, multipleEvent.AffectedSubjectIds);
            Assert.Equal(multipleEvent.EventId, equivalent.GetSubjectState("subject-1").EffectiveEventId);
            Assert.Equal(multipleEvent.EventId, equivalent.GetSubjectState("subject-2").EffectiveEventId);

            var mutableGroupMembers = new List<string> { "subject-2", "subject-1" };
            ReviewDecisionRequest group = ReviewDecisionRequest.ForGroup(
                equivalent.ProjectId,
                equivalent.BaseGenerationId,
                equivalent.MappingGenerationId,
                equivalent.GenerationId,
                "group-12",
                mutableGroupMembers,
                ReviewDecisionStatus.NeedsInvestigation,
                Reviewer,
                When.AddMinutes(1));
            mutableGroupMembers[0] = "subject-3";
            ReviewDecisionGeneration investigating = ReviewDecisionEngine.Apply(equivalent, group);
            DecisionEvent groupEvent = investigating.Events[1];
            Assert.Equal(DecisionSelectionScope.Group, groupEvent.SelectionScope);
            Assert.Equal("group-12", groupEvent.GroupId);
            Assert.Equal(multipleEvent.EventId, groupEvent.SupersededEventId);
            Assert.Equal(new[] { "subject-1", "subject-2" }, groupEvent.AffectedSubjectIds);
            Assert.All(new[] { "subject-1", "subject-2" }, subjectId =>
                Assert.Equal(ReviewDecisionStatus.NeedsInvestigation,
                    investigating.GetSubjectState(subjectId).Status));

            ReviewDecisionGeneration cleared = ReviewDecisionEngine.Apply(
                investigating,
                ReviewDecisionRequest.ClearGroup(
                    investigating.ProjectId,
                    investigating.BaseGenerationId,
                    investigating.MappingGenerationId,
                    investigating.GenerationId,
                    "group-12",
                    new[] { "subject-2", "subject-1" },
                    Reviewer,
                    When.AddMinutes(2)));
            Assert.Equal(groupEvent.EventId, cleared.Events[2].SupersededEventId);
            Assert.Empty(cleared.ActiveDecisions);
            Assert.Equal(3, cleared.Events.Count);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, cleared.GetSubjectState("subject-1").Status);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, cleared.GetSubjectState("subject-2").Status);
            Assert.Equal(ReviewDecisionStatus.Unreviewed, cleared.GetSubjectState("subject-3").Status);
        }

        [Fact]
        public void Selection_must_exist_belong_to_the_subject_and_match_master_or_candidate_semantics()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());

            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(root, "subject-1", ReviewDecisionStatus.UseCandidate, "unknown", When)));
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(
                    root,
                    "subject-1",
                    ReviewDecisionStatus.UseCandidate,
                    "observation-2-candidate",
                    When)));
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(
                    root,
                    "subject-1",
                    ReviewDecisionStatus.UseMaster,
                    "observation-1-candidate",
                    When)));
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(
                    root,
                    "subject-1",
                    ReviewDecisionStatus.UseCandidate,
                    "observation-1-master",
                    When)));

            Assert.Throws<ArgumentException>(() => ReviewDecisionRequest.ForSubject(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "subject-1",
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When,
                selectedObservationId: "observation-1-master"));
            Assert.Throws<ArgumentException>(() => ReviewDecisionRequest.ForSubjects(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                new[] { "subject-1", "subject-2" },
                ReviewDecisionStatus.UseCandidate,
                Reviewer,
                When,
                selectedObservationId: "observation-1-candidate"));
        }

        [Fact]
        public void Selection_locations_are_empty_only_when_selected_and_master_provenance_are_absent()
        {
            var catalog = new ReviewDecisionCatalog(
                "project-empty-provenance",
                "evidence-empty-provenance",
                "mapping-empty-provenance",
                "master",
                new[] { SelectableSubject("subject") },
                new[]
                {
                    new Observation(
                        "master-observation",
                        "subject",
                        "master",
                        ObservationExactType.String,
                        "master-hash",
                        value: new CanonicalScalar(ObservationExactType.String, "master")),
                    new Observation(
                        "candidate-observation",
                        "subject",
                        "candidate",
                        ObservationExactType.String,
                        "candidate-hash",
                        value: new CanonicalScalar(ObservationExactType.String, "candidate")),
                });
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(catalog);
            ReviewDecisionGeneration applied = ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    "subject",
                    ReviewDecisionStatus.UseCandidate,
                    Reviewer,
                    When,
                    selectedObservationId: "candidate-observation"));

            Assert.Empty(Assert.Single(applied.Events).Locations);
        }

        [Fact]
        public void Selection_locations_preserve_source_only_and_coordinate_only_provenance()
        {
            var catalog = new ReviewDecisionCatalog(
                "project-sparse-provenance",
                "evidence-sparse-provenance",
                "mapping-sparse-provenance",
                "master",
                new[] { SelectableSubject("subject") },
                new[]
                {
                    new Observation(
                        "master-observation",
                        "subject",
                        "master",
                        ObservationExactType.String,
                        "master-hash",
                        value: new CanonicalScalar(ObservationExactType.String, "master"),
                        provenance: new ObservationProvenance(physicalRow: 17, physicalColumn: 4)),
                    new Observation(
                        "candidate-observation",
                        "subject",
                        "candidate",
                        ObservationExactType.String,
                        "candidate-hash",
                        value: new CanonicalScalar(ObservationExactType.String, "candidate"),
                        provenance: new ObservationProvenance(sourceLabel: "emailed-copy.xlsx")),
                });
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(catalog);
            ReviewDecisionGeneration applied = ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    "subject",
                    ReviewDecisionStatus.UseCandidate,
                    Reviewer,
                    When,
                    selectedObservationId: "candidate-observation"));

            Assert.Collection(
                Assert.Single(applied.Events).Locations,
                selected =>
                {
                    Assert.Equal(ReviewDecisionLocationRoles.Selected, selected.Role);
                    Assert.Equal("emailed-copy.xlsx", selected.SourceLabel);
                    Assert.Empty(selected.SheetName);
                    Assert.Empty(selected.Address);
                    Assert.Equal(0, selected.PhysicalRow);
                    Assert.Equal(0, selected.PhysicalColumn);
                },
                master =>
                {
                    Assert.Equal(ReviewDecisionLocationRoles.Master, master.Role);
                    Assert.Empty(master.SourceLabel);
                    Assert.Empty(master.SheetName);
                    Assert.Empty(master.Address);
                    Assert.Equal(17, master.PhysicalRow);
                    Assert.Equal(4, master.PhysicalColumn);
                });
        }

        [Fact]
        public void Selection_location_preserves_logical_sheet_only_provenance()
        {
            var catalog = new ReviewDecisionCatalog(
                "project-logical-provenance",
                "evidence-logical-provenance",
                "mapping-logical-provenance",
                "master",
                new[] { SelectableSubject("subject") },
                new[]
                {
                    new Observation(
                        "master-observation",
                        "subject",
                        "master",
                        ObservationExactType.String,
                        "master-hash",
                        value: new CanonicalScalar(ObservationExactType.String, "master")),
                    new Observation(
                        "candidate-observation",
                        "subject",
                        "candidate",
                        ObservationExactType.String,
                        "candidate-hash",
                        value: new CanonicalScalar(ObservationExactType.String, "candidate"),
                        provenance: new ObservationProvenance(logicalSheetId: "logical-sheet")),
                });
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(catalog);
            ReviewDecisionGeneration applied = ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    "subject",
                    ReviewDecisionStatus.UseCandidate,
                    Reviewer,
                    When,
                    selectedObservationId: "candidate-observation"));

            DecisionLocationContext location = Assert.Single(Assert.Single(applied.Events).Locations);
            Assert.Equal(ReviewDecisionLocationRoles.Selected, location.Role);
            Assert.Equal("logical-sheet", location.LogicalSheetId);
            Assert.Empty(location.SourceLabel);
            Assert.Empty(location.SheetName);
            Assert.Empty(location.Address);
            Assert.Equal(0, location.PhysicalRow);
            Assert.Equal(0, location.PhysicalColumn);
        }

        [Fact]
        public void Decision_locations_reject_negative_physical_coordinates()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new DecisionLocationContext("selected", "revision", physicalRow: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new DecisionLocationContext("selected", "revision", physicalColumn: -1));
        }

        [Fact]
        public void Unknown_subject_group_and_changed_group_membership_are_rejected()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(root, "unknown-subject", ReviewDecisionStatus.Equivalent, "", When)));

            ReviewDecisionRequest unknownGroup = ReviewDecisionRequest.ForGroup(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "unknown-group",
                new[] { "subject-1" },
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When);
            Assert.Throws<InvalidOperationException>(() =>
                ReviewDecisionEngine.Apply(root, unknownGroup));

            ReviewDecisionRequest wrongMembers = ReviewDecisionRequest.ForGroup(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "group-12",
                new[] { "subject-1", "subject-3" },
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When);
            Assert.Throws<InvalidOperationException>(() =>
                ReviewDecisionEngine.Apply(root, wrongMembers));
        }

        [Theory]
        [InlineData("wrong-project", "evidence-1", "mapping-1")]
        [InlineData("project-1", "stale-evidence", "mapping-1")]
        [InlineData("project-1", "evidence-1", "stale-mapping")]
        public void Stale_project_evidence_or_mapping_context_is_rejected(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId)
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionRequest request = ReviewDecisionRequest.ForSubject(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                root.GenerationId,
                "subject-1",
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When);

            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(root, request));
        }

        [Fact]
        public void Stale_decision_generation_is_rejected_by_optimistic_concurrency_check()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionRequest stale = SubjectRequest(
                root,
                "subject-2",
                ReviewDecisionStatus.Equivalent,
                "",
                When);
            ReviewDecisionGeneration advanced = ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(root, "subject-1", ReviewDecisionStatus.Equivalent, "", When));

            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(advanced, stale));
        }

        [Fact]
        public void Duplicate_no_op_and_mixed_provenance_batches_are_rejected()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            Assert.Throws<ArgumentException>(() => ReviewDecisionRequest.ForSubjects(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                new[] { "subject-1", "subject-1" },
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When));

            ReviewDecisionGeneration first = ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(root, "subject-1", ReviewDecisionStatus.Equivalent, "", When));
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                first,
                SubjectRequest(
                    first,
                    "subject-1",
                    ReviewDecisionStatus.Equivalent,
                    "",
                    When.AddMinutes(1))));

            ReviewDecisionGeneration second = ReviewDecisionEngine.Apply(
                first,
                SubjectRequest(
                    first,
                    "subject-2",
                    ReviewDecisionStatus.Equivalent,
                    "",
                    When.AddMinutes(1)));
            ReviewDecisionRequest ambiguous = ReviewDecisionRequest.ForSubjects(
                second.ProjectId,
                second.BaseGenerationId,
                second.MappingGenerationId,
                second.GenerationId,
                new[] { "subject-1", "subject-2" },
                ReviewDecisionStatus.NeedsInvestigation,
                Reviewer,
                When.AddMinutes(2));
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                ReviewDecisionEngine.Apply(second, ambiguous));
            Assert.Contains("mixed prior decision provenance", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Catalog_rejects_duplicate_or_ambiguous_evidence_and_unknown_group_members()
        {
            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                new[] { "subject-1", "subject-1" },
                Array.Empty<Observation>()));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                new[] { "subject-1" },
                new[]
                {
                    Observation("same-id", "subject-1", "master", "one"),
                    Observation("same-id", "subject-1", "candidate", "two"),
                }));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                new[] { "subject-1" },
                new[]
                {
                    Observation("observation-a", "subject-1", "master", "one"),
                    Observation("observation-b", "subject-1", "master", "two"),
                }));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                new[] { "subject-1" },
                new[] { Observation("observation-a", "unknown", "master", "one") }));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                new[] { "subject-1" },
                Array.Empty<Observation>(),
                new[] { new ReviewDecisionGroup("group", new[] { "unknown" }) }));
            Assert.Throws<ArgumentException>(() =>
                new ReviewDecisionGroup("group", new[] { "subject-1", "subject-1" }));
        }

        [Fact]
        public void Public_catalog_request_and_generation_views_are_read_only()
        {
            var sourceSubjects = new List<string> { "subject-3", "subject-1", "subject-2" };
            var sourceObservations = Observations().ToList();
            var sourceGroups = new List<ReviewDecisionGroup>
            {
                new ReviewDecisionGroup("group-12", new[] { "subject-2", "subject-1" }),
            };
            var catalog = new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                sourceSubjects,
                sourceObservations,
                sourceGroups);
            sourceSubjects.Clear();
            sourceObservations.Clear();
            sourceGroups.Clear();

            Assert.Equal(3, catalog.SubjectIds.Count);
            Assert.Equal(6, catalog.Observations.Count);
            Assert.Single(catalog.Groups);
            Assert.Throws<NotSupportedException>(() =>
                ((IList<string>)catalog.SubjectIds).Add("subject-4"));
            Assert.Throws<NotSupportedException>(() =>
                ((IList<Observation>)catalog.Observations).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<string>)catalog.Groups[0].SubjectIds)[0] = "subject-3");

            ReviewDecisionGeneration generation = ReviewDecisionEngine.Apply(
                ReviewDecisionEngine.Create(catalog),
                ReviewDecisionRequest.ForSubjects(
                    catalog.ProjectId,
                    catalog.BaseGenerationId,
                    catalog.MappingGenerationId,
                    ReviewDecisionEngine.Create(catalog).GenerationId,
                    new[] { "subject-1", "subject-2" },
                    ReviewDecisionStatus.Equivalent,
                    Reviewer,
                    When));
            Assert.Throws<NotSupportedException>(() =>
                ((IDictionary<string, ReviewDecisionState>)generation.SubjectStates).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IDictionary<string, DecisionEvent>)generation.ActiveDecisions).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<DecisionEvent>)generation.Events).Clear());
        }

        [Fact]
        public void Requests_require_utc_timestamps_and_clear_of_untouched_state_is_rejected_as_no_op()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            Assert.Throws<ArgumentException>(() => ReviewDecisionRequest.ForSubject(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "subject-1",
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                new DateTimeOffset(2026, 7, 17, 10, 0, 0, TimeSpan.FromHours(2))));
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ClearSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    "subject-1",
                    Reviewer,
                    When)));
        }

        private static ReviewDecisionRequest SubjectRequest(
            ReviewDecisionGeneration generation,
            string subjectId,
            ReviewDecisionStatus decision,
            string selectedObservationId,
            DateTimeOffset createdUtc,
            string note = "") =>
            ReviewDecisionRequest.ForSubject(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.GenerationId,
                subjectId,
                decision,
                Reviewer,
                createdUtc,
                selectedObservationId,
                "run-1",
                "1.2.3",
                note);

        private static ReviewDecisionCatalog Catalog(bool reverseInputs = false)
        {
            IEnumerable<ReviewDecisionSubject> subjects = new[]
            {
                SelectableSubject("subject-3"),
                SelectableSubject("subject-1", ComparisonAspect.Formula),
                SelectableSubject("subject-2"),
            };
            IEnumerable<Observation> observations = Observations();
            IEnumerable<ReviewDecisionGroup> groups = new[]
            {
                new ReviewDecisionGroup("group-12", new[] { "subject-2", "subject-1" }),
            };
            if (reverseInputs)
            {
                subjects = subjects.Reverse();
                observations = observations.Reverse();
                groups = groups.Reverse();
            }
            return new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                subjects,
                observations,
                groups);
        }

        private static IEnumerable<Observation> Observations() => new[]
        {
            Observation("observation-1-master", "subject-1", "master", "1-master"),
            Observation("observation-1-candidate", "subject-1", "candidate", "1-candidate"),
            Observation("observation-2-master", "subject-2", "master", "2-master"),
            Observation("observation-2-candidate", "subject-2", "candidate", "2-candidate"),
            Observation("observation-3-master", "subject-3", "master", "3-master"),
            Observation("observation-3-candidate", "subject-3", "candidate", "3-candidate"),
        };

        private static Observation Observation(
            string observationId,
            string subjectId,
            string revisionId,
            string text) =>
            new Observation(
                observationId,
                subjectId,
                revisionId,
                ObservationExactType.String,
                "hash-" + text,
                value: new CanonicalScalar(ObservationExactType.String, text),
                provenance: new ObservationProvenance(
                    sourceLabel: revisionId,
                    sheetName: "Sheet-" + subjectId,
                    address: revisionId == "master" ? "C10" : "C11",
                    physicalRow: revisionId == "master" ? 10 : 11,
                    physicalColumn: 3,
                    logicalSheetId: "logical-sheet-" + subjectId));

        private static ReviewDecisionSubject SelectableSubject(
            string subjectId,
            ComparisonAspect aspect = ComparisonAspect.Value) =>
            new ReviewDecisionSubject(
                subjectId,
                "logical-" + subjectId,
                aspect,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
    }
}
