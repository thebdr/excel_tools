using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Merge;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergePlanBuilderTests
    {
        private const string CandidateRevision = "candidate";
        private const string MasterRevision = "master";
        private const string PrivateSourceLabel =
            @"C:\private\emailed copies\customer-secret-master.xlsx";
        private const string PrivateReviewer = "Private Reviewer";
        private static readonly DateTimeOffset When =
            new DateTimeOffset(2026, 7, 23, 10, 11, 12, TimeSpan.Zero);

        [Fact]
        public void Builds_deterministic_exact_formula_and_value_plan()
        {
            SubjectDefinition[] definitions =
            {
                Value("value", 2, "old-value", "new-value"),
                Formula("formula", 1, "=OLD(A1)", "=NEW(A1)"),
            };
            ReviewDecisionGeneration generation = Generation(
                definitions,
                new[]
                {
                    Decision("formula", ReviewDecisionStatus.UseCandidate),
                    Decision("value", ReviewDecisionStatus.UseCandidate),
                });
            MergePlanRequest reversed = MergePlanRequest.Capture(
                generation,
                new[] { "value", "formula" });
            MergePlanRequest forward = MergePlanRequest.Capture(
                generation,
                new[] { "formula", "value" });

            MergePlan left = MergePlanBuilder.Build(generation, reversed);
            MergePlan right = MergePlanBuilder.Build(generation, forward);

            Assert.Equal(left.DecisionSetHash, right.DecisionSetHash);
            Assert.Equal(left.PlanHash, right.PlanHash);
            Assert.Equal(64, left.DecisionSetHash.Length);
            Assert.Equal(64, left.PlanHash.Length);
            Assert.Equal(generation.ProjectId, left.ProjectId);
            Assert.Equal(generation.BaseGenerationId, left.EvidenceGenerationId);
            Assert.Equal(generation.MappingGenerationId, left.MappingGenerationId);
            Assert.Equal(generation.CatalogId, left.CatalogId);
            Assert.Equal(generation.GenerationId, left.DecisionGenerationId);
            Assert.Equal(MasterRevision, left.MasterRevisionId);
            Assert.Equal(Hash("master-file"), left.MasterByteContentId);
            Assert.Equal(new[] { "formula", "value" }, left.Operations.Select(item => item.SubjectId));
            Assert.Equal(
                left.Operations.Select(item => item.DecisionEventId),
                left.SelectedDecisionIds);
            Assert.All(left.Operations, operation => Assert.Equal(64, operation.OperationId.Length));

            MergePlanOperation formula = left.Operations[0];
            Assert.Equal(ComparisonAspect.Formula, formula.Aspect);
            Assert.Equal("Data!A1", formula.Target.SheetName + "!" + formula.Target.Address);
            Assert.Equal(ObservationExactType.Formula, formula.MasterPrecondition.ExactType);
            Assert.Equal("=OLD(A1)", formula.MasterPrecondition.Formula);
            Assert.Null(formula.MasterPrecondition.Value);
            Assert.Equal(ObservationExactType.Formula, formula.SelectedOutput.ExactType);
            Assert.Equal("=NEW(A1)", formula.SelectedOutput.Formula);
            Assert.Null(formula.SelectedOutput.Value);
            Assert.Equal("Candidate Data", formula.SelectedSource.SheetName);
            Assert.Equal("A1", formula.SelectedSource.Address);
            Assert.Equal(
                formula.Target.LogicalSheetId,
                formula.SelectedSource.LogicalSheetId);

            MergePlanOperation value = left.Operations[1];
            Assert.Equal(ComparisonAspect.Value, value.Aspect);
            Assert.Equal("A2", value.Target.Address);
            Assert.Equal("old-value", value.MasterPrecondition.Value.CanonicalText);
            Assert.Equal("new-value", value.SelectedOutput.Value.CanonicalText);
            Assert.Equal(ObservationExactType.String, value.SelectedOutput.ExactType);
            Assert.Equal(Hash("candidate-file"), value.SelectedRevisionByteContentId);

            Assert.True(left.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.CellWriteSynthesis));
            Assert.True(left.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FormulaDependencies));
            Assert.True(left.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FormulaGroups));
            Assert.True(left.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FullOutputSemanticDiff));
            Assert.True(left.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity));
        }

        [Fact]
        public void Catalog_and_request_enumeration_order_do_not_change_plan()
        {
            SubjectDefinition[] definitions =
            {
                Formula("formula", 3, "=A1", "=A2"),
                Value("value", 7, "10", "11"),
            };
            ReviewDecisionSpec[] decisions =
            {
                Decision("formula", ReviewDecisionStatus.UseCandidate),
                Decision("value", ReviewDecisionStatus.UseCandidate),
            };
            ReviewDecisionGeneration normal = Generation(definitions, decisions);
            ReviewDecisionGeneration reversedCatalog = Generation(
                definitions.AsEnumerable().Reverse().ToArray(),
                decisions,
                reverseCatalogInputs: true);

            MergePlan normalPlan = MergePlanBuilder.Build(
                normal,
                MergePlanRequest.Capture(normal, new[] { "value", "formula" }));
            MergePlan reversedPlan = MergePlanBuilder.Build(
                reversedCatalog,
                MergePlanRequest.Capture(reversedCatalog, new[] { "formula", "value" }));

            Assert.Equal(normal.CatalogId, reversedCatalog.CatalogId);
            Assert.Equal(normal.GenerationId, reversedCatalog.GenerationId);
            Assert.Equal(normalPlan.DecisionSetHash, reversedPlan.DecisionSetHash);
            Assert.Equal(normalPlan.PlanHash, reversedPlan.PlanHash);
            Assert.Equal(
                normalPlan.Operations.Select(item => item.OperationId),
                reversedPlan.Operations.Select(item => item.OperationId));
        }

        [Fact]
        public void Value_only_plan_does_not_claim_formula_checks_are_needed()
        {
            SubjectDefinition definition = Value("value", 1, "before", "after");
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlan plan = Build(generation, "value");

            Assert.False(plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FormulaDependencies));
            Assert.False(plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FormulaGroups));
            Assert.True(plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.SourceMasterByteIdentity));
            Assert.True(plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.MasterTargetPreconditions));
        }

        [Fact]
        public void Formula_and_value_aspects_can_share_a_cell_but_remain_nonsequential()
        {
            SubjectDefinition formula = Formula("formula", 1, "=OLD()", "=NEW()");
            SubjectDefinition value = Value("value", 1, "old", "new");
            value.LogicalSubjectKey = formula.LogicalSubjectKey;
            ReviewDecisionGeneration generation = Generation(
                new[] { value, formula },
                new[]
                {
                    Decision("value", ReviewDecisionStatus.UseCandidate),
                    Decision("formula", ReviewDecisionStatus.UseCandidate),
                });

            MergePlan plan = MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(generation, new[] { "formula", "value" }));

            Assert.Equal(2, plan.Operations.Count);
            Assert.All(plan.Operations, operation => Assert.Equal("A1", operation.Target.Address));
            Assert.Equal(
                new[] { ComparisonAspect.Formula, ComparisonAspect.Value },
                plan.Operations.Select(operation => operation.Aspect));
            Assert.True(plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.CellWriteSynthesis));
        }

        [Fact]
        public void Review_summary_freezes_every_effective_status_and_partial_selection()
        {
            SubjectDefinition[] definitions =
            {
                Value("selected", 1, "a", "b"),
                Value("candidate-not-selected", 2, "a", "b"),
                Value("master", 3, "a", "b"),
                Value("equivalent", 4, "a", "b"),
                Value("investigate", 5, "a", "b"),
                Structure("structural", 6),
                Value("unreviewed", 7, "a", "b"),
            };
            ReviewDecisionGeneration generation = Generation(
                definitions,
                new[]
                {
                    Decision("selected", ReviewDecisionStatus.UseCandidate),
                    Decision("candidate-not-selected", ReviewDecisionStatus.UseCandidate),
                    Decision("master", ReviewDecisionStatus.UseMaster),
                    Decision("equivalent", ReviewDecisionStatus.Equivalent),
                    Decision("investigate", ReviewDecisionStatus.NeedsInvestigation),
                    Decision("structural", ReviewDecisionStatus.ReviewedNoStructuralMerge),
                });

            MergePlan plan = Build(generation, "selected");

            Assert.Equal(7, plan.ReviewSummary.TotalSubjects);
            Assert.Equal(1, plan.ReviewSummary.Unreviewed);
            Assert.Equal(1, plan.ReviewSummary.UseMaster);
            Assert.Equal(2, plan.ReviewSummary.UseCandidate);
            Assert.Equal(1, plan.ReviewSummary.Equivalent);
            Assert.Equal(1, plan.ReviewSummary.NeedsInvestigation);
            Assert.Equal(1, plan.ReviewSummary.ReviewedNoStructuralMerge);
            Assert.Equal(1, plan.ReviewSummary.SelectedCandidate);
            Assert.Equal(1, plan.ReviewSummary.UnselectedCandidate);
        }

        [Theory]
        [InlineData("project")]
        [InlineData("evidence")]
        [InlineData("mapping")]
        [InlineData("catalog")]
        [InlineData("decision")]
        [InlineData("master")]
        public void Rejects_every_stale_request_binding(string changedBinding)
        {
            ReviewDecisionGeneration generation = Generation(
                new[] { Value("value", 1, "before", "after") },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });
            MergePlanRequest captured = MergePlanRequest.Capture(
                generation,
                new[] { "value" });
            string Wrong(string name, string value) =>
                changedBinding == name ? value + "-stale" : value;
            var stale = new MergePlanRequest(
                Wrong("project", captured.ProjectId),
                Wrong("evidence", captured.EvidenceGenerationId),
                Wrong("mapping", captured.MappingGenerationId),
                Wrong("catalog", captured.CatalogId),
                Wrong("decision", captured.ExpectedDecisionGenerationId),
                Wrong("master", captured.MasterRevisionId),
                captured.Selections);

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                MergePlanBuilder.Build(generation, stale));

            Assert.Equal(MergePlanBuildCodes.ContextMismatch, exception.Code);
            Assert.Contains("Corrective action:", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Request_becomes_stale_after_an_unrelated_decision_changes_generation()
        {
            SubjectDefinition[] definitions =
            {
                Value("selected", 1, "a", "b"),
                Value("other", 2, "a", "b"),
            };
            ReviewDecisionGeneration selected = Generation(
                definitions,
                new[] { Decision("selected", ReviewDecisionStatus.UseCandidate) });
            MergePlanRequest captured = MergePlanRequest.Capture(
                selected,
                new[] { "selected" });
            ReviewDecisionGeneration advanced = Apply(
                selected,
                Decision("other", ReviewDecisionStatus.Equivalent),
                selected.Events.Count);

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                MergePlanBuilder.Build(advanced, captured));

            Assert.Equal(MergePlanBuildCodes.ContextMismatch, exception.Code);
        }

        [Fact]
        public void Rejects_opaque_non_evidence_bound_catalog()
        {
            SubjectDefinition definition = Value("value", 1, "before", "after");
            ReviewDecisionCatalog catalog = Catalog(
                new[] { definition },
                evidenceBound: false);
            ReviewDecisionGeneration generation = Apply(
                ReviewDecisionEngine.Create(catalog),
                Decision("value", ReviewDecisionStatus.UseCandidate),
                0);

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "value"));

            Assert.Equal(MergePlanBuildCodes.EvidenceRequired, exception.Code);
        }

        [Fact]
        public void Rejects_selectable_but_nonmergeable_subject()
        {
            SubjectDefinition definition = Value("value", 1, "before", "after");
            definition.Policy = ReviewDecisionSubjectPolicy.Selectable;
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "value"));

            Assert.Equal(MergePlanBuildCodes.SubjectNotMergeable, exception.Code);
            Assert.Equal("value", exception.SubjectId);
        }

        [Fact]
        public void Corrupt_structural_candidate_state_cannot_become_an_operation()
        {
            SubjectDefinition definition = Structure("structural", 1);
            ReviewDecisionCatalog catalog = Catalog(new[] { definition });
            Observation selected = catalog.Observations.Single(
                item => item.SubjectId == "structural" &&
                        item.RevisionId == CandidateRevision);
            ReviewDecisionGeneration generation = CorruptCandidateGeneration(
                catalog,
                selected,
                supportingContentHash: selected.ContentHash);
            var request = new MergePlanRequest(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.CatalogId,
                generation.GenerationId,
                generation.Catalog.MasterRevisionId,
                new[] { new MergePlanSelection("structural", generation.Events[0].EventId) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                MergePlanBuilder.Build(generation, request));

            Assert.Equal(MergePlanBuildCodes.SubjectNotMergeable, exception.Code);
        }

        [Fact]
        public void Rejects_redundant_selection_with_no_actual_change()
        {
            SubjectDefinition definition = Value("value", 1, "same", "same");
            definition.MasterContentSeed = "same-state";
            definition.CandidateContentSeed = "same-state";
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "value"));

            Assert.Equal(MergePlanBuildCodes.NoActualChange, exception.Code);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Rejects_payload_and_content_identity_disagreement(bool sameHash)
        {
            SubjectDefinition definition = Value(
                "value",
                1,
                sameHash ? "before" : "same",
                sameHash ? "after" : "same");
            definition.MasterContentSeed = sameHash ? "forged-shared-hash" : "master-hash";
            definition.CandidateContentSeed = sameHash ? "forged-shared-hash" : "candidate-hash";
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "value"));

            Assert.Equal(MergePlanBuildCodes.EvidenceInconsistent, exception.Code);
        }

        [Theory]
        [InlineData(ComparisonAspect.Formula, ObservationExactType.Missing)]
        [InlineData(ComparisonAspect.Value, ObservationExactType.Missing)]
        [InlineData(ComparisonAspect.Value, ObservationExactType.Unsupported)]
        public void Rejects_states_without_unambiguous_v1_write_meaning(
            ComparisonAspect aspect,
            ObservationExactType candidateType)
        {
            SubjectDefinition definition = aspect == ComparisonAspect.Formula
                ? Formula("subject", 1, "=OLD()", "=NEW()")
                : Value("subject", 1, "before", "after");
            definition.CandidateType = candidateType;
            definition.CandidateText = candidateType == ObservationExactType.Unsupported
                ? "unsupported-payload"
                : "";
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("subject", ReviewDecisionStatus.UseCandidate) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "subject"));

            Assert.Equal(MergePlanBuildCodes.StateUnsupported, exception.Code);
        }

        [Theory]
        [InlineData("")]
        [InlineData("B1")]
        [InlineData("$A$1")]
        public void Rejects_missing_or_noncanonical_master_target_address(string address)
        {
            SubjectDefinition definition = Value("value", 1, "before", "after");
            definition.MasterAddress = address;
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "value"));

            Assert.Equal(MergePlanBuildCodes.TargetInvalid, exception.Code);
        }

        [Fact]
        public void Rejects_missing_selected_source_occurrence()
        {
            SubjectDefinition definition = Value("value", 1, "before", "after");
            definition.EmptyCandidateProvenance = true;
            ReviewDecisionGeneration generation = Generation(
                new[] { definition },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                Build(generation, "value"));

            Assert.Equal(MergePlanBuildCodes.SourceInvalid, exception.Code);
        }

        [Fact]
        public void Rejects_duplicate_physical_target_for_same_aspect()
        {
            SubjectDefinition first = Value("first", 1, "a", "b");
            SubjectDefinition second = Value("second", 1, "c", "d");
            ReviewDecisionGeneration generation = Generation(
                new[] { first, second },
                new[]
                {
                    Decision("first", ReviewDecisionStatus.UseCandidate),
                    Decision("second", ReviewDecisionStatus.UseCandidate),
                });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                MergePlanBuilder.Build(
                    generation,
                    MergePlanRequest.Capture(generation, new[] { "first", "second" })));

            Assert.Equal(MergePlanBuildCodes.TargetAmbiguous, exception.Code);
        }

        [Fact]
        public void Rejects_one_logical_sheet_mapped_to_multiple_master_names()
        {
            SubjectDefinition first = Value("first", 1, "a", "b");
            SubjectDefinition second = Value("second", 2, "c", "d");
            second.MasterSheetName = "Other";
            ReviewDecisionGeneration generation = Generation(
                new[] { first, second },
                new[]
                {
                    Decision("first", ReviewDecisionStatus.UseCandidate),
                    Decision("second", ReviewDecisionStatus.UseCandidate),
                });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                MergePlanBuilder.Build(
                    generation,
                    MergePlanRequest.Capture(generation, new[] { "first", "second" })));

            Assert.Equal(MergePlanBuildCodes.TargetAmbiguous, exception.Code);
        }

        [Fact]
        public void Rejects_corrupt_decision_support_even_inside_generation_object()
        {
            SubjectDefinition definition = Value("value", 1, "before", "after");
            ReviewDecisionCatalog catalog = Catalog(new[] { definition });
            Observation selected = catalog.Observations.Single(
                item => item.SubjectId == "value" &&
                        item.RevisionId == CandidateRevision);
            ReviewDecisionGeneration generation = CorruptCandidateGeneration(
                catalog,
                selected,
                supportingContentHash: Hash("forged-support"));
            var request = new MergePlanRequest(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.CatalogId,
                generation.GenerationId,
                generation.Catalog.MasterRevisionId,
                new[] { new MergePlanSelection("value", generation.Events[0].EventId) });

            MergePlanBuildException exception = Assert.Throws<MergePlanBuildException>(() =>
                MergePlanBuilder.Build(generation, request));

            Assert.Equal(MergePlanBuildCodes.EvidenceInconsistent, exception.Code);
        }

        [Fact]
        public void Portable_plan_surface_excludes_source_reviewer_note_and_timestamp_fields()
        {
            ReviewDecisionGeneration generation = Generation(
                new[] { Value("value", 1, "before", "after") },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });

            MergePlan plan = Build(generation, "value");
            string[] publicStrings =
            {
                plan.ProjectId,
                plan.EvidenceGenerationId,
                plan.MappingGenerationId,
                plan.CatalogId,
                plan.DecisionGenerationId,
                plan.MasterRevisionId,
                plan.MasterByteContentId,
                plan.DecisionSetHash,
                plan.PlanHash,
                plan.Operations[0].OperationId,
                plan.Operations[0].SubjectId,
                plan.Operations[0].DecisionEventId,
                plan.Operations[0].SelectedObservationId,
                plan.Operations[0].SelectedRevisionId,
                plan.Operations[0].SelectedRevisionByteContentId,
                plan.Operations[0].MasterObservationId,
                plan.Operations[0].SelectedSource.LogicalSheetId,
                plan.Operations[0].SelectedSource.SheetName,
                plan.Operations[0].SelectedSource.Address,
                plan.Operations[0].Target.LogicalSheetId,
                plan.Operations[0].Target.SheetName,
                plan.Operations[0].Target.Address,
            };

            Assert.All(publicStrings, value =>
            {
                Assert.DoesNotContain(PrivateSourceLabel, value, StringComparison.Ordinal);
                Assert.DoesNotContain(PrivateReviewer, value, StringComparison.Ordinal);
                Assert.DoesNotContain("private-note", value, StringComparison.Ordinal);
                Assert.DoesNotContain(When.ToString("O"), value, StringComparison.Ordinal);
            });
            Assert.DoesNotContain(
                typeof(MergePlan).GetProperties(),
                property => property.Name.Contains("Path", StringComparison.Ordinal) ||
                            property.Name.Contains("Reviewer", StringComparison.Ordinal) ||
                            property.Name.Contains("Timestamp", StringComparison.Ordinal));
            Assert.DoesNotContain(
                typeof(MergePlanTarget).GetProperties(),
                property => property.Name.Contains("Source", StringComparison.Ordinal));
            Assert.DoesNotContain(
                typeof(MergePlanSourceOccurrence).GetProperties(),
                property => property.Name.Contains("Path", StringComparison.Ordinal) ||
                            property.Name.Contains("Label", StringComparison.Ordinal));
        }

        [Fact]
        public void Plan_hashes_bind_directly_into_prepared_recovery_record()
        {
            ReviewDecisionGeneration generation = Generation(
                new[] { Value("value", 1, "before", "after") },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });
            MergePlan plan = Build(generation, "value");
            var binding = new MergeRecoveryPreparedBinding(
                plan.ProjectId,
                "merge-1",
                plan.MasterRevisionId,
                Hash("project-state"),
                123,
                plan.MasterByteContentId,
                plan.EvidenceGenerationId,
                plan.MappingGenerationId,
                plan.DecisionGenerationId,
                plan.DecisionSetHash,
                plan.PlanHash,
                "destination-locator",
                "merged.xlsx");

            MergeRecoveryJournalRecord prepared =
                MergeRecoveryJournal.CreatePrepared(binding);

            Assert.Equal(plan.DecisionSetHash.ToLowerInvariant(), prepared.DecisionSetHash);
            Assert.Equal(plan.PlanHash.ToLowerInvariant(), prepared.MergePlanHash);
            Assert.Equal(plan.MasterByteContentId.ToLowerInvariant(), prepared.BaseSha256);
        }

        [Fact]
        public void Cancellation_stops_plan_construction()
        {
            ReviewDecisionGeneration generation = Generation(
                new[] { Value("value", 1, "before", "after") },
                new[] { Decision("value", ReviewDecisionStatus.UseCandidate) });
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() => MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(generation, new[] { "value" }),
                cancellation.Token));
        }

        private static MergePlan Build(
            ReviewDecisionGeneration generation,
            params string[] selectedSubjects) =>
            MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(generation, selectedSubjects));

        private static ReviewDecisionGeneration Generation(
            IReadOnlyList<SubjectDefinition> definitions,
            IReadOnlyList<ReviewDecisionSpec> decisions,
            bool reverseCatalogInputs = false)
        {
            ReviewDecisionGeneration generation = ReviewDecisionEngine.Create(
                Catalog(definitions, reverseCatalogInputs: reverseCatalogInputs));
            for (int index = 0; index < decisions.Count; index++)
                generation = Apply(generation, decisions[index], index);
            return generation;
        }

        private static ReviewDecisionGeneration Apply(
            ReviewDecisionGeneration generation,
            ReviewDecisionSpec specification,
            int index)
        {
            string selectedObservationId = "";
            if (specification.Status == ReviewDecisionStatus.UseCandidate ||
                specification.Status == ReviewDecisionStatus.UseMaster)
            {
                string revisionId = specification.Status == ReviewDecisionStatus.UseCandidate
                    ? CandidateRevision
                    : MasterRevision;
                selectedObservationId = generation.Catalog.Observations.Single(
                    item => item.SubjectId == specification.SubjectId &&
                            item.RevisionId == revisionId).ObservationId;
            }
            return ReviewDecisionEngine.Apply(
                generation,
                ReviewDecisionRequest.ForSubject(
                    generation.ProjectId,
                    generation.BaseGenerationId,
                    generation.MappingGenerationId,
                    generation.GenerationId,
                    specification.SubjectId,
                    specification.Status,
                    new ReviewerSnapshot(PrivateReviewer),
                    When.AddMinutes(index),
                    selectedObservationId,
                    runId: "run",
                    addInBuild: "test-build",
                    note: "private-note"));
        }

        private static ReviewDecisionGeneration CorruptCandidateGeneration(
            ReviewDecisionCatalog catalog,
            Observation selected,
            string supportingContentHash)
        {
            string eventId = Hash("corrupt-event-" + selected.SubjectId);
            var decision = new DecisionEvent(
                eventId,
                catalog.ProjectId,
                catalog.BaseGenerationId,
                catalog.MappingGenerationId,
                ReviewDecisionStatus.Unreviewed,
                ReviewDecisionStatus.UseCandidate,
                When,
                DecisionSelectionScope.SingleSubject,
                runId: "run",
                subjectId: selected.SubjectId,
                affectedSubjectIds: new[] { selected.SubjectId },
                selectedObservationId: selected.ObservationId,
                supportingRevisionIds: new[] { selected.RevisionId },
                supportingContentHashes: new[] { supportingContentHash },
                requestId: Hash("request-" + selected.SubjectId));
            return new ReviewDecisionGeneration(
                catalog,
                Hash("corrupt-generation-" + selected.SubjectId),
                "",
                catalog.SubjectIds.Select(subjectId =>
                    subjectId == selected.SubjectId
                        ? new ReviewDecisionState(
                            subjectId,
                            ReviewDecisionStatus.UseCandidate,
                            selected.ObservationId,
                            eventId)
                        : new ReviewDecisionState(
                            subjectId,
                            ReviewDecisionStatus.Unreviewed,
                            "",
                            "")),
                new[] { decision });
        }

        private static ReviewDecisionCatalog Catalog(
            IReadOnlyList<SubjectDefinition> definitions,
            bool evidenceBound = true,
            bool reverseCatalogInputs = false)
        {
            IEnumerable<ReviewDecisionSubject> subjects = definitions.Select(
                definition => new ReviewDecisionSubject(
                    definition.SubjectId,
                    definition.LogicalSubjectKey,
                    definition.Aspect,
                    definition.Policy));
            IEnumerable<Observation> observations = definitions.SelectMany(
                definition => new[]
                {
                    Observation(definition, CandidateRevision),
                    Observation(definition, MasterRevision),
                });
            if (reverseCatalogInputs)
            {
                subjects = subjects.Reverse();
                observations = observations.Reverse();
            }

            return evidenceBound
                ? new ReviewDecisionCatalog(
                    Evidence(),
                    MasterRevision,
                    subjects,
                    observations)
                : new ReviewDecisionCatalog(
                    "project",
                    "opaque-evidence",
                    "mapping",
                    MasterRevision,
                    subjects,
                    observations);
        }

        private static Observation Observation(
            SubjectDefinition definition,
            string revisionId)
        {
            bool master = revisionId == MasterRevision;
            ObservationExactType type = master
                ? definition.MasterType
                : definition.CandidateType;
            string text = master
                ? definition.MasterText
                : definition.CandidateText;
            string? seed = master
                ? definition.MasterContentSeed
                : definition.CandidateContentSeed;
            string contentHash = Hash(
                seed ?? definition.Aspect + ":" + type + ":" + text);
            CanonicalScalar? value = Scalar(type, text);
            string? formula = type == ObservationExactType.Formula ? text : null;
            ObservationProvenance provenance = master
                ? new ObservationProvenance(
                    PrivateSourceLabel,
                    definition.MasterSheetName,
                    definition.MasterAddress ?? A1(definition.Row, definition.Column),
                    definition.Row,
                    definition.Column,
                    definition.LogicalSheetId)
                : definition.EmptyCandidateProvenance
                    ? ObservationProvenance.Empty
                    : new ObservationProvenance(
                    @"C:\private\candidate.xlsx",
                    "Candidate Data",
                    A1(definition.Row, definition.Column),
                    definition.Row,
                    definition.Column,
                    definition.LogicalSheetId);
            return new Observation(
                ComparisonIds.Observation(
                    definition.SubjectId,
                    revisionId,
                    type,
                    contentHash),
                definition.SubjectId,
                revisionId,
                type,
                contentHash,
                formula,
                value,
                provenance: provenance);
        }

        private static CanonicalScalar? Scalar(
            ObservationExactType type,
            string text)
        {
            switch (type)
            {
                case ObservationExactType.Formula: return null;
                case ObservationExactType.Missing: return CanonicalScalar.Missing;
                case ObservationExactType.Blank: return CanonicalScalar.Blank;
                case ObservationExactType.DateSerial:
                    return new CanonicalScalar(type, text, ExcelDateSystem.Date1900);
                default:
                    return new CanonicalScalar(type, text);
            }
        }

        private static ReviewEvidenceGeneration Evidence() =>
            new ReviewEvidenceGeneration(
                "project",
                "mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity(
                        CandidateRevision,
                        0,
                        Hash("candidate-file")),
                    new ReviewEvidenceRevisionIdentity(
                        MasterRevision,
                        1,
                        Hash("master-file")),
                });

        private static SubjectDefinition Value(
            string subjectId,
            int row,
            string master,
            string candidate) =>
            new SubjectDefinition
            {
                SubjectId = subjectId,
                LogicalSubjectKey = "logical-key-" + subjectId,
                Aspect = ComparisonAspect.Value,
                Policy = ReviewDecisionSubjectPolicy.SelectableAndMergeable,
                Row = row,
                Column = 1,
                MasterType = ObservationExactType.String,
                CandidateType = ObservationExactType.String,
                MasterText = master,
                CandidateText = candidate,
            };

        private static SubjectDefinition Formula(
            string subjectId,
            int row,
            string master,
            string candidate) =>
            new SubjectDefinition
            {
                SubjectId = subjectId,
                LogicalSubjectKey = "logical-key-" + subjectId,
                Aspect = ComparisonAspect.Formula,
                Policy = ReviewDecisionSubjectPolicy.SelectableAndMergeable,
                Row = row,
                Column = 1,
                MasterType = ObservationExactType.Formula,
                CandidateType = ObservationExactType.Formula,
                MasterText = master,
                CandidateText = candidate,
            };

        private static SubjectDefinition Structure(string subjectId, int row) =>
            new SubjectDefinition
            {
                SubjectId = subjectId,
                LogicalSubjectKey = "logical-key-" + subjectId,
                Aspect = ComparisonAspect.Structure,
                Policy = ReviewDecisionSubjectPolicy.ReviewOnly,
                Row = row,
                Column = 1,
                MasterType = ObservationExactType.String,
                CandidateType = ObservationExactType.String,
                MasterText = "structure-before",
                CandidateText = "structure-after",
            };

        private static ReviewDecisionSpec Decision(
            string subjectId,
            ReviewDecisionStatus status) =>
            new ReviewDecisionSpec(subjectId, status);

        private static string A1(int row, int column)
        {
            var result = new StringBuilder();
            int remaining = column;
            while (remaining > 0)
            {
                remaining--;
                result.Insert(0, (char)('A' + remaining % 26));
                remaining /= 26;
            }
            return result.Append(row).ToString();
        }

        private static string Hash(string value) =>
            StableId.ForBytes(Encoding.UTF8.GetBytes(value));

        private sealed class SubjectDefinition
        {
            public string SubjectId { get; set; } = "";
            public string LogicalSubjectKey { get; set; } = "";
            public ComparisonAspect Aspect { get; set; }
            public ReviewDecisionSubjectPolicy Policy { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public string LogicalSheetId { get; set; } = "logical-sheet";
            public string MasterSheetName { get; set; } = "Data";
            public string? MasterAddress { get; set; }
            public ObservationExactType MasterType { get; set; }
            public ObservationExactType CandidateType { get; set; }
            public string MasterText { get; set; } = "";
            public string CandidateText { get; set; } = "";
            public string? MasterContentSeed { get; set; }
            public string? CandidateContentSeed { get; set; }
            public bool EmptyCandidateProvenance { get; set; }
        }

        private sealed class ReviewDecisionSpec
        {
            internal ReviewDecisionSpec(
                string subjectId,
                ReviewDecisionStatus status)
            {
                SubjectId = subjectId;
                Status = status;
            }

            internal string SubjectId { get; }
            internal ReviewDecisionStatus Status { get; }
        }
    }
}
