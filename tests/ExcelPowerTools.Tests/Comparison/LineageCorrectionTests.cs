using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Lineage;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class LineageCorrectionTests
    {
        [Fact]
        public void Corrections_append_events_and_create_new_generations_without_mutating_history()
        {
            LineageMappingGeneration initial = Initial();
            var reviewer = new ReviewerSnapshot("Alice Reviewer", "Alice", new[] { "confirmed", "office" });
            DateTimeOffset at = new DateTimeOffset(2026, 2, 1, 10, 30, 0, TimeSpan.Zero);

            LineageMappingGeneration labelled = LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetLabel("r2", "Reviewed revision 2", reviewer, at, "clarify label"));
            LineageMappingGeneration mastered = LineageCorrectionEngine.Apply(
                labelled,
                LineageCorrectionRequest.SetMaster("r2", reviewer, at.AddMinutes(1), "selected after review"));
            LineageMappingGeneration rebased = LineageCorrectionEngine.Apply(
                mastered,
                LineageCorrectionRequest.SetTrustedBaseline("r2", reviewer, at.AddMinutes(2)));

            Assert.Equal("Plant_r2", initial.GetLabel("r2"));
            Assert.Equal("r3", initial.MasterRevisionId);
            Assert.Equal("r1", initial.TrustedBaselineRevisionId);
            Assert.Empty(initial.Events);

            Assert.Equal("Reviewed revision 2", labelled.GetLabel("r2"));
            Assert.Equal("r3", labelled.MasterRevisionId);
            Assert.Single(labelled.Events);
            Assert.Equal(initial.GenerationId, labelled.PriorGenerationId);

            Assert.Equal("r2", mastered.MasterRevisionId);
            Assert.Equal(2, mastered.Events.Count);
            Assert.Equal(labelled.GenerationId, mastered.Events[1].PriorGenerationId);
            Assert.Equal(mastered.GenerationId, mastered.Events[1].NewGenerationId);

            Assert.Equal("r2", rebased.TrustedBaselineRevisionId);
            Assert.Equal(3, rebased.Events.Count);
            Assert.Equal("Alice Reviewer", rebased.Events[2].Reviewer.ConfirmedDisplayName);
            Assert.Equal(at.AddMinutes(2), rebased.Events[2].CreatedUtc);
        }

        [Fact]
        public void Parent_corrections_enforce_known_earlier_occurrences_and_prevent_cycles()
        {
            LineageMappingGeneration initial = Initial();
            DateTimeOffset at = DateTimeOffset.Parse("2026-02-01T00:00:00Z");

            LineageMappingGeneration changed = LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetParent("r3", "r1", ReviewerSnapshot.Empty, at));

            Assert.Equal("r1", changed.GetParent("r3"));
            Assert.Equal("r2", initial.GetParent("r3"));
            Assert.Throws<InvalidOperationException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetParent("r1", "r3", ReviewerSnapshot.Empty, at)));
            Assert.Throws<ArgumentException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetParent("r3", "missing", ReviewerSnapshot.Empty, at)));
        }

        [Fact]
        public void Reorder_requires_a_complete_permutation_and_preserves_parent_before_child_invariant()
        {
            LineageMappingGeneration initial = Initial();
            DateTimeOffset at = DateTimeOffset.Parse("2026-02-01T00:00:00Z");

            Assert.Throws<ArgumentException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.ReplaceOrder(
                    new[] { "r1", "r2" }, ReviewerSnapshot.Empty, at)));
            Assert.Throws<ArgumentException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.ReplaceOrder(
                    new[] { "r1", "r1", "r3" }, ReviewerSnapshot.Empty, at)));
            Assert.Throws<InvalidOperationException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.ReplaceOrder(
                    new[] { "r1", "r3", "r2" }, ReviewerSnapshot.Empty, at)));
        }

        [Fact]
        public void Parents_can_be_explicitly_cleared_before_a_reorder()
        {
            LineageMappingGeneration initial = Initial();
            DateTimeOffset at = DateTimeOffset.Parse("2026-02-01T00:00:00Z");
            LineageMappingGeneration clearedR2 = LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.ClearParent("r2", ReviewerSnapshot.Empty, at));
            LineageMappingGeneration clearedR3 = LineageCorrectionEngine.Apply(
                clearedR2,
                LineageCorrectionRequest.ClearParent("r3", ReviewerSnapshot.Empty, at.AddSeconds(1)));
            LineageMappingGeneration reordered = LineageCorrectionEngine.Apply(
                clearedR3,
                LineageCorrectionRequest.ReplaceOrder(
                    new[] { "r3", "r1", "r2" },
                    ReviewerSnapshot.Empty,
                    at.AddSeconds(2),
                    "manual content-based order"));

            Assert.Equal(new[] { "r3", "r1", "r2" }, reordered.OrderedRevisionIds);
            LineageCorrectionEvent reorderEvent = reordered.Events.Last();
            Assert.Equal(LineageCorrectionKind.ReplaceOrder, reorderEvent.Kind);
            Assert.Equal(new[] { "r1", "r2", "r3" }, reorderEvent.PreviousOrder);
            Assert.Equal(new[] { "r3", "r1", "r2" }, reorderEvent.NewOrder);
            Assert.Equal(clearedR3.GenerationId, reorderEvent.PriorGenerationId);
        }

        [Fact]
        public void Event_and_generation_ids_are_deterministic_for_the_same_correction()
        {
            LineageMappingGeneration initial = Initial();
            var reviewer = new ReviewerSnapshot("Reviewer");
            DateTimeOffset at = DateTimeOffset.Parse("2026-02-01T00:00:00Z");
            LineageCorrectionRequest request = LineageCorrectionRequest.SetLabel(
                "r2", "R2 confirmed", reviewer, at, "same event");

            LineageMappingGeneration first = LineageCorrectionEngine.Apply(initial, request);
            LineageMappingGeneration second = LineageCorrectionEngine.Apply(initial, request);

            Assert.Equal(first.GenerationId, second.GenerationId);
            Assert.Equal(first.Events.Single().EventId, second.Events.Single().EventId);
            Assert.Equal(initial.GenerationId, first.Events.Single().PriorGenerationId);
            Assert.Equal(first.GenerationId, first.Events.Single().NewGenerationId);
        }

        [Fact]
        public void No_op_corrections_are_rejected_instead_of_polluting_the_audit_chain()
        {
            LineageMappingGeneration initial = Initial();
            DateTimeOffset at = DateTimeOffset.Parse("2026-02-01T00:00:00Z");

            Assert.Throws<InvalidOperationException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetMaster("r3", ReviewerSnapshot.Empty, at)));
            Assert.Throws<InvalidOperationException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetLabel("r2", "Plant_r2", ReviewerSnapshot.Empty, at)));
            Assert.Throws<InvalidOperationException>(() => LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.ReplaceOrder(
                    initial.OrderedRevisionIds, ReviewerSnapshot.Empty, at)));
        }

        [Fact]
        public void Mapping_snapshots_and_events_do_not_expose_mutable_collections()
        {
            LineageMappingGeneration initial = Initial();
            LineageMappingGeneration changed = LineageCorrectionEngine.Apply(
                initial,
                LineageCorrectionRequest.SetLabel(
                    "r2",
                    "R2 confirmed",
                    ReviewerSnapshot.Empty,
                    DateTimeOffset.Parse("2026-02-01T00:00:00Z")));

            Assert.Throws<NotSupportedException>(() =>
                ((IList<string>)changed.OrderedRevisionIds).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IDictionary<string, string>)changed.ParentByRevisionId)["r2"] = "tampered");
            Assert.Throws<NotSupportedException>(() =>
                ((IList<LineageCorrectionEvent>)changed.Events).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<string>)changed.Events.Single().Reviewer.IdentitySourceLabels).Clear());
        }

        private static LineageMappingGeneration Initial()
        {
            LineageInferenceResult inference = MultiRevisionLineageInference.Infer(
                new[]
                {
                    Evidence("r1", "Plant_r1.xlsx", 'A', 1),
                    Evidence("r2", "Plant_r2.xlsx", 'B', 2),
                    Evidence("r3", "Plant_r3.xlsx", 'C', 3),
                },
                new LineageExplicitSelections("r3", "r1"));
            return LineageCorrectionEngine.FromInference(inference);
        }

        private static LineageRevisionEvidence Evidence(
            string revisionId,
            string fileName,
            char hashCharacter,
            int day) =>
            new LineageRevisionEvidence(
                revisionId,
                fileName,
                new string(hashCharacter, 64),
                modifiedUtc: new DateTimeOffset(2026, 1, day, 12, 0, 0, TimeSpan.Zero));
    }
}
