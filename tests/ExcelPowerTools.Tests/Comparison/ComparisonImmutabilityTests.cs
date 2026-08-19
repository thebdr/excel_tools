using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Revision;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ComparisonImmutabilityTests
    {
        [Fact]
        public void Revision_snapshot_does_not_expose_its_dictionary()
        {
            var snapshot = new RevisionSnapshot(
                "revision",
                new[] { CanonicalCellState.Constant("A1", CanonicalValue.Text("before")) });

            Assert.False(snapshot.Cells is Dictionary<string, CanonicalCellState>);
            Assert.Throws<NotSupportedException>(() =>
                ((IDictionary<string, CanonicalCellState>)snapshot.Cells).Add(
                    "A2",
                    CanonicalCellState.Constant("A2", CanonicalValue.Text("after"))));
            Assert.Single(snapshot.Cells);
        }

        [Fact]
        public void Exact_comparison_does_not_expose_its_result_list()
        {
            var before = new RevisionSnapshot(
                "before",
                new[] { CanonicalCellState.Constant("A1", CanonicalValue.Text("before")) });
            var after = new RevisionSnapshot(
                "after",
                new[] { CanonicalCellState.Constant("A1", CanonicalValue.Text("after")) });
            IReadOnlyList<AtomicCellChange> changes = ExactComparisonEngine.Compare(before, after);

            Assert.False(changes is List<AtomicCellChange>);
            Assert.Throws<NotSupportedException>(() => ((IList<AtomicCellChange>)changes).Clear());
            Assert.Single(changes);
        }

        [Fact]
        public void Character_diff_does_not_expose_its_run_list()
        {
            CharacterDiffResult result = CharacterDiff.Compare("before", "after");

            Assert.False(result.Runs is List<CharacterDiffRun>);
            Assert.Throws<NotSupportedException>(() => ((IList<CharacterDiffRun>)result.Runs).Clear());
            Assert.NotEmpty(result.Runs);
        }

        [Fact]
        public void Revision_inference_does_not_expose_components_or_candidates()
        {
            var number = new RevisionNumber(2, 1);
            RevisionLabelInferenceResult result = RevisionLabelInference.Infer("artifact v2.1.xlsx");

            Assert.False(number.Components is int[]);
            Assert.Throws<NotSupportedException>(() => ((IList<int>)number.Components)[0] = 9);
            Assert.Equal("2.1", number.ToString());
            Assert.False(result.Candidates is List<RevisionTokenCandidate>);
            Assert.Throws<NotSupportedException>(() =>
                ((IList<RevisionTokenCandidate>)result.Candidates).Clear());
            Assert.True(result.HasStrongCandidate);
        }
    }
}
