using System;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class ExactComparisonEngineTests
    {
        [Fact]
        public void Identical_snapshots_have_no_atomic_changes()
        {
            RevisionSnapshot before = Snapshot("before", Constant("A1", "same"), Formula("B1", "=A1", "same"));
            RevisionSnapshot after = Snapshot("after", Constant("A1", "same"), Formula("B1", "=A1", "same"));

            Assert.Empty(ExactComparisonEngine.Compare(before, after));
        }

        [Fact]
        public void Typed_value_change_is_one_primary_atomic_change()
        {
            RevisionSnapshot before = Snapshot("before", Constant("A1", "old"));
            RevisionSnapshot after = Snapshot("after", Constant("A1", "new"));

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before, after));
            Assert.Equal(CellChangeAspect.Value, change.Aspect);
            Assert.Equal(CellChangeKind.Modified, change.Kind);
            Assert.Equal("old", change.BeforeText);
            Assert.Equal("new", change.AfterText);
            Assert.False(change.IsSecondary);
        }

        [Fact]
        public void Same_formula_with_different_cache_is_secondary_only()
        {
            RevisionSnapshot before = Snapshot("before", Formula("A1", "=1+1", "2"));
            RevisionSnapshot after = Snapshot("after", Formula("A1", "=1+1", "3"));

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before, after));
            Assert.Equal(CellChangeAspect.CachedResult, change.Aspect);
            Assert.True(change.IsSecondary);
        }

        [Fact]
        public void Formula_change_is_primary_and_suppresses_cache_noise_by_default()
        {
            RevisionSnapshot before = Snapshot("before", Formula("A1", "=1+1", "2"));
            RevisionSnapshot after = Snapshot("after", Formula("A1", "=1+2", "3"));

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before, after));
            Assert.Equal(CellChangeAspect.Formula, change.Aspect);
            Assert.Equal("=1+1", change.BeforeText);
            Assert.Equal("=1+2", change.AfterText);
        }

        [Fact]
        public void Cache_change_can_be_retained_when_formula_also_changed()
        {
            RevisionSnapshot before = Snapshot("before", Formula("A1", "=1+1", "2"));
            RevisionSnapshot after = Snapshot("after", Formula("A1", "=1+2", "3"));
            var options = new ExactComparisonOptions { CompareCachedResultsWhenFormulaChanges = true };

            var changes = ExactComparisonEngine.Compare(before, after, options);

            Assert.Collection(
                changes,
                change => Assert.Equal(CellChangeAspect.Formula, change.Aspect),
                change => Assert.Equal(CellChangeAspect.CachedResult, change.Aspect));
        }

        [Fact]
        public void Formula_to_constant_transition_keeps_both_atomic_aspects()
        {
            RevisionSnapshot before = Snapshot("before", Formula("A1", "=2+2", "4"));
            RevisionSnapshot after = Snapshot("after", Constant("A1", "four"));

            var changes = ExactComparisonEngine.Compare(before, after);

            Assert.Collection(
                changes,
                formula =>
                {
                    Assert.Equal(CellChangeAspect.Formula, formula.Aspect);
                    Assert.Equal(CellChangeKind.Removed, formula.Kind);
                },
                value =>
                {
                    Assert.Equal(CellChangeAspect.Value, value.Aspect);
                    Assert.Equal(CellChangeKind.Added, value.Kind);
                });
        }

        [Fact]
        public void Missing_nonblank_cell_is_reported_as_content_not_presence_noise()
        {
            RevisionSnapshot before = Snapshot("before");
            RevisionSnapshot after = Snapshot("after", Constant("A1", "added"));

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before, after));
            Assert.Equal(CellChangeAspect.Value, change.Aspect);
            Assert.Equal(CellChangeKind.Added, change.Kind);
        }

        [Fact]
        public void Explicit_blank_versus_missing_cell_is_secondary_presence_evidence()
        {
            RevisionSnapshot before = Snapshot("before");
            RevisionSnapshot after = Snapshot("after", CanonicalCellState.Blank("A1"));

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before, after));
            Assert.Equal(CellChangeAspect.Presence, change.Aspect);
            Assert.True(change.IsSecondary);
        }

        [Fact]
        public void Atomic_changes_are_deterministically_sorted_by_subject_then_aspect()
        {
            RevisionSnapshot before = Snapshot("before", Constant("Z9", "old"), Formula("A1", "=1", "1"));
            RevisionSnapshot after = Snapshot("after", Constant("Z9", "new"), Constant("A1", "x"));

            var changes = ExactComparisonEngine.Compare(before, after);

            Assert.Equal(new[] { "A1", "A1", "Z9" }, changes.Select(c => c.SubjectKey));
            Assert.Equal(
                new[] { CellChangeAspect.Formula, CellChangeAspect.Value, CellChangeAspect.Value },
                changes.Select(c => c.Aspect));
        }

        [Fact]
        public void Cancellation_is_observed_before_scanning_subjects()
        {
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                ExactComparisonEngine.Compare(
                    Snapshot("before", Constant("A1", "a")),
                    Snapshot("after", Constant("A1", "b")),
                    cancellationToken: cancellation.Token));
        }

        private static RevisionSnapshot Snapshot(string id, params CanonicalCellState[] cells) =>
            new RevisionSnapshot(id, cells);

        private static CanonicalCellState Constant(string subject, string value) =>
            CanonicalCellState.Constant(subject, CanonicalValue.Text(value));

        private static CanonicalCellState Formula(string subject, string formula, string cached) =>
            CanonicalCellState.Formula(subject, formula, CanonicalValue.Text(cached));
    }
}
