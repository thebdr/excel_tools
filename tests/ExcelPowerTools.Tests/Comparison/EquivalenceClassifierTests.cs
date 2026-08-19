using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class EquivalenceClassifierTests
    {
        [Fact]
        public void Case_only_difference_is_labelled_without_removing_exact_evidence()
        {
            AtomicCellChange change = ValueChange(CanonicalValue.Text("Alpha"), CanonicalValue.Text("alpha"));

            EquivalenceTag tags = EquivalenceClassifier.Classify(change);

            Assert.True(tags.HasFlag(EquivalenceTag.CaseOnly));
            Assert.Equal("Alpha", change.BeforeText);
            Assert.Equal("alpha", change.AfterText);
        }

        [Fact]
        public void Whitespace_and_punctuation_are_independent_labels()
        {
            EquivalenceTag whitespace = EquivalenceClassifier.Classify(
                ValueChange(CanonicalValue.Text("A B"), CanonicalValue.Text("AB")));
            EquivalenceTag punctuation = EquivalenceClassifier.Classify(
                ValueChange(CanonicalValue.Text("A-B"), CanonicalValue.Text("AB")));

            Assert.True(whitespace.HasFlag(EquivalenceTag.WhitespaceOnly));
            Assert.False(whitespace.HasFlag(EquivalenceTag.PunctuationOnly));
            Assert.True(punctuation.HasFlag(EquivalenceTag.PunctuationOnly));
            Assert.False(punctuation.HasFlag(EquivalenceTag.WhitespaceOnly));
        }

        [Fact]
        public void Numeric_text_and_number_can_be_quietly_equivalent_but_remain_typed_different()
        {
            AtomicCellChange change = ValueChange(CanonicalValue.Text("001.0"), CanonicalValue.Number(1));

            EquivalenceTag tags = EquivalenceClassifier.Classify(change);

            Assert.True(tags.HasFlag(EquivalenceTag.NumericallyEquivalent));
            Assert.NotEqual(change.BeforeValue!.Kind, change.AfterValue!.Kind);
        }

        [Fact]
        public void Blank_and_explicit_empty_text_are_labelled()
        {
            EquivalenceTag tags = EquivalenceClassifier.Classify(
                ValueChange(CanonicalValue.Blank, CanonicalValue.Text("")));

            Assert.True(tags.HasFlag(EquivalenceTag.BlankAndEmptyText));
        }

        [Fact]
        public void Substantive_text_change_has_no_equivalence_tags()
        {
            Assert.Equal(
                EquivalenceTag.None,
                EquivalenceClassifier.Classify(
                    ValueChange(CanonicalValue.Text("north"), CanonicalValue.Text("south"))));
        }

        private static AtomicCellChange ValueChange(CanonicalValue before, CanonicalValue after)
        {
            var beforeSnapshot = new RevisionSnapshot(
                "before",
                new[] { CanonicalCellState.Constant("A1", before) });
            var afterSnapshot = new RevisionSnapshot(
                "after",
                new[] { CanonicalCellState.Constant("A1", after) });
            return Assert.Single(ExactComparisonEngine.Compare(beforeSnapshot, afterSnapshot));
        }
    }
}
