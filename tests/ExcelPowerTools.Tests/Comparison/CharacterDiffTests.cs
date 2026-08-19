using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Comparison.Diff;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class CharacterDiffTests
    {
        [Fact]
        public void Replacement_has_equal_removed_added_equal_runs()
        {
            CharacterDiffResult result = CharacterDiff.Compare("abc", "axc");

            Assert.Collection(
                result.Runs,
                run => AssertRun(run, CharacterDiffKind.Equal, "a"),
                run => AssertRun(run, CharacterDiffKind.Removed, "b"),
                run => AssertRun(run, CharacterDiffKind.Added, "x"),
                run => AssertRun(run, CharacterDiffKind.Equal, "c"));
            Assert.False(result.IsSimplified);
            AssertReconstructs(result);
        }

        [Fact]
        public void Insertions_and_deletions_reconstruct_both_exact_inputs()
        {
            CharacterDiffResult result = CharacterDiff.Compare("=SUM(A1:A9)", "=SUM(A1:A10) + 1");

            AssertReconstructs(result);
            Assert.Contains(result.Runs, r => r.Kind == CharacterDiffKind.Added);
            Assert.Contains(result.Runs, r => r.Kind == CharacterDiffKind.Removed);
        }

        [Fact]
        public void Unicode_combining_sequence_is_not_split_inside_a_text_element()
        {
            CharacterDiffResult result = CharacterDiff.Compare("Cafe\u0301", "Café");

            AssertReconstructs(result);
            CharacterDiffRun removed = Assert.Single(result.Runs, r => r.Kind == CharacterDiffKind.Removed);
            Assert.Equal("e\u0301", removed.Text);
        }

        [Fact]
        public void Visible_whitespace_marks_spaces_tabs_and_line_endings()
        {
            Assert.Equal("a·⇥b␍␊c", CharacterDiff.MakeWhitespaceVisible("a \tb\r\nc"));
        }

        [Fact]
        public void Large_middle_falls_back_to_exact_simplified_blocks()
        {
            string before = "prefix-" + new string('a', 50) + "-suffix";
            string after = "prefix-" + new string('b', 50) + "-suffix";

            CharacterDiffResult result = CharacterDiff.Compare(before, after, maxMatrixCells: 10);

            Assert.True(result.IsSimplified);
            AssertReconstructs(result);
            Assert.Contains(result.Runs, r => r.Kind == CharacterDiffKind.Removed && r.Text.Length == 50);
            Assert.Contains(result.Runs, r => r.Kind == CharacterDiffKind.Added && r.Text.Length == 50);
        }

        [Fact]
        public void Null_inputs_are_exactly_equivalent_to_empty_text()
        {
            CharacterDiffResult result = CharacterDiff.Compare(null!, "");

            Assert.Empty(result.Runs);
            Assert.Equal("", result.BeforeText);
            Assert.Equal("", result.AfterText);
        }

        private static void AssertRun(CharacterDiffRun run, CharacterDiffKind kind, string text)
        {
            Assert.Equal(kind, run.Kind);
            Assert.Equal(text, run.Text);
        }

        private static void AssertReconstructs(CharacterDiffResult result)
        {
            var before = new StringBuilder();
            var after = new StringBuilder();
            foreach (CharacterDiffRun run in result.Runs)
            {
                if (run.Kind != CharacterDiffKind.Added) before.Append(run.Text);
                if (run.Kind != CharacterDiffKind.Removed) after.Append(run.Text);
            }

            Assert.Equal(result.BeforeText, before.ToString());
            Assert.Equal(result.AfterText, after.ToString());
        }
    }
}
