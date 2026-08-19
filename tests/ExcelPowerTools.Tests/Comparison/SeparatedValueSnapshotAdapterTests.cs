using System.Linq;
using System.IO;
using System.Text;
using ExcelPowerTools.Core.Comparison.Adapters;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class SeparatedValueSnapshotAdapterTests
    {
        [Fact]
        public void Every_field_is_retained_as_exact_text_including_comments_and_formula_looking_values()
        {
            byte[] bytes = Utf8("# comment;\r\nName;Expression;Empty\r\nrow1;=1+1;\r\n");

            SeparatedValueSnapshot result =
                SeparatedValueSnapshotAdapter.FromBytes("revision-1", bytes, "input.csv");

            Assert.Equal(3, result.RecordCount);
            Assert.Equal(3, result.MaxFieldCount);
            Assert.Equal(StableId.ComputeByteContentId(bytes), result.ByteContentId);

            CanonicalCellState formulaLooking = result.Snapshot.Cells[
                SeparatedValueSnapshotAdapter.PositionalSubjectKey(3, 2)];
            Assert.False(formulaLooking.HasFormula);
            Assert.Equal(CanonicalValueKind.Text, formulaLooking.ConstantValue.Kind);
            Assert.Equal("=1+1", formulaLooking.ConstantValue.ExactText);

            CanonicalCellState explicitEmpty = result.Snapshot.Cells[
                SeparatedValueSnapshotAdapter.PositionalSubjectKey(3, 3)];
            Assert.Equal(CanonicalValueKind.Text, explicitEmpty.ConstantValue.Kind);
            Assert.Equal("", explicitEmpty.ConstantValue.ExactText);
        }

        [Fact]
        public void Header_detector_output_is_only_an_untrusted_candidate()
        {
            byte[] bytes = Utf8("Name;Amount\r\na;1\r\nb;2\r\n");

            SeparatedValueSnapshot result =
                SeparatedValueSnapshotAdapter.FromBytes("revision-1", bytes, "input.csv");

            Assert.Equal(0, result.HeaderCandidateRecordIndex);
            Assert.Equal(3, result.RecordCount); // header candidate remains ordinary snapshot data
        }

        [Fact]
        public void Tsv_path_hint_preserves_tab_for_single_column_input()
        {
            SeparatedValueSnapshot result =
                SeparatedValueSnapshotAdapter.FromBytes("revision-1", Utf8("only one column\n"), "input.tsv");

            Assert.Equal('\t', result.Format.Delimiter);
        }

        [Fact]
        public void Positional_snapshot_feeds_the_exact_engine_deterministically()
        {
            SeparatedValueSnapshot before =
                SeparatedValueSnapshotAdapter.FromBytes("before", Utf8("a;b\r\n1;2\r\n"), "before.csv");
            SeparatedValueSnapshot after =
                SeparatedValueSnapshotAdapter.FromBytes("after", Utf8("a;b\r\n1;9\r\n"), "after.csv");

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));
            Assert.Equal(SeparatedValueSnapshotAdapter.PositionalSubjectKey(2, 2), change.SubjectKey);
            Assert.Equal(CellChangeAspect.Value, change.Aspect);
            Assert.Equal("2", change.BeforeValue!.ExactText);
            Assert.Equal("9", change.AfterValue!.ExactText);
        }

        [Fact]
        public void Missing_field_and_explicit_empty_field_remain_distinct()
        {
            SeparatedValueSnapshot before =
                SeparatedValueSnapshotAdapter.FromBytes("before", Utf8("a\r\n"), "before.csv");
            SeparatedValueSnapshot after =
                SeparatedValueSnapshotAdapter.FromBytes("after", Utf8("a;\r\n"), "after.csv");

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));
            Assert.Equal(CellChangeKind.Added, change.Kind);
            Assert.Equal(CanonicalValueKind.Text, change.AfterValue!.Kind);
            Assert.Equal("", change.AfterValue.ExactText);
        }

        [Fact]
        public void Subject_keys_sort_rows_and_columns_ordinally()
        {
            string[] keys =
            {
                SeparatedValueSnapshotAdapter.PositionalSubjectKey(10, 2),
                SeparatedValueSnapshotAdapter.PositionalSubjectKey(2, 10),
                SeparatedValueSnapshotAdapter.PositionalSubjectKey(2, 2),
            };

            Assert.Equal(
                new[]
                {
                    SeparatedValueSnapshotAdapter.PositionalSubjectKey(2, 2),
                    SeparatedValueSnapshotAdapter.PositionalSubjectKey(2, 10),
                    SeparatedValueSnapshotAdapter.PositionalSubjectKey(10, 2),
                },
                keys.OrderBy(key => key, System.StringComparer.Ordinal));
        }

        [Fact]
        public void Malformed_bom_declared_unicode_is_rejected_instead_of_replacement_decoded()
        {
            byte[] malformedUtf8 = { 0xEF, 0xBB, 0xBF, 0xC3, 0x28 };
            byte[] malformedUtf16 = { 0xFF, 0xFE, 0x00, 0xD8 };

            Assert.Throws<InvalidDataException>(() =>
                SeparatedValueSnapshotAdapter.FromBytes("bad-utf8", malformedUtf8, "bad.csv"));
            Assert.Throws<InvalidDataException>(() =>
                SeparatedValueSnapshotAdapter.FromBytes("bad-utf16", malformedUtf16, "bad.csv"));
        }

        private static byte[] Utf8(string text) => new UTF8Encoding(false).GetBytes(text);
    }
}
