using System.Collections.Generic;
using System.Text;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvDiffTests
    {
        private static List<CsvRecord> Rows(params string[][] rows)
        {
            var list = new List<CsvRecord>();
            foreach (var r in rows) list.Add(CsvRecord.Data(r));
            return list;
        }

        [Fact]
        public void Identical_sets_have_no_changes()
        {
            var a = Rows(new[] { "x", "y" }, new[] { "1", "2" });
            var b = Rows(new[] { "x", "y" }, new[] { "1", "2" });
            var s = CsvDiff.Summarize(a, b);
            Assert.False(s.HasChanges);
            Assert.Equal("no value changes", s.Describe());
        }

        [Fact]
        public void Changed_cell_counted_with_first_position()
        {
            var a = Rows(new[] { "x", "y" }, new[] { "1", "2" });
            var b = Rows(new[] { "x", "y" }, new[] { "1", "9" });
            var s = CsvDiff.Summarize(a, b);
            Assert.Equal(1, s.ChangedCells);
            Assert.Equal(2, s.FirstChangeRow);
            Assert.Equal(2, s.FirstChangeColumn);
        }

        [Fact]
        public void Added_and_removed_rows_counted()
        {
            var s1 = CsvDiff.Summarize(Rows(new[] { "a" }), Rows(new[] { "a" }, new[] { "b" }, new[] { "c" }));
            Assert.Equal(2, s1.AddedRows);
            Assert.Equal(0, s1.RemovedRows);

            var s2 = CsvDiff.Summarize(Rows(new[] { "a" }, new[] { "b" }), Rows(new[] { "a" }));
            Assert.Equal(1, s2.RemovedRows);
        }

        [Fact]
        public void Width_difference_counts_as_changed_cells()
        {
            var s = CsvDiff.Summarize(Rows(new[] { "a", "b", "c" }), Rows(new[] { "a", "b" }));
            Assert.Equal(1, s.ChangedCells);   // the vanished "c" vs ""
        }

        [Fact]
        public void Missing_field_differs_from_an_explicit_empty_field()
        {
            var s = CsvDiff.Summarize(Rows(new[] { "a" }), Rows(new[] { "a", "" }));
            Assert.True(s.HasChanges);
            Assert.Equal(1, s.ChangedCells);
            Assert.Equal(2, s.FirstChangeColumn);
        }

        [Fact]
        public void VerifyRoundTrip_passes_for_written_bytes()
        {
            var fmt = new CsvFormat { Delimiter = ';' };
            var doc = new CsvDocument
            {
                Format = fmt,
                Records = Rows(new[] { "a", "b;c", "he said \"hi\"" }, new[] { "#x", "line1\nline2", "007" }),
            };
            Assert.Null(CsvDiff.VerifyRoundTrip(doc.Save(), fmt, doc.Records));
        }

        [Fact]
        public void VerifyRoundTrip_catches_corrupted_bytes()
        {
            var fmt = new CsvFormat { Delimiter = ';' };
            var doc = new CsvDocument { Format = fmt, Records = Rows(new[] { "abc", "def" }) };

            byte[] bytes = doc.Save();
            bytes[1] = (byte)'X';   // corrupt one byte ("abc" -> "aXc")

            string mismatch = CsvDiff.VerifyRoundTrip(bytes, fmt, doc.Records);
            Assert.NotNull(mismatch);
            Assert.Contains("row 1", mismatch);
        }

        [Fact]
        public void VerifyRoundTrip_catches_a_lost_trailing_empty_field()
        {
            var fmt = new CsvFormat { Delimiter = ';', HasTrailingNewline = true };
            byte[] missingField = new UTF8Encoding(false).GetBytes("a\r\n");

            string mismatch = CsvDiff.VerifyRoundTrip(missingField, fmt, Rows(new[] { "a", "" }));

            Assert.NotNull(mismatch);
            Assert.Contains("field count", mismatch);
        }

        [Fact]
        public void VerifyRoundTrip_passes_for_real_file_shape()
        {
            // The production shape: padded '#' count row, '#' header, '@' data rows.
            byte[] input = new UTF8Encoding(false).GetBytes(
                "#;2;;;;\r\n#;Name;Type;Node;Offset;Bit\r\n@;AUTO1_Z01;1;80;3;0\r\n@;AUTO1_Z02;5;8;0;0\r\n");
            var doc = CsvFile.Load(input);
            Assert.Null(CsvDiff.VerifyRoundTrip(doc.Save(), doc.Format, doc.Records));
        }
    }
}
