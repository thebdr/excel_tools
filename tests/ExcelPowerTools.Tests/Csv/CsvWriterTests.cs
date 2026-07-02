using System.Collections.Generic;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvWriterTests
    {
        private static List<CsvRecord> Rows(params string[][] rows)
        {
            var list = new List<CsvRecord>();
            foreach (var r in rows) list.Add(CsvRecord.Data(r));
            return list;
        }

        [Fact]
        public void Minimal_quoting_only_when_needed()
        {
            var fmt = new CsvFormat { Delimiter = ';', HasTrailingNewline = false };
            string outp = CsvWriter.Write(Rows(new[] { "a", "b;c", "d\"e", "f" }), fmt);
            Assert.Equal("a;\"b;c\";\"d\"\"e\";f", outp);
        }

        [Fact]
        public void QuoteAll_policy_quotes_every_field()
        {
            var fmt = new CsvFormat { Delimiter = ';', QuotePolicy = CsvQuotePolicy.All, HasTrailingNewline = false };
            string outp = CsvWriter.Write(Rows(new[] { "a", "b" }), fmt);
            Assert.Equal("\"a\";\"b\"", outp);
        }

        [Fact]
        public void Newline_and_delimiter_force_quotes()
        {
            var fmt = new CsvFormat { Delimiter = ',', HasTrailingNewline = false };
            string outp = CsvWriter.Write(Rows(new[] { "line1\nline2", "x" }), fmt);
            Assert.Equal("\"line1\nline2\",x", outp);
        }

        [Theory]
        [InlineData(CsvLineEnding.Crlf, "a;b\r\nc;d\r\n")]
        [InlineData(CsvLineEnding.Lf, "a;b\nc;d\n")]
        [InlineData(CsvLineEnding.Cr, "a;b\rc;d\r")]
        public void Honors_line_ending_and_trailing_newline(CsvLineEnding eol, string expected)
        {
            var fmt = new CsvFormat { Delimiter = ';', LineEnding = eol, HasTrailingNewline = true };
            Assert.Equal(expected, CsvWriter.Write(Rows(new[] { "a", "b" }, new[] { "c", "d" }), fmt));
        }

        [Fact]
        public void No_trailing_newline_when_disabled()
        {
            var fmt = new CsvFormat { Delimiter = ';', HasTrailingNewline = false };
            Assert.Equal("a;b", CsvWriter.Write(Rows(new[] { "a", "b" }), fmt));
        }

        [Fact]
        public void Comment_prefix_row_written_like_any_other_row()
        {
            // No special handling: '#' fields are neither quoted nor kept verbatim — just joined.
            var fmt = new CsvFormat { Delimiter = ';', HasTrailingNewline = true };
            Assert.Equal("#!format=2;x\r\na;b\r\n",
                CsvWriter.Write(Rows(new[] { "#!format=2", "x" }, new[] { "a", "b" }), fmt));
        }
    }
}
