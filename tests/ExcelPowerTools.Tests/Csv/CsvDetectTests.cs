using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvDetectTests
    {
        [Theory]
        [InlineData("a;b;c\r\nd;e;f", ';')]
        [InlineData("a,b,c\r\nd,e,f", ',')]
        [InlineData("a\tb\tc\r\nd\te\tf", '\t')]
        [InlineData("a|b|c\r\nd|e|f", '|')]
        public void Detects_delimiter_by_consistency(string text, char expected)
        {
            Assert.Equal(expected, CsvDetect.DetectDelimiter(text, CsvDetect.DefaultDelimiters, '"', ';'));
        }

        [Fact]
        public void Delimiter_inside_quotes_is_not_counted()
        {
            // Commas only appear inside a quoted field; ';' is the real delimiter.
            string text = "a;\"x,y,z\";b\r\nc;\"p,q,r\";d";
            Assert.Equal(';', CsvDetect.DetectDelimiter(text, CsvDetect.DefaultDelimiters, '"', ','));
        }

        [Fact]
        public void Falls_back_when_no_delimiter_present()
        {
            Assert.Equal(';', CsvDetect.DetectDelimiter("singlecolumn\r\nanother", CsvDetect.DefaultDelimiters, '"', ';'));
        }

        [Theory]
        [InlineData("a\r\nb\r\n", CsvLineEnding.Crlf)]
        [InlineData("a\nb\n", CsvLineEnding.Lf)]
        [InlineData("a\rb\r", CsvLineEnding.Cr)]
        [InlineData("nolineending", CsvLineEnding.Crlf)]
        public void Detects_line_ending(string text, CsvLineEnding expected)
        {
            Assert.Equal(expected, CsvDetect.DetectLineEnding(text));
        }

        [Theory]
        [InlineData("a;b\r\n", true)]
        [InlineData("a;b", false)]
        public void Detects_trailing_newline(string text, bool expected)
        {
            Assert.Equal(expected, CsvDetect.HasTrailingNewline(text));
        }

        [Theory]
        [InlineData("sep=;", true, ';')]
        [InlineData("sep=,", true, ',')]
        [InlineData("SEP=|", true, '|')]
        [InlineData("a;b;c", false, '\0')]
        [InlineData("sep=abc", false, '\0')]
        public void Parses_sep_directive(string line, bool ok, char expected)
        {
            bool got = CsvDetect.TryParseSepDirective(line, out char sep);
            Assert.Equal(ok, got);
            if (ok) Assert.Equal(expected, sep);
        }

        [Fact]
        public void Header_is_first_wordy_full_width_row()
        {
            var rows = new[]
            {
                new[] { "Station", "IP", "Slot" },   // header (all wordy)
                new[] { "S1", "192.168.0.1", "1" },
                new[] { "S2", "192.168.0.2", "2" },
            };
            Assert.Equal(0, CsvDetect.DetectHeaderRow(rows, "#"));
        }

        [Fact]
        public void Header_detected_after_preamble_rows()
        {
            var rows = new[]
            {
                new[] { "10", "20", "30" },          // numeric row (not a header)
                new[] { "Name", "Value", "Unit" },   // header
                new[] { "x", "1", "V" },
                new[] { "y", "2", "A" },
            };
            Assert.Equal(1, CsvDetect.DetectHeaderRow(rows, "#"));
        }

        [Fact]
        public void Commented_header_row_is_detected()
        {
            // The hardware-config convention: the header itself is a '#' comment line.
            var rows = new[]
            {
                new[] { "# exported 2026-07-02" },            // prose comment (width 1 — not a candidate)
                new[] { "#Name", "IP", "Slot" },              // commented header
                new[] { "S1", "192.168.0.1", "1" },
                new[] { "S2", "192.168.0.2", "2" },
            };
            Assert.Equal(1, CsvDetect.DetectHeaderRow(rows, "#"));
        }

        [Fact]
        public void Padded_directive_is_not_a_header()
        {
            // "#!format=2,,,," has full width but is mostly empty — must not beat the real header.
            var rows = new[]
            {
                new[] { "#!format=2", "", "" },
                new[] { "#Name", "IP", "Slot" },
                new[] { "S1", "192.168.0.1", "1" },
            };
            Assert.Equal(1, CsvDetect.DetectHeaderRow(rows, "#"));
        }

        [Fact]
        public void Ragged_rows_flagged_against_modal_width()
        {
            var rows = new[]
            {
                new[] { "a", "b", "c" },
                new[] { "d", "e" },        // short
                new[] { "f", "g", "h" },
                new[] { "i", "j", "k", "l" }, // long
            };
            var ragged = CsvDetect.FindRaggedRows(rows, out int expected);
            Assert.Equal(3, expected);
            Assert.Equal(new[] { 1, 3 }, ragged);
        }

        [Theory]
        [InlineData("123", true)]
        [InlineData("-42", true)]
        [InlineData("3,14", true)]
        [InlineData("1.234,56", true)]
        [InlineData("1.23E+10", true)]
        [InlineData("Station", false)]
        [InlineData("192.168.0.1", false)]
        [InlineData("", false)]
        public void LooksNumeric_classifies_tokens(string s, bool expected)
        {
            Assert.Equal(expected, CsvDetect.LooksNumeric(s));
        }
    }
}
