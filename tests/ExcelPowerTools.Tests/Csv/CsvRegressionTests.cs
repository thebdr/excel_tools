using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    /// <summary>Regression tests for defects found in the Core.Csv adversarial review.</summary>
    public class CsvRegressionTests
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);
        private static CsvFormat Semi() => new CsvFormat { Delimiter = ';', Quote = '"', CommentPrefix = "#" };
        private static byte[] Utf8(string s) => Utf8NoBom.GetBytes(s);

        private static void AssertRoundTrip(byte[] input)
        {
            byte[] output = CsvFile.Load(input).Save();
            Assert.True(input.SequenceEqual(output),
                $"Round-trip differed.\n in: {Show(input)}\nout: {Show(output)}");
        }
        private static string Show(byte[] b) => string.Join(" ", b.Select(x => x.ToString("X2")));

        // #1 — a quote mid-field must be literal and must NOT swallow the delimiter.
        [Fact]
        public void Quote_mid_field_is_literal_and_delimiter_survives()
        {
            var recs = CsvParser.Parse("x\"y;z", Semi());
            Assert.Single(recs);
            Assert.Equal(new[] { "x\"y", "z" }, recs[0].Fields);
        }

        // #3/#12 — a space (or text) before a quote keeps the quotes as literal characters.
        [Fact]
        public void Space_before_quote_keeps_quotes_as_literal()
        {
            var recs = CsvParser.Parse(" \"a\";b", Semi());
            Assert.Equal(new[] { " \"a\"", "b" }, recs[0].Fields);
        }

        // '#' rows are ordinary rows: split like data, values stable across round-trips. A source that
        // quoted "#x" unnecessarily is unquoted once (minimal writer), then the file is a stable fixpoint.
        [Fact]
        public void Comment_prefix_first_field_is_stable_across_roundtrips()
        {
            var doc = CsvFile.Load(Utf8("\"#x\";y\r\n"));
            Assert.Equal(new[] { "#x", "y" }, doc.Records[0].Fields);

            byte[] save1 = doc.Save();
            Assert.Equal("#x;y\r\n", Utf8NoBom.GetString(save1));

            var doc2 = CsvFile.Load(save1);
            Assert.Equal(new[] { "#x", "y" }, doc2.Records[0].Fields);   // still the same values
            Assert.True(save1.SequenceEqual(doc2.Save()));               // and byte-stable from here on
        }

        // Comment rows split into fields and round-trip byte-identically, including Excel-padded directives.
        [Fact] public void Comment_row_with_delimiters_round_trips() => AssertRoundTrip(Utf8("#a;b;c\r\nx;y;z\r\n"));
        [Fact] public void Padded_directive_line_round_trips() => AssertRoundTrip(Utf8("#!format=2,,,,\r\na,b,c,d,e\r\n"));

        // #4 — UTF-16 without a BOM is detected and round-trips.
        [Fact]
        public void Utf16_le_without_bom_detected_and_round_trips()
        {
            byte[] input = Encoding.Unicode.GetBytes("a;b\r\nc;d\r\n");   // no preamble
            var doc = CsvFile.Load(input);
            Assert.Equal(1200, doc.Format.Encoding.CodePage);
            Assert.False(doc.Format.HasBom);
            Assert.True(input.SequenceEqual(doc.Save()));
        }

        // #6/#11 — mixed line endings and a bare CR at EOF round-trip exactly.
        [Fact] public void Mixed_line_endings_round_trip() => AssertRoundTrip(Utf8("a;b\r\nc;d\nx;y\r\n"));
        [Fact] public void Bare_cr_at_eof_round_trips() => AssertRoundTrip(Utf8("a;b\r\nc;d\r"));

        // #10 — a sep= line whose terminator differs from the data rows round-trips exactly.
        [Fact] public void Sep_line_with_different_terminator_round_trips() => AssertRoundTrip(Utf8("sep=;\na;b\r\nc;d\r\n"));

        // #7/#8 — a programmatically-null Fields array never throws in the writer or detectors.
        [Fact]
        public void Null_fields_are_safe()
        {
            var doc = new CsvDocument { Format = new CsvFormat() };
            doc.Records.Add(CsvRecord.Data(null));
            Assert.NotNull(doc.Save());
            var ragged = CsvDetect.FindRaggedRows(doc.DataRows(), out _);
            Assert.NotNull(ragged);
        }

        // #13 — a width-frequency tie breaks toward the first-seen width.
        [Fact]
        public void Ragged_tie_break_prefers_first_seen_width()
        {
            var rows = new[]
            {
                new[] { "a", "b" }, new[] { "c", "d" },
                new[] { "1", "2", "3" }, new[] { "4", "5", "6" },
            };
            var ragged = CsvDetect.FindRaggedRows(rows, out int expected);
            Assert.Equal(2, expected);
            Assert.Equal(new[] { 2, 3 }, ragged);
        }

        // #14 — a sep= directive whose delimiter would equal the quote char is ignored.
        [Fact]
        public void Sep_directive_with_quote_char_is_ignored()
        {
            var doc = CsvFile.Load(Utf8("sep=\"\r\na;b\r\n"));
            Assert.False(doc.Format.HasSepDirective);
        }
    }
}
