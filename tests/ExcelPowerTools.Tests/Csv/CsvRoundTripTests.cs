using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvRoundTripTests
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        private static void AssertRoundTrip(byte[] input)
        {
            var doc = CsvFile.Load(input);
            byte[] output = doc.Save();
            Assert.True(input.SequenceEqual(output),
                $"Round-trip differed.\n in: {Show(input)}\nout: {Show(output)}");
        }

        private static byte[] Utf8(string s) => Utf8NoBom.GetBytes(s);
        private static string Show(byte[] b) => string.Join(" ", b.Select(x => x.ToString("X2")));

        [Fact] public void Crlf_trailing_newline() => AssertRoundTrip(Utf8("a;b;c\r\nd;e;f\r\n"));
        [Fact] public void Lf_trailing_newline() => AssertRoundTrip(Utf8("a;b;c\nd;e;f\n"));
        [Fact] public void No_trailing_newline() => AssertRoundTrip(Utf8("a;b\r\nc;d"));
        [Fact] public void Cr_only_endings() => AssertRoundTrip(Utf8("a;b\rc;d\r"));
        [Fact] public void Empty_fields_preserved() => AssertRoundTrip(Utf8("a;;c\r\n;;\r\n"));
        [Fact] public void Leading_trailing_spaces_significant() => AssertRoundTrip(Utf8("a; b ;c\r\n"));

        [Fact] public void Quoted_field_with_delimiter() => AssertRoundTrip(Utf8("a;\"b;c\";d\r\n"));
        [Fact] public void Quoted_field_with_embedded_newline() => AssertRoundTrip(Utf8("a;\"line1\r\nline2\";c\r\n"));
        [Fact] public void Doubled_quotes() => AssertRoundTrip(Utf8("a;\"he said \"\"hi\"\"\";b\r\n"));

        [Fact] public void Comment_lines_preserved() => AssertRoundTrip(Utf8("#comment\r\n#!format=2\r\na;b\r\n"));
        [Fact] public void Sep_directive_preserved() => AssertRoundTrip(Utf8("sep=;\r\na;b;c\r\n"));

        [Fact]
        public void Utf8_bom_preserved()
        {
            byte[] bom = { 0xEF, 0xBB, 0xBF };
            byte[] input = bom.Concat(Utf8("a;b\r\n")).ToArray();
            AssertRoundTrip(input);
        }

        [Fact]
        public void Windows1252_bytes_preserved()
        {
            // "aäböc" in Windows-1252: ä=0xE4, ö=0xF6. Not valid UTF-8, so detection falls back to CP1252.
            byte[] input = { 0x61, 0xE4, 0x62, 0xF6, 0x63, 0x3B, 0x78, 0x0D, 0x0A }; // aäböc;x\r\n
            AssertRoundTrip(input);
        }

        [Fact]
        public void Utf16_le_bom_preserved()
        {
            byte[] input = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("a;b\r\n")).ToArray();
            AssertRoundTrip(input);
        }

        [Fact]
        public void Detected_delimiter_comma_round_trips()
        {
            AssertRoundTrip(Utf8("a,b,c\r\nd,e,f\r\n"));
        }

        [Fact]
        public void Editing_a_cell_only_changes_that_cell()
        {
            var doc = CsvFile.Load(Utf8("a;b;c\r\nd;e;f\r\n"));
            doc.Records[1].Fields[1] = "E";   // change 'e' -> 'E'
            string outText = Utf8NoBom.GetString(doc.Save());
            Assert.Equal("a;b;c\r\nd;E;f\r\n", outText);
        }
    }
}
