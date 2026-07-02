using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvParserTests
    {
        private static CsvFormat Semi() => new CsvFormat { Delimiter = ';', Quote = '"', CommentPrefix = "#" };

        [Fact]
        public void Splits_simple_fields()
        {
            var recs = CsvParser.Parse("a;b;c", Semi());
            Assert.Single(recs);
            Assert.Equal(new[] { "a", "b", "c" }, recs[0].Fields);
        }

        [Fact]
        public void Quoted_field_keeps_delimiter()
        {
            var recs = CsvParser.Parse("a;\"b;c\";d", Semi());
            Assert.Equal(new[] { "a", "b;c", "d" }, recs[0].Fields);
        }

        [Fact]
        public void Doubled_quote_becomes_single()
        {
            var recs = CsvParser.Parse("\"he said \"\"hi\"\"\"", Semi());
            Assert.Equal("he said \"hi\"", recs[0].Fields[0]);
        }

        [Fact]
        public void Embedded_newline_stays_in_one_record()
        {
            var recs = CsvParser.Parse("a;\"x\ny\";b\r\nc;d;e", Semi());
            Assert.Equal(2, recs.Count);
            Assert.Equal(new[] { "a", "x\ny", "b" }, recs[0].Fields);
            Assert.Equal(new[] { "c", "d", "e" }, recs[1].Fields);
        }

        [Fact]
        public void Trailing_newline_does_not_create_empty_record()
        {
            var recs = CsvParser.Parse("a;b\r\n", Semi());
            Assert.Single(recs);
        }

        [Fact]
        public void Blank_line_is_an_empty_record()
        {
            var recs = CsvParser.Parse("a;b\r\n\r\nc;d", Semi());
            Assert.Equal(3, recs.Count);
            Assert.Equal(new[] { "" }, recs[1].Fields);
        }

        [Fact]
        public void Comment_prefix_line_is_split_like_any_other_row()
        {
            var recs = CsvParser.Parse("#hello;world\r\na;b", Semi());
            Assert.Equal(new[] { "#hello", "world" }, recs[0].Fields);
            Assert.Equal(new[] { "a", "b" }, recs[1].Fields);
        }

        [Fact]
        public void Empty_text_yields_no_records()
        {
            Assert.Empty(CsvParser.Parse("", Semi()));
        }

        [Fact]
        public void Empty_trailing_field_preserved()
        {
            var recs = CsvParser.Parse("a;b;", Semi());
            Assert.Equal(new[] { "a", "b", "" }, recs[0].Fields);
        }
    }
}
