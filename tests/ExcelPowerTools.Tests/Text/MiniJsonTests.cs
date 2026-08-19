using System.Text;
using ExcelPowerTools.Core.Text;
using Xunit;

namespace ExcelPowerTools.Tests.Text
{
    public class MiniJsonTests
    {
        private static JsonValue Parse(string json)
        {
            Assert.True(MiniJson.TryParse(json, out JsonValue value, out string error), error);
            return value;
        }

        [Fact]
        public void Parses_nested_objects_arrays_and_scalars()
        {
            JsonValue v = Parse("{\"a\":{\"b\":[1,\"two\",true,false,null]},\"c\":\"x\"}");

            Assert.Equal(JsonKind.Object, v.Kind);
            Assert.Equal("x", v.Str("c"));

            JsonValue b = v.Get("a")!.Get("b")!;
            Assert.Equal(JsonKind.Array, b.Kind);
            Assert.Equal(5, b.Items!.Count);
            Assert.Equal(JsonKind.Number, b.Items[0].Kind);
            Assert.Equal("1", b.Items[0].Text);
            Assert.Equal("two", b.Items[1].Text);
            Assert.True(b.Items[2].Flag);
            Assert.False(b.Items[3].Flag);
            Assert.Equal(JsonKind.Null, b.Items[4].Kind);
        }

        [Theory]
        [InlineData("\"a\\\"b\"", "a\"b")]
        [InlineData("\"a\\\\b\"", "a\\b")]
        [InlineData("\"a\\/b\"", "a/b")]
        [InlineData("\"\\b\\f\\n\\r\\t\"", "\b\f\n\r\t")]
        [InlineData("\"\\u0041\\u00e9\"", "Aé")]
        [InlineData("\"\\u2028\"", "\u2028")]
        public void Decodes_escapes(string json, string expected)
        {
            JsonValue v = Parse(json);
            Assert.Equal(JsonKind.String, v.Kind);
            Assert.Equal(expected, v.Text);
        }

        [Fact]
        public void Surrogate_pair_escapes_recombine()
        {
            JsonValue v = Parse("\"\\uD83D\\uDE00\"");
            Assert.Equal("\U0001F600", v.Text);
        }

        [Theory]
        [InlineData("-12.5e+3")]
        [InlineData("0")]
        [InlineData("-0.001")]
        [InlineData("1E9")]
        public void Numbers_keep_their_raw_text(string number)
        {
            JsonValue v = Parse(number);
            Assert.Equal(JsonKind.Number, v.Kind);
            Assert.Equal(number, v.Text);
        }

        [Theory]
        [InlineData("")]
        [InlineData("{")]
        [InlineData("{\"a\":}")]
        [InlineData("{\"a\":1,}")]         // trailing comma
        [InlineData("[1,]")]
        [InlineData("{'a':1}")]            // single quotes
        [InlineData("{\"a\":1} extra")]
        [InlineData("\"unterminated")]
        [InlineData("\"bad\\q\"")]
        [InlineData("01")]                 // leading zero then digit → trailing content
        [InlineData("nulll")]
        public void Rejects_malformed_input(string json)
        {
            Assert.False(MiniJson.TryParse(json, out _, out string error));
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void Accessors_tolerate_missing_and_mistyped_properties()
        {
            JsonValue v = Parse("{\"n\":5,\"s\":\"x\",\"b\":true,\"arr\":[1]}");

            Assert.Equal("", v.Str("absent"));
            Assert.Equal("fallback", v.Str("absent", "fallback"));
            Assert.Equal("5", v.Str("n"));                 // numbers read as text
            Assert.True(v.BoolOf("b", false));
            Assert.False(v.BoolOf("s", false));            // mistyped → fallback
            Assert.Empty(v.Arr("s"));                      // mistyped → empty list
            Assert.Single(v.Arr("arr"));
            Assert.Null(v.Get("absent"));
        }

        [Fact]
        public void Writer_output_round_trips_through_the_reader()
        {
            var sb = new StringBuilder();
            sb.Append('{');
            MiniJsonWriter.Property(sb, "plain", "value", first: true);
            MiniJsonWriter.Property(sb, "tricky", "a\"b\\c\nd\te\u2028f😀", first: false);
            MiniJsonWriter.BoolProperty(sb, "flag", true, first: false);
            MiniJsonWriter.NumberProperty(sb, "count", -42, first: false);
            sb.Append('}');

            JsonValue v = Parse(sb.ToString());
            Assert.Equal("value", v.Str("plain"));
            Assert.Equal("a\"b\\c\nd\te\u2028f😀", v.Str("tricky"));
            Assert.True(v.BoolOf("flag", false));
            Assert.Equal("-42", v.Str("count"));
        }

        [Fact]
        public void Deep_nesting_is_bounded_not_fatal()
        {
            string json = new string('[', 200) + new string(']', 200);
            Assert.False(MiniJson.TryParse(json, out _, out string error));
            Assert.Contains("deep", error);
        }
    }
}
