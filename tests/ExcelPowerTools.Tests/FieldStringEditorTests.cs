using ExcelPowerTools.Core;
using Xunit;

namespace ExcelPowerTools.Tests
{
    public class FieldStringEditorTests
    {
        [Fact]
        public void GetValuesBetween_ReturnsValuesBetweenDelimiters()
        {
            var values = FieldStringEditor.GetValuesBetween("aa(1).bb(2).cc(3)", "(", ")");
            Assert.Equal(new[] { "1", "2", "3" }, values);
        }

        [Fact]
        public void GetValuesBetween_NoMatch_ReturnsEmpty()
        {
            var values = FieldStringEditor.GetValuesBetween("no delimiters here", "(", ")");
            Assert.Empty(values);
        }

        [Fact]
        public void RebuildWithValues_ReplacesEachValue()
        {
            var rebuilt = FieldStringEditor.RebuildWithValues(
                "aa(1).bb(2).cc(3)", "(", ")", new[] { "10", "20", "30" });
            Assert.Equal("aa(10).bb(20).cc(30)", rebuilt);
        }

        [Fact]
        public void RebuildWithValues_ExtraOccurrencesLeftUntouched()
        {
            var rebuilt = FieldStringEditor.RebuildWithValues(
                "aa(1).bb(2).cc(3)", "(", ")", new[] { "10" });
            Assert.Equal("aa(10).bb(2).cc(3)", rebuilt);
        }

        [Theory]
        [InlineData("+5", FieldOperation.Add, 5)]
        [InlineData("-3", FieldOperation.Subtract, 3)]
        public void ParseNewValue_SignedNumbers(string input, FieldOperation expected, long delta)
        {
            var op = FieldStringEditor.ParseNewValue(input);
            Assert.Equal(expected, op.Operation);
            Assert.Equal(delta, op.Delta);
        }

        [Theory]
        [InlineData("7")]          // plain number is a literal replacement, not an operation
        [InlineData("hello")]
        [InlineData("+abc")]       // signed but non-numeric -> falls back to replace
        public void ParseNewValue_LiteralReplacement(string input)
        {
            var op = FieldStringEditor.ParseNewValue(input);
            Assert.Equal(FieldOperation.Replace, op.Operation);
            Assert.Equal(input, op.ReplaceValue);
        }

        [Fact]
        public void ChangeField_SingleSeparator_Replace()
        {
            var op = FieldStringEditor.ParseNewValue("905");
            var r = FieldStringEditor.ChangeField("I902.0", ".", "", 0, op);
            Assert.True(r.Changed);
            Assert.Equal("905.0", r.NewValue);
        }

        [Fact]
        public void ChangeField_SingleSeparator_AddPreservesWidth()
        {
            var op = FieldStringEditor.ParseNewValue("+2");
            var r = FieldStringEditor.ChangeField("008.100", ".", "", 0, op);
            Assert.True(r.Changed);
            Assert.Equal("010.100", r.NewValue);
        }

        [Fact]
        public void ChangeField_SingleSeparator_SubtractPreservesWidth()
        {
            var op = FieldStringEditor.ParseNewValue("-3");
            var r = FieldStringEditor.ChangeField("010.100", ".", "", 0, op);
            Assert.True(r.Changed);
            Assert.Equal("007.100", r.NewValue);
        }

        [Fact]
        public void ChangeField_DoubleSeparator_Replace()
        {
            var op = FieldStringEditor.ParseNewValue("5");
            var r = FieldStringEditor.ChangeField("aa(1).bb(2)", "(", ")", 1, op);
            Assert.True(r.Changed);
            Assert.Equal("aa(1).bb(5)", r.NewValue);
        }

        [Fact]
        public void ChangeField_DoubleSeparator_Add()
        {
            var op = FieldStringEditor.ParseNewValue("+10");
            var r = FieldStringEditor.ChangeField("aa(01).bb(02)", "(", ")", 0, op);
            Assert.True(r.Changed);
            Assert.Equal("aa(11).bb(02)", r.NewValue);
        }

        [Fact]
        public void ChangeField_NoSeparator_Skipped()
        {
            var op = FieldStringEditor.ParseNewValue("9");
            var r = FieldStringEditor.ChangeField("plaintext", ".", "", 0, op);
            Assert.Equal(FieldChangeStatus.SkippedNoSeparator, r.Status);
            Assert.Equal("plaintext", r.NewValue);
        }

        [Fact]
        public void ChangeField_IndexOutOfRange_Skipped()
        {
            var op = FieldStringEditor.ParseNewValue("9");
            var r = FieldStringEditor.ChangeField("a.b", ".", "", 5, op);
            Assert.Equal(FieldChangeStatus.SkippedIndexOutOfRange, r.Status);
            Assert.Equal(2, r.FieldCount);
        }

        [Fact]
        public void ChangeField_AddOnNonNumericField_Skipped()
        {
            var op = FieldStringEditor.ParseNewValue("+1");
            var r = FieldStringEditor.ChangeField("abc.def", ".", "", 0, op);
            Assert.Equal(FieldChangeStatus.SkippedNotNumeric, r.Status);
        }

        [Fact]
        public void ZeroPad_HandlesNegativeAndWidth()
        {
            Assert.Equal("009", FieldStringEditor.ZeroPad(9, 3));
            Assert.Equal("-007", FieldStringEditor.ZeroPad(-7, 3));
            Assert.Equal("1234", FieldStringEditor.ZeroPad(1234, 3)); // wider than pad -> unchanged
        }
    }
}
