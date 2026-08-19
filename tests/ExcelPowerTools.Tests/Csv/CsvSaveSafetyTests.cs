using System.Text;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvSaveSafetyTests
    {
        private static CsvSourceFingerprint F(string text) =>
            CsvSourceFingerprint.FromBytes(Encoding.UTF8.GetBytes(text));

        [Theory]
        [InlineData("A", "A", "A", false, CsvSaveDisposition.UpToDate)]
        [InlineData("A", "A", "C", false, CsvSaveDisposition.UpToDate)]
        [InlineData("A", "A", "C", true, CsvSaveDisposition.SafeToWrite)]
        [InlineData("A", "B", "B", false, CsvSaveDisposition.AlreadySynchronized)]
        [InlineData("A", "B", "B", true, CsvSaveDisposition.AlreadySynchronized)]
        [InlineData("A", "B", "A", false, CsvSaveDisposition.ExternalChange)]
        [InlineData("A", "B", "C", false, CsvSaveDisposition.ExternalChange)]
        [InlineData("A", "B", "C", true, CsvSaveDisposition.Conflict)]
        public void Present_source_decision_matrix(
            string baseline, string observed, string proposed, bool local, CsvSaveDisposition expected)
        {
            Assert.Equal(expected, CsvSaveSafety.Assess(F(baseline), F(observed), F(proposed), local));
        }

        [Fact]
        public void Missing_source_distinguishes_local_edits()
        {
            Assert.Equal(CsvSaveDisposition.SourceMissing,
                CsvSaveSafety.Assess(F("A"), null, F("A"), localValuesChanged: false));
            Assert.Equal(CsvSaveDisposition.SourceMissingWithLocalChanges,
                CsvSaveSafety.Assess(F("A"), null, F("C"), localValuesChanged: true));
        }

        [Fact]
        public void Unavailable_source_cannot_be_verified() =>
            Assert.Equal(CsvSaveDisposition.CannotVerify,
                CsvSaveSafety.Assess(F("A"), null, F("C"), true, sourceReadable: false));
    }
}
