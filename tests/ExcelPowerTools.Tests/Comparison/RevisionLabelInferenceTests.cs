using System.Linq;
using ExcelPowerTools.Core.Comparison.Revision;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class RevisionLabelInferenceTests
    {
        [Theory]
        [InlineData("Book_r1.xlsx", "1")]
        [InlineData("Book-r02.xlsm", "2")]
        [InlineData("Book v2.1.xlsb", "2.1")]
        [InlineData("Book_rev3.csv", "3")]
        [InlineData("Book_VERSION_4_2.txt", "4.2")]
        public void Explicit_revision_tokens_are_strong(string fileName, string expected)
        {
            RevisionLabelInferenceResult result = RevisionLabelInference.Infer(fileName);

            Assert.True(result.HasStrongCandidate);
            Assert.Equal(RevisionTokenKind.ExplicitPrefix, result.StrongCandidate!.Kind);
            Assert.Equal(expected, result.StrongCandidate.Number.ToString());
        }

        [Fact]
        public void Terminal_31_keeps_whole_and_3_1_interpretations_without_choosing()
        {
            RevisionLabelInferenceResult result = RevisionLabelInference.Infer("Reviewed workbook 31.xlsx");

            Assert.False(result.HasStrongCandidate);
            Assert.Equal(2, result.Candidates.Count);
            Assert.Contains(result.Candidates, candidate =>
                candidate.Kind == RevisionTokenKind.TrailingWholeNumber && candidate.Number.ToString() == "31");
            Assert.Contains(result.Candidates, candidate =>
                candidate.Kind == RevisionTokenKind.TrailingSplitDecimal && candidate.Number.ToString() == "3.1");
        }

        [Fact]
        public void Explicit_version_at_end_does_not_add_a_false_terminal_candidate()
        {
            RevisionLabelInferenceResult result = RevisionLabelInference.Infer("Book_v2.1.xlsx");

            RevisionTokenCandidate candidate = Assert.Single(result.Candidates);
            Assert.Equal(RevisionTokenKind.ExplicitPrefix, candidate.Kind);
            Assert.Equal("2.1", candidate.Number.ToString());
        }

        [Fact]
        public void Product_number_without_prefix_is_only_weak_evidence()
        {
            RevisionLabelInferenceResult result = RevisionLabelInference.Infer("Pipeline4.xlsx");

            Assert.False(result.HasStrongCandidate);
            Assert.Equal("4", Assert.Single(result.Candidates).Number.ToString());
        }

        [Fact]
        public void Revision_numbers_sort_numerically_and_ignore_trailing_zero_components()
        {
            var values = new[]
            {
                new RevisionNumber(10),
                new RevisionNumber(2, 1),
                new RevisionNumber(2),
            };

            RevisionNumber[] ordered = values.OrderBy(value => value).ToArray();
            Assert.Equal("2", ordered[0].ToString());
            Assert.Equal("2.1", ordered[1].ToString());
            Assert.Equal("10", ordered[2].ToString());
            Assert.Equal(new RevisionNumber(2), new RevisionNumber(2, 0));
            Assert.Equal(new RevisionNumber(2).GetHashCode(), new RevisionNumber(2, 0).GetHashCode());
        }

        [Fact]
        public void Empty_file_name_is_rejected()
        {
            Assert.Throws<System.ArgumentException>(() => RevisionLabelInference.Infer(" "));
        }
    }
}
