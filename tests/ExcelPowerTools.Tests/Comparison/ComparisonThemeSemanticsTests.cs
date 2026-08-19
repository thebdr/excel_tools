using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Theme;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class ComparisonThemeSemanticsTests
    {
        [Theory]
        [InlineData(ComparisonAspect.Formula, ComparisonThemeRole.AspectFormula)]
        [InlineData(ComparisonAspect.Value, ComparisonThemeRole.AspectValue)]
        [InlineData(ComparisonAspect.CachedResult, ComparisonThemeRole.AspectCachedResult)]
        [InlineData(ComparisonAspect.Structure, ComparisonThemeRole.AspectStructure)]
        [InlineData(ComparisonAspect.InteriorFill, ComparisonThemeRole.AspectInteriorFill)]
        [InlineData(ComparisonAspect.Coverage, ComparisonThemeRole.AspectCoverage)]
        [InlineData(ComparisonAspect.Unknown, ComparisonThemeRole.AspectUnknown)]
        public void Aspect_maps_to_a_semantic_role(ComparisonAspect aspect, ComparisonThemeRole expected)
        {
            Assert.Equal(expected, ComparisonThemeSemantics.ForAspect(aspect));
        }

        [Theory]
        [InlineData(ReviewDecisionStatus.Unreviewed, ComparisonThemeRole.StatusUnreviewed)]
        [InlineData(ReviewDecisionStatus.UseMaster, ComparisonThemeRole.StatusUseMaster)]
        [InlineData(ReviewDecisionStatus.UseCandidate, ComparisonThemeRole.StatusUseCandidate)]
        [InlineData(ReviewDecisionStatus.Equivalent, ComparisonThemeRole.StatusEquivalent)]
        [InlineData(ReviewDecisionStatus.NeedsInvestigation, ComparisonThemeRole.StatusNeedsInvestigation)]
        [InlineData(ReviewDecisionStatus.ReviewedNoStructuralMerge, ComparisonThemeRole.StatusReviewedNoStructuralMerge)]
        public void Review_status_maps_to_a_semantic_role(
            ReviewDecisionStatus status,
            ComparisonThemeRole expected)
        {
            Assert.Equal(expected, ComparisonThemeSemantics.ForStatus(status));
        }
    }
}
