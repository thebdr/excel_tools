using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Theme
{
    /// <summary>Semantic channels consumed by WinForms, Excel, report and snapshot renderers.</summary>
    public enum ComparisonThemeRole
    {
        Neutral = 0,
        Quality = 1,
        Special = 2,
        DiffAdded = 3,
        DiffRemoved = 4,
        AspectUnknown = 10,
        AspectFormula = 11,
        AspectValue = 12,
        AspectCachedResult = 13,
        AspectStructure = 14,
        AspectInteriorFill = 15,
        AspectCoverage = 16,
        StatusUnreviewed = 20,
        StatusUseMaster = 21,
        StatusUseCandidate = 22,
        StatusEquivalent = 23,
        StatusNeedsInvestigation = 24,
        StatusReviewedNoStructuralMerge = 25
    }

    public enum ComparisonThemeLevel
    {
        Lowest = 0,
        Low = 1,
        Medium = 2,
        High = 3,
        Highest = 4
    }

    /// <summary>Mappings only; concrete colors, fonts, icons and patterns belong to the add-in renderer.</summary>
    public static class ComparisonThemeSemantics
    {
        public static ComparisonThemeRole ForAspect(ComparisonAspect aspect)
        {
            switch (aspect)
            {
                case ComparisonAspect.Formula: return ComparisonThemeRole.AspectFormula;
                case ComparisonAspect.Value: return ComparisonThemeRole.AspectValue;
                case ComparisonAspect.CachedResult: return ComparisonThemeRole.AspectCachedResult;
                case ComparisonAspect.Structure: return ComparisonThemeRole.AspectStructure;
                case ComparisonAspect.InteriorFill: return ComparisonThemeRole.AspectInteriorFill;
                case ComparisonAspect.Coverage: return ComparisonThemeRole.AspectCoverage;
                default: return ComparisonThemeRole.AspectUnknown;
            }
        }

        public static ComparisonThemeRole ForStatus(ReviewDecisionStatus status)
        {
            switch (status)
            {
                case ReviewDecisionStatus.UseMaster: return ComparisonThemeRole.StatusUseMaster;
                case ReviewDecisionStatus.UseCandidate: return ComparisonThemeRole.StatusUseCandidate;
                case ReviewDecisionStatus.Equivalent: return ComparisonThemeRole.StatusEquivalent;
                case ReviewDecisionStatus.NeedsInvestigation: return ComparisonThemeRole.StatusNeedsInvestigation;
                case ReviewDecisionStatus.ReviewedNoStructuralMerge:
                    return ComparisonThemeRole.StatusReviewedNoStructuralMerge;
                default: return ComparisonThemeRole.StatusUnreviewed;
            }
        }
    }
}
