namespace ExcelPowerTools.Core.Comparison.Diff
{
    public sealed class ExactComparisonOptions
    {
        public bool ComparePhysicalPresence { get; set; } = true;
        public bool CompareCachedResults { get; set; } = true;
        public bool CompareCachedResultsWhenFormulaChanges { get; set; }
    }
}
