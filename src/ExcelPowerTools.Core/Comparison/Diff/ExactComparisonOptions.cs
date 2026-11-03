namespace ExcelPowerTools.Core.Comparison.Diff
{
    public sealed class ExactComparisonOptions
    {
        // Blank-versus-missing storage is package evidence, not user content. Formula/value additions and
        // removals are still reported independently, so physical presence stays opt-in.
        public bool ComparePhysicalPresence { get; set; }
        public bool CompareCachedResults { get; set; } = true;
        // A changed formula and its changed stored calculation result are separate facts that both matter.
        public bool CompareCachedResultsWhenFormulaChanges { get; set; } = true;
    }
}
