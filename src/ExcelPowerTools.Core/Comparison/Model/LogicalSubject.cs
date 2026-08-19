namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>
    /// A stable logical cell aspect. Physical row numbers and addresses are intentionally only provenance and
    /// never form the identity after row/column alignment.
    /// </summary>
    public sealed class LogicalSubject
    {
        public string SubjectId { get; set; } = "";
        public string LogicalSheetId { get; set; } = "";
        public string RowIdentity { get; set; } = "";
        public string ColumnIdentity { get; set; } = "";
        public ComparisonAspect Aspect { get; set; }
    }
}
