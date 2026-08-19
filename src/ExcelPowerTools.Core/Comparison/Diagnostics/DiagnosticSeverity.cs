namespace ExcelPowerTools.Core.Comparison.Diagnostics
{
    /// <summary>Ordered operational impact; presentation color is assigned by the add-in theme renderer.</summary>
    public enum DiagnosticSeverity
    {
        Information = 0,
        Warning = 1,
        Correctable = 2,
        ArtifactBlocking = 3,
        ProjectBlocking = 4
    }
}
