namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// The handful of Excel enum values we use, as plain ints, so the late-bound (dynamic)
    /// COM calls never depend on interop enum binding.
    /// </summary>
    internal static class Xl
    {
        public const int xlByRows = 1;
        public const int xlByColumns = 2;

        public const int xlNext = 1;
        public const int xlPrevious = 2;

        public const int xlWhole = 1;
        public const int xlPart = 2;
        public const int xlFormulas = -4123;
        public const int xlValues = -4163;

        public const int xlCellTypeVisible = 12;

        public const int xlCalculationAutomatic = -4105;
        public const int xlCalculationManual = -4135;

        public const int xlWBATWorksheet = -4167;       // Workbooks.Add: exactly one worksheet

        public const int xlColorIndexNone = -4142;      // no fill
        public const int xlColorIndexAutomatic = -4105; // automatic font colour

        public const int xlNumberAsText = 3;            // Range.Errors(3): "number stored as text"

        public const int xlSummaryOnLeft = -4131;       // Outline.SummaryColumn
        public const int xlSummaryOnRight = -4152;
    }
}
