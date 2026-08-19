using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Diff
{
    public enum CellChangeAspect
    {
        Presence,
        Formula,
        FormulaStructure,
        Value,
        CachedResult,
    }

    public enum CellChangeKind
    {
        Added,
        Removed,
        Modified,
    }

    /// <summary>One exact, non-destructive formula/value observation between two revision snapshots.</summary>
    public sealed class AtomicCellChange
    {
        internal AtomicCellChange(
            string subjectKey,
            CellChangeAspect aspect,
            CellChangeKind kind,
            CanonicalCellState beforeCell,
            CanonicalCellState afterCell,
            string beforeText,
            string afterText,
            CanonicalValue beforeValue,
            CanonicalValue afterValue)
        {
            SubjectKey = subjectKey;
            Aspect = aspect;
            Kind = kind;
            BeforeCell = beforeCell;
            AfterCell = afterCell;
            BeforeText = beforeText;
            AfterText = afterText;
            BeforeValue = beforeValue;
            AfterValue = afterValue;
        }

        public string SubjectKey { get; }
        public CellChangeAspect Aspect { get; }
        public CellChangeKind Kind { get; }
        public CanonicalCellState BeforeCell { get; }
        public CanonicalCellState AfterCell { get; }
        public string BeforeText { get; }
        public string AfterText { get; }
        public CanonicalValue BeforeValue { get; }
        public CanonicalValue AfterValue { get; }
        public bool IsSecondary => Aspect == CellChangeAspect.CachedResult || Aspect == CellChangeAspect.Presence;
    }
}
