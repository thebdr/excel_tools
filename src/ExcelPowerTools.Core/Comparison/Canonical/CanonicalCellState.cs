using System;

namespace ExcelPowerTools.Core.Comparison.Canonical
{
    /// <summary>Exact primary content observed for one already-aligned logical cell subject.</summary>
    public sealed class CanonicalCellState
    {
        private CanonicalCellState(
            string subjectKey,
            string formulaText,
            string formulaMetadata,
            CanonicalValue constantValue,
            CanonicalValue cachedResult,
            string displayText)
        {
            if (string.IsNullOrEmpty(subjectKey))
                throw new ArgumentException("A stable logical subject key is required.", nameof(subjectKey));

            SubjectKey = subjectKey;
            FormulaText = formulaText;
            FormulaMetadata = formulaMetadata;
            ConstantValue = constantValue ?? CanonicalValue.Blank;
            CachedResult = cachedResult;
            DisplayText = displayText;
        }

        public string SubjectKey { get; }
        public string FormulaText { get; }
        /// <summary>Canonical exact formula attributes/group evidence, separate from character-diff text.</summary>
        public string FormulaMetadata { get; }
        public CanonicalValue ConstantValue { get; }
        public CanonicalValue CachedResult { get; }
        public string DisplayText { get; }
        public bool HasFormula => FormulaText != null;

        public static CanonicalCellState Blank(string subjectKey) =>
            new CanonicalCellState(subjectKey, null, null, CanonicalValue.Blank, null, null);

        public static CanonicalCellState Constant(
            string subjectKey,
            CanonicalValue value,
            string displayText = null)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            return new CanonicalCellState(subjectKey, null, null, value, null, displayText);
        }

        public static CanonicalCellState Formula(
            string subjectKey,
            string exactFormulaText,
            CanonicalValue cachedResult = null,
            string displayText = null,
            string exactFormulaMetadata = null)
        {
            if (string.IsNullOrEmpty(exactFormulaText))
                throw new ArgumentException("Exact formula text is required.", nameof(exactFormulaText));

            return new CanonicalCellState(
                subjectKey,
                exactFormulaText,
                exactFormulaMetadata,
                CanonicalValue.Blank,
                cachedResult,
                displayText);
        }
    }
}
