using System;
using System.Globalization;
using System.Text;
using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Diff
{
    [Flags]
    public enum EquivalenceTag
    {
        None = 0,
        CaseOnly = 1,
        WhitespaceOnly = 2,
        PunctuationOnly = 4,
        NumericallyEquivalent = 8,
        BlankAndEmptyText = 16,
    }

    /// <summary>
    /// Adds quiet semantic labels to an exact atomic difference. These labels never alter or suppress the
    /// underlying before/after evidence.
    /// </summary>
    public static class EquivalenceClassifier
    {
        public static EquivalenceTag Classify(AtomicCellChange change)
        {
            if (change == null) throw new ArgumentNullException(nameof(change));
            if (change.Aspect != CellChangeAspect.Formula && change.Aspect != CellChangeAspect.Value)
                return EquivalenceTag.None;

            string before = change.BeforeText ?? "";
            string after = change.AfterText ?? "";
            EquivalenceTag tags = TextTags(before, after);

            if (change.Aspect == CellChangeAspect.Value)
            {
                if (IsBlankAndEmptyText(change.BeforeValue, change.AfterValue))
                    tags |= EquivalenceTag.BlankAndEmptyText;
                if (AreNumericallyEquivalent(change.BeforeValue, change.AfterValue))
                    tags |= EquivalenceTag.NumericallyEquivalent;
            }

            return tags;
        }

        private static EquivalenceTag TextTags(string before, string after)
        {
            EquivalenceTag tags = EquivalenceTag.None;
            if (!string.Equals(before, after, StringComparison.Ordinal) &&
                string.Equals(before, after, StringComparison.OrdinalIgnoreCase))
                tags |= EquivalenceTag.CaseOnly;

            if (!string.Equals(before, after, StringComparison.Ordinal) &&
                string.Equals(RemoveWhitespace(before), RemoveWhitespace(after), StringComparison.Ordinal))
                tags |= EquivalenceTag.WhitespaceOnly;

            if (!string.Equals(before, after, StringComparison.Ordinal) &&
                string.Equals(RemovePunctuation(before), RemovePunctuation(after), StringComparison.Ordinal))
                tags |= EquivalenceTag.PunctuationOnly;

            return tags;
        }

        private static bool IsBlankAndEmptyText(CanonicalValue before, CanonicalValue after)
        {
            if (before == null || after == null) return false;
            return (before.Kind == CanonicalValueKind.Blank &&
                    after.Kind == CanonicalValueKind.Text && after.ExactText.Length == 0) ||
                   (after.Kind == CanonicalValueKind.Blank &&
                    before.Kind == CanonicalValueKind.Text && before.ExactText.Length == 0);
        }

        private static bool AreNumericallyEquivalent(CanonicalValue before, CanonicalValue after)
        {
            double beforeNumber;
            double afterNumber;
            return TryComparableNumber(before, out beforeNumber) &&
                   TryComparableNumber(after, out afterNumber) &&
                   beforeNumber.Equals(afterNumber);
        }

        private static bool TryComparableNumber(CanonicalValue value, out double number)
        {
            number = 0;
            if (value == null) return false;
            if (value.Kind != CanonicalValueKind.Number && value.Kind != CanonicalValueKind.Text) return false;
            return double.TryParse(
                value.ExactText,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out number) &&
                   !double.IsNaN(number) &&
                   !double.IsInfinity(number);
        }

        private static string RemoveWhitespace(string text)
        {
            var result = new StringBuilder(text.Length);
            foreach (char ch in text)
                if (!char.IsWhiteSpace(ch)) result.Append(ch);
            return result.ToString();
        }

        private static string RemovePunctuation(string text)
        {
            var result = new StringBuilder(text.Length);
            foreach (char ch in text)
                if (!char.IsPunctuation(ch)) result.Append(ch);
            return result.ToString();
        }
    }
}
