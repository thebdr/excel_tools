using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelPowerTools.Core
{
    /// <summary>What to do to the targeted field.</summary>
    public enum FieldOperation
    {
        Replace,
        Add,
        Subtract,
    }

    /// <summary>Outcome of trying to change a field in one cell string.</summary>
    public enum FieldChangeStatus
    {
        Changed,
        SkippedNoSeparator,
        SkippedIndexOutOfRange,
        SkippedNotNumeric,
    }

    /// <summary>Parsed "new value" box: a leading + / - means add/subtract, anything else is a literal replacement.</summary>
    public sealed class FieldChangeOperation
    {
        public FieldOperation Operation { get; }
        public string ReplaceValue { get; }
        public long Delta { get; }

        public FieldChangeOperation(FieldOperation operation, string replaceValue, long delta)
        {
            Operation = operation;
            ReplaceValue = replaceValue ?? string.Empty;
            Delta = delta;
        }
    }

    /// <summary>Result of <see cref="FieldStringEditor.ChangeField"/> for a single cell.</summary>
    public sealed class FieldChangeResult
    {
        public FieldChangeStatus Status { get; }
        public string NewValue { get; }
        public int FieldCount { get; }

        public FieldChangeResult(FieldChangeStatus status, string newValue, int fieldCount)
        {
            Status = status;
            NewValue = newValue;
            FieldCount = fieldCount;
        }

        public bool Changed => Status == FieldChangeStatus.Changed;
    }

    /// <summary>
    /// Pure, Excel-independent port of the VBA IoListAux "ChangeField" logic:
    /// split a delimited string, replace / add / subtract one field, then rebuild it.
    /// Supports single-separator mode ("a.b.c") and value-between-separators mode ("a(1).b(2)").
    /// </summary>
    public static class FieldStringEditor
    {
        /// <summary>
        /// Returns every value found between <paramref name="sep1"/> and the next <paramref name="sep2"/>.
        /// GetValuesBetween("aa(1).bb(2).cc(3)", "(", ")") -> ["1","2","3"].
        /// </summary>
        public static string[] GetValuesBetween(string text, string sep1, string sep2)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            if (string.IsNullOrEmpty(sep1)) throw new ArgumentException("sep1 must be non-empty", nameof(sep1));
            if (string.IsNullOrEmpty(sep2)) throw new ArgumentException("sep2 must be non-empty", nameof(sep2));

            var result = new List<string>();
            int searchFrom = 0;
            while (true)
            {
                int posStart = text.IndexOf(sep1, searchFrom, StringComparison.Ordinal);
                if (posStart < 0) break;

                int valueStart = posStart + sep1.Length;
                int posEnd = text.IndexOf(sep2, valueStart, StringComparison.Ordinal);
                if (posEnd < 0) break;

                result.Add(text.Substring(valueStart, posEnd - valueStart));
                searchFrom = posEnd + sep2.Length;
            }

            return result.ToArray();
        }

        /// <summary>
        /// Rebuilds <paramref name="text"/>, replacing the i-th value between sep1/sep2 with newValues[i].
        /// Occurrences beyond the supplied values are left untouched.
        /// RebuildWithValues("aa(1).bb(2).cc(3)", "(", ")", ["10","20","30"]) -> "aa(10).bb(20).cc(30)".
        /// </summary>
        public static string RebuildWithValues(string text, string sep1, string sep2, IReadOnlyList<string> newValues)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            if (newValues == null) throw new ArgumentNullException(nameof(newValues));

            var sb = new StringBuilder();
            int searchFrom = 0;
            int idx = 0;

            while (idx < newValues.Count)
            {
                int posStart = text.IndexOf(sep1, searchFrom, StringComparison.Ordinal);
                if (posStart < 0) break;

                int valueStart = posStart + sep1.Length;
                int posEnd = text.IndexOf(sep2, valueStart, StringComparison.Ordinal);
                if (posEnd < 0) break;

                // everything up to and including the opening delimiter, then the new value, then the closing delimiter
                sb.Append(text.Substring(searchFrom, valueStart - searchFrom));
                sb.Append(newValues[idx]);
                sb.Append(sep2);

                searchFrom = posEnd + sep2.Length;
                idx++;
            }

            sb.Append(text.Substring(searchFrom));
            return sb.ToString();
        }

        /// <summary>
        /// Parses the "new value" textbox. "+5" -> add 5, "-3" -> subtract 3, anything else -> literal replace.
        /// A +/- prefix with a non-numeric remainder falls back to a literal replacement (the original VBA would error).
        /// </summary>
        public static FieldChangeOperation ParseNewValue(string input)
        {
            input ??= string.Empty;

            if (input.Length > 1 && (input[0] == '+' || input[0] == '-'))
            {
                string rest = input.Substring(1);
                if (long.TryParse(rest, NumberStyles.Integer, CultureInfo.InvariantCulture, out long delta))
                {
                    return new FieldChangeOperation(
                        input[0] == '+' ? FieldOperation.Add : FieldOperation.Subtract,
                        string.Empty,
                        delta);
                }
            }

            return new FieldChangeOperation(FieldOperation.Replace, input, 0);
        }

        /// <summary>
        /// Applies <paramref name="op"/> to field <paramref name="fieldIndex"/> (0-based) of <paramref name="cellValue"/>.
        /// Mirrors the VBA skip conditions: no separator present -> skipped, index past the last field -> skipped.
        /// Add/Subtract preserve the original whole-number field width by zero-padding (VBA's Format(n, "000...")).
        /// Deliberate deviation from the VBA: a non-integer field (e.g. "1.5") is left UNCHANGED rather than
        /// coerced-and-rounded (the VBA would turn "1.5" +2 into "004", which is lossy). The tool targets integer
        /// address tokens, so this is safer; callers see <see cref="FieldChangeStatus.SkippedNotNumeric"/>.
        /// </summary>
        public static FieldChangeResult ChangeField(
            string cellValue,
            string sep1,
            string sep2,
            int fieldIndex,
            FieldChangeOperation op)
        {
            if (cellValue == null) throw new ArgumentNullException(nameof(cellValue));
            if (op == null) throw new ArgumentNullException(nameof(op));
            if (string.IsNullOrEmpty(sep1)) throw new ArgumentException("sep1 must be non-empty", nameof(sep1));

            bool doubleSeparator = !string.IsNullOrEmpty(sep2);

            // must contain at least one usable field, else skip untouched (matches VBA)
            int firstSep = cellValue.IndexOf(sep1, StringComparison.Ordinal);
            if (firstSep < 0)
                return new FieldChangeResult(FieldChangeStatus.SkippedNoSeparator, cellValue, 0);

            if (doubleSeparator &&
                cellValue.IndexOf(sep2, firstSep + sep1.Length, StringComparison.Ordinal) < 0)
                return new FieldChangeResult(FieldChangeStatus.SkippedNoSeparator, cellValue, 0);

            string[] fields = doubleSeparator
                ? GetValuesBetween(cellValue, sep1, sep2)
                : cellValue.Split(new[] { sep1 }, StringSplitOptions.None);

            if (fieldIndex < 0 || fieldIndex > fields.Length - 1)
                return new FieldChangeResult(FieldChangeStatus.SkippedIndexOutOfRange, cellValue, fields.Length);

            string original = fields[fieldIndex];

            switch (op.Operation)
            {
                case FieldOperation.Replace:
                    fields[fieldIndex] = op.ReplaceValue;
                    break;

                case FieldOperation.Add:
                case FieldOperation.Subtract:
                    if (!long.TryParse(original, NumberStyles.Integer, CultureInfo.InvariantCulture, out long current))
                        return new FieldChangeResult(FieldChangeStatus.SkippedNotNumeric, cellValue, fields.Length);

                    long updated = op.Operation == FieldOperation.Add
                        ? current + op.Delta
                        : current - op.Delta;

                    fields[fieldIndex] = ZeroPad(updated, original.Length);
                    break;
            }

            string rebuilt = doubleSeparator
                ? RebuildWithValues(cellValue, sep1, sep2, fields)
                : string.Join(sep1, fields);

            return new FieldChangeResult(FieldChangeStatus.Changed, rebuilt, fields.Length);
        }

        /// <summary>Zero-pads to at least <paramref name="width"/> digits, keeping a leading '-' for negatives (like VBA Format(n,"000")).</summary>
        public static string ZeroPad(long value, int width)
        {
            bool negative = value < 0;
            string digits = Math.Abs(value).ToString(CultureInfo.InvariantCulture);
            if (digits.Length < width) digits = digits.PadLeft(width, '0');
            return negative ? "-" + digits : digits;
        }
    }
}
