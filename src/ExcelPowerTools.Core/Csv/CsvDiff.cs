using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>What changed between two record sets (e.g. the file on disk vs the grid about to be saved).</summary>
    public sealed class CsvDiffSummary
    {
        public int ChangedCells { get; set; }
        public int AddedRows { get; set; }
        public int RemovedRows { get; set; }

        /// <summary>1-based position of the first differing cell (row = record index + 1), or -1.</summary>
        public int FirstChangeRow { get; set; } = -1;
        public int FirstChangeColumn { get; set; } = -1;

        public bool HasChanges => ChangedCells > 0 || AddedRows > 0 || RemovedRows > 0;

        public string Describe()
        {
            if (!HasChanges) return "no value changes";

            var sb = new StringBuilder();
            if (ChangedCells > 0)
            {
                sb.Append($"{ChangedCells} cell(s) changed");
                if (FirstChangeRow > 0) sb.Append($" (first at row {FirstChangeRow}, column {FirstChangeColumn})");
            }
            if (AddedRows > 0) sb.Append(sb.Length > 0 ? ", " : "").Append($"{AddedRows} row(s) added");
            if (RemovedRows > 0) sb.Append(sb.Length > 0 ? ", " : "").Append($"{RemovedRows} row(s) removed");
            return sb.ToString();
        }
    }

    /// <summary>Record-level comparison and the pre-save round-trip self-check.</summary>
    public static class CsvDiff
    {
        /// <summary>Compares two record sets row-by-row (extra rows in <paramref name="current"/> are additions).</summary>
        public static CsvDiffSummary Summarize(IReadOnlyList<CsvRecord> original, IReadOnlyList<CsvRecord> current)
        {
            var s = new CsvDiffSummary();
            original = original ?? new List<CsvRecord>();
            current = current ?? new List<CsvRecord>();

            int common = original.Count < current.Count ? original.Count : current.Count;
            for (int r = 0; r < common; r++)
            {
                string[] a = original[r].Fields;
                string[] b = current[r].Fields;
                int commonWidth = a.Length < b.Length ? a.Length : b.Length;
                for (int c = 0; c < commonWidth; c++)
                {
                    string va = a[c] ?? "";
                    string vb = b[c] ?? "";
                    if (va != vb)
                    {
                        s.ChangedCells++;
                        if (s.FirstChangeRow < 0) { s.FirstChangeRow = r + 1; s.FirstChangeColumn = c + 1; }
                    }
                }
                if (a.Length != b.Length)
                {
                    s.ChangedCells += Math.Abs(a.Length - b.Length);
                    if (s.FirstChangeRow < 0)
                    {
                        s.FirstChangeRow = r + 1;
                        s.FirstChangeColumn = commonWidth + 1;
                    }
                }
            }

            s.AddedRows = current.Count > common ? current.Count - common : 0;
            s.RemovedRows = original.Count > common ? original.Count - common : 0;
            return s;
        }

        /// <summary>
        /// The last line of defense before writing: re-parses <paramref name="bytesToWrite"/> with the same
        /// format and verifies every field survives exactly. Returns null when everything matches, else a
        /// human-readable description of the first mismatch — in which case the file must NOT be written.
        /// </summary>
        public static string VerifyRoundTrip(byte[] bytesToWrite, CsvFormat format, IReadOnlyList<CsvRecord> intended)
        {
            var reparsed = CsvFile.Load(bytesToWrite, format);
            intended = intended ?? new List<CsvRecord>();

            if (reparsed.Records.Count != intended.Count)
                return $"row count after re-parse is {reparsed.Records.Count}, expected {intended.Count}";

            for (int r = 0; r < intended.Count; r++)
            {
                string[] a = intended[r].Fields;
                string[] b = reparsed.Records[r].Fields;
                if (a.Length != b.Length)
                    return $"row {r + 1}: field count after re-parse is {b.Length}, expected {a.Length}";

                for (int c = 0; c < a.Length; c++)
                {
                    string va = a[c] ?? "";
                    string vb = b[c] ?? "";
                    if (va != vb)
                        return $"row {r + 1}, column {c + 1}: value would change from \"{va}\" to \"{vb}\"";
                }
            }

            return null;
        }
    }
}
