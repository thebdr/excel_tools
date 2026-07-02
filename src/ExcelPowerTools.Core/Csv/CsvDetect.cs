using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>Heuristic detection helpers used when loading a CSV.</summary>
    public static class CsvDetect
    {
        public static readonly char[] DefaultDelimiters = { ',', ';', '\t', '|' };

        private static readonly Regex NumericLike =
            new Regex(@"^[+-]?[\d]{1,3}([.,\s]?\d{3})*([.,]\d+)?([eE][+-]?\d+)?$|^[+-]?[.,]?\d+([.,]\d+)?$",
                RegexOptions.Compiled);

        /// <summary>
        /// Picks the delimiter that yields the most consistent column count across the sample (high mean
        /// field count, low variance). Quote-aware because it parses with the real reader for each candidate.
        /// </summary>
        public static char DetectDelimiter(string sample, char[] candidates, char quote, char fallback)
        {
            if (string.IsNullOrEmpty(sample)) return fallback;
            candidates = candidates ?? DefaultDelimiters;

            char best = fallback;
            double bestScore = double.NegativeInfinity;
            bool found = false;

            foreach (char d in candidates)
            {
                var counts = FieldCountsPerRecord(sample, d, quote, "#");
                if (counts.Count == 0) continue;

                double mean = Average(counts);
                if (mean <= 1.0 + 1e-9) continue;   // delimiter effectively absent

                double variance = Variance(counts, mean);
                double score = mean - 2.0 * variance;   // reward columns, punish ragged inconsistency
                if (score > bestScore)
                {
                    bestScore = score;
                    best = d;
                    found = true;
                }
            }

            return found ? best : fallback;
        }

        /// <summary>
        /// Counts fields per record for a candidate delimiter using quote-TOGGLE semantics (any quote flips
        /// in/out of a quoted region). This protects quoted regions regardless of field position, so the
        /// score of a wrong delimiter isn't inflated by delimiters that actually sit inside quotes.
        /// </summary>
        private static List<int> FieldCountsPerRecord(string text, char delim, char quote, string commentPrefix)
        {
            var counts = new List<int>();
            int i = 0, n = text.Length;
            while (i < n)
            {
                if (!string.IsNullOrEmpty(commentPrefix) && StartsWithAt(text, i, commentPrefix))
                {
                    while (i < n && text[i] != '\r' && text[i] != '\n') i++;
                    i = SkipEol(text, i);
                    continue;
                }

                int delims = 0;
                bool inQuotes = false, ended = false;
                while (i < n && !ended)
                {
                    char c = text[i];
                    if (c == quote) { inQuotes = !inQuotes; i++; }
                    else if (!inQuotes && c == delim) { delims++; i++; }
                    else if (!inQuotes && (c == '\r' || c == '\n')) { i = SkipEol(text, i); ended = true; }
                    else i++;
                }
                counts.Add(delims + 1);
            }
            return counts;
        }

        private static bool StartsWithAt(string s, int i, string prefix)
        {
            if (i + prefix.Length > s.Length) return false;
            for (int k = 0; k < prefix.Length; k++)
                if (s[i + k] != prefix[k]) return false;
            return true;
        }

        private static int SkipEol(string s, int i)
        {
            if (i < s.Length && s[i] == '\r') { i++; if (i < s.Length && s[i] == '\n') i++; }
            else if (i < s.Length && s[i] == '\n') i++;
            return i;
        }

        public static CsvLineEnding DetectLineEnding(string text)
        {
            int crlf = 0, lf = 0, cr = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\r')
                {
                    if (i + 1 < text.Length && text[i + 1] == '\n') { crlf++; i++; }
                    else cr++;
                }
                else if (text[i] == '\n') lf++;
            }

            if (crlf >= lf && crlf >= cr) return CsvLineEnding.Crlf;   // also the default when there are no newlines
            if (lf >= cr) return CsvLineEnding.Lf;
            return CsvLineEnding.Cr;
        }

        public static bool HasTrailingNewline(string text) =>
            !string.IsNullOrEmpty(text) &&
            (text[text.Length - 1] == '\n' || text[text.Length - 1] == '\r');

        /// <summary>
        /// True when a row is a comment row: its FIRST field starts with the comment prefix. Comment rows
        /// are ordinary records (split, edited, saved like data) — this only drives detection-skipping,
        /// save-side trailing-trim, and the green styling in the Excel layer.
        /// </summary>
        public static bool IsCommentFields(string[] fields, string commentPrefix)
        {
            return !string.IsNullOrEmpty(commentPrefix)
                && fields != null && fields.Length > 0
                && fields[0] != null
                && fields[0].StartsWith(commentPrefix, StringComparison.Ordinal);
        }

        /// <summary>Parses an Excel "sep=&lt;char&gt;" directive line; returns false if the line isn't one.</summary>
        public static bool TryParseSepDirective(string line, out char separator)
        {
            separator = '\0';
            if (line == null || line.Length < 5) return false;
            if (!line.Substring(0, 4).Equals("sep=", StringComparison.OrdinalIgnoreCase)) return false;

            string rest = line.Substring(4);
            if (rest.Length != 1) return false;   // Excel writes a single literal delimiter char
            separator = rest[0];
            return true;
        }

        /// <summary>
        /// First "wordy", mostly-filled row whose column count matches the modal DATA width — the header.
        /// Comment rows are candidates too (a commented header like "#Name;IP;Slot" is the norm in the
        /// hardware-config files), but the modal width is computed from data rows only, and mostly-empty
        /// rows (e.g. a padded "#!format=2,,,," directive) are rejected.
        /// Returns the index into <paramref name="rows"/>, or -1 if none looks like a header.
        /// </summary>
        public static int DetectHeaderRow(IReadOnlyList<string[]> rows, string commentPrefix, int maxScan = 25)
        {
            if (rows == null || rows.Count == 0) return -1;

            var dataRows = new List<string[]>();
            foreach (var r in rows)
                if (!IsCommentFields(r, commentPrefix)) dataRows.Add(r);

            int modal = ModalLength(dataRows.Count > 0 ? dataRows : rows);
            if (modal <= 0) return -1;

            int minFilled = Math.Max(1, (modal + 1) / 2);   // reject mostly-empty rows (padded directives)

            int limit = Math.Min(rows.Count, maxScan);
            for (int i = 0; i < limit; i++)
            {
                string[] row = rows[i];
                if (row.Length != modal) continue;
                if (NonEmptyCount(row) < minFilled) continue;
                if (IsWordyRow(row)) return i;
            }
            return -1;
        }

        private static int NonEmptyCount(string[] row)
        {
            int n = 0;
            foreach (string cell in row)
                if (!string.IsNullOrWhiteSpace(cell)) n++;
            return n;
        }

        /// <summary>
        /// Rows whose field count differs from the modal (expected) width. Returns indices into
        /// <paramref name="dataRows"/> and reports the expected width.
        /// </summary>
        public static IReadOnlyList<int> FindRaggedRows(IReadOnlyList<string[]> dataRows, out int expected)
        {
            var ragged = new List<int>();
            expected = dataRows == null || dataRows.Count == 0 ? 0 : ModalLength(dataRows);
            if (dataRows == null) return ragged;

            for (int i = 0; i < dataRows.Count; i++)
                if (dataRows[i].Length != expected) ragged.Add(i);

            return ragged;
        }

        // ---- helpers -----------------------------------------------------------------------------

        private static bool IsWordyRow(string[] row)
        {
            int nonEmpty = 0, wordy = 0;
            foreach (string cell in row)
            {
                if (string.IsNullOrWhiteSpace(cell)) continue;
                nonEmpty++;
                if (!LooksNumeric(cell)) wordy++;
            }
            if (nonEmpty == 0) return false;
            return wordy >= Math.Max(1, (int)Math.Ceiling(nonEmpty * 0.6));
        }

        public static bool LooksNumeric(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            return NumericLike.IsMatch(s.Trim());
        }

        private static int ModalLength(IReadOnlyList<string[]> rows)
        {
            var freq = new Dictionary<int, int>();
            var firstSeen = new Dictionary<int, int>();
            int order = 0;
            foreach (var r in rows)
            {
                int len = r.Length;
                if (!firstSeen.ContainsKey(len)) firstSeen[len] = order++;
                freq.TryGetValue(len, out int c);
                freq[len] = c + 1;
            }

            int best = 0, bestCount = -1, bestOrder = int.MaxValue;
            foreach (var kv in freq)
            {
                // Most frequent width wins; ties break toward the width seen first (stable, avoids
                // flagging leading rows as ragged and avoids hiding a narrow header behind a wide body).
                if (kv.Value > bestCount || (kv.Value == bestCount && firstSeen[kv.Key] < bestOrder))
                {
                    best = kv.Key;
                    bestCount = kv.Value;
                    bestOrder = firstSeen[kv.Key];
                }
            }
            return best;
        }

        private static double Average(List<int> xs)
        {
            long sum = 0;
            foreach (int x in xs) sum += x;
            return (double)sum / xs.Count;
        }

        private static double Variance(List<int> xs, double mean)
        {
            double sum = 0;
            foreach (int x in xs) { double d = x - mean; sum += d * d; }
            return sum / xs.Count;
        }
    }
}
