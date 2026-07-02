using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Commands
{
    /// <summary>Bridges <see cref="CsvRecord"/>s and an Excel worksheet (COM). Never lets Excel coerce values.</summary>
    internal static class CsvSheet
    {
        /// <summary>
        /// Writes records to <paramref name="sheet"/> starting at A1 — every row (comment rows included) is
        /// split across columns. The ENTIRE sheet is formatted as Text first, so neither the written values
        /// nor anything the user later types/pastes is coerced (leading zeros, long IDs, dates, formulas all
        /// stay literal). <paramref name="dataWidth"/> is the widest record — remembered so Save re-emits the
        /// same column count.
        /// </summary>
        public static void WriteRecords(dynamic sheet, IReadOnlyList<CsvRecord> records, out int rowCount, out int dataWidth)
        {
            rowCount = records.Count;
            dataWidth = 1;
            foreach (var r in records)
                if (r.FieldCount > dataWidth) dataWidth = r.FieldCount;

            sheet.Cells.NumberFormat = "@";   // Text — the whole sheet, before any value lands
            if (rowCount == 0) return;

            var data = new object[rowCount, dataWidth];
            for (int r = 0; r < rowCount; r++)
            {
                CsvRecord rec = records[r];
                for (int c = 0; c < dataWidth; c++)
                    data[r, c] = c < rec.FieldCount ? (object)rec.Fields[c] : null;
            }

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(rowCount, dataWidth));
            range.Value = data;
        }

        /// <summary>
        /// Reads a worksheet back into records. Non-string cells are read via their displayed <c>.Text</c> so
        /// error values (#REF!) and formatted numbers survive; doubles are stringified invariant/integer-aware
        /// (no locale separator, no scientific notation). Every row — comment rows included — is emitted at
        /// <paramref name="forcedWidth"/> (or the widest non-empty extent), matching how Excel itself pads
        /// short rows when saving a CSV.
        /// </summary>
        public static List<CsvRecord> ReadRecords(dynamic sheet, int forcedWidth)
        {
            var records = new List<CsvRecord>();

            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);
            if (lastRow < 1) return records;

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            object raw = range.Value2;

            object[,] grid;
            int rows, cols, rBase, cBase;
            if (raw is object[,] g)
            {
                grid = g;
                rows = g.GetLength(0); cols = g.GetLength(1);
                rBase = g.GetLowerBound(0); cBase = g.GetLowerBound(1);
            }
            else
            {
                grid = new object[1, 1]; grid[0, 0] = raw;
                rows = 1; cols = 1; rBase = 0; cBase = 0;
            }

            // Never below the widest non-empty extent, so user-added columns are kept, not truncated.
            int width = Math.Max(forcedWidth, EffectiveWidth(grid, rows, cols, rBase, cBase));
            if (width < 1) width = 1;

            for (int r = 0; r < rows; r++)
            {
                int sheetRow = r + 1;
                var row = new string[cols];
                for (int c = 0; c < cols; c++)
                    row[c] = CellText(grid[rBase + r, cBase + c], sheet, sheetRow, c + 1);

                records.Add(CsvRecord.Data(Resize(row, width)));
            }

            return records;
        }

        private static string[] Resize(string[] row, int width)
        {
            if (row.Length == width) return row;
            var outRow = new string[width];
            for (int c = 0; c < width; c++)
                outRow[c] = c < row.Length ? row[c] : "";
            return outRow;
        }

        private static int EffectiveWidth(object[,] grid, int rows, int cols, int rBase, int cBase)
        {
            int width = 1;
            for (int r = 0; r < rows; r++)
            {
                int lastNonEmpty = 0;
                for (int c = 0; c < cols; c++)
                    if (!IsEmpty(grid[rBase + r, cBase + c])) lastNonEmpty = c + 1;
                if (lastNonEmpty > width) width = lastNonEmpty;
            }
            return width;
        }

        private static bool IsEmpty(object v) => v == null || (v is string s && s.Length == 0);

        /// <summary>Faithful cell text: exact for strings, displayed .Text for errors/formatted cells, invariant for doubles.</summary>
        private static string CellText(object v, dynamic sheet, int sheetRow, int sheetCol)
        {
            if (v == null) return "";
            if (v is string s) return s;
            if (v is bool b) return b ? "TRUE" : "FALSE";
            if (v is double d) return FormatDouble(d);

            // CVErr (an Excel error like #REF!) or an unexpected type — use the displayed value.
            try { object t = sheet.Cells(sheetRow, sheetCol).Text; return t == null ? "" : t.ToString(); }
            catch { return Convert.ToString(v, CultureInfo.InvariantCulture); }
        }

        private static string FormatDouble(double d)
        {
            if (!double.IsInfinity(d) && !double.IsNaN(d) && d == Math.Floor(d) && Math.Abs(d) < 9.0e18)
                return ((long)d).ToString(CultureInfo.InvariantCulture);           // integer, no exponent
            return d.ToString("0.###############", CultureInfo.InvariantCulture);   // no locale sep, no scientific
        }
    }
}
