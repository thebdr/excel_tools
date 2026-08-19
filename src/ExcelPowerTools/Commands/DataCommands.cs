using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.Text;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Data QA tools over the active cell's contiguous block (CurrentRegion): natural sort (stable,
    /// undo = restore original order) and duplicate-row removal keyed on the selected column(s).
    /// The header row (detected, '#'-comment headers included) and everything above it stay put.
    /// </summary>
    public static class DataCommands
    {
        [ExcelCommand(Name = "PT_ColumnStats", MenuName = "Power Tools", MenuText = "Column Stats...")]
        public static void ShowColumnStats() => ColumnStatsForm.ShowPanel();

        [ExcelCommand(Name = "PT_NaturalSortAsc", MenuName = "Power Tools", MenuText = "Natural Sort (A-Z)")]
        public static void SortAscending() => Sort(ascending: true);

        [ExcelCommand(Name = "PT_NaturalSortDesc", MenuName = "Power Tools", MenuText = "Natural Sort (Z-A)")]
        public static void SortDescending() => Sort(ascending: false);

        // ---- sort ---------------------------------------------------------------------------------

        private static void Sort(bool ascending)
        {
            var b = ReadBlock();
            if (b == null) return;

            int keyOffset = b.ActiveColumn - b.C0;
            if (keyOffset < 0 || keyOffset >= b.DataCols) keyOffset = 0;

            var order = Enumerable.Range(b.DataStart, b.Rows - b.DataStart).ToList();
            var sorted = ascending
                ? order.OrderBy(r => b.Texts[r, keyOffset], NaturalComparer.Instance).ToList()
                : order.OrderByDescending(r => b.Texts[r, keyOffset], NaturalComparer.Instance).ToList();

            if (sorted.SequenceEqual(order))
            {
                Dialogs.Info("Already in order (column " + ColumnStatsForm.ColLetters(b.ActiveColumn) + ", natural).");
                return;
            }

            var undo = UndoManager.NewEntry("Natural Sort", b.Sheet);
            using (new ExcelStateGuard(manualCalculation: true))
            {
                dynamic dataRange = b.DataRange();
                UndoManager.CaptureBlock(undo, dataRange);
                if (b.Session != null && b.RecordMarkerColumn > 0)
                    UndoManager.CaptureCell(undo, b.Sheet, b.R0 + b.Rows, b.RecordMarkerColumn);

                var outGrid = new object[b.Rows - b.DataStart, b.Cols];
                for (int i = 0; i < sorted.Count; i++)
                    for (int c = 0; c < b.Cols; c++)
                        outGrid[i, c] = b.Formulas[sorted[i], c];

                dataRange.FormulaR1C1 = outGrid;

                // Typing below the sentinel expands the logical source without first moving it. Sorting
                // then moves the token with its row, so restore the boundary after the final record.
                if (b.Session != null)
                    CsvSheet.SyncRecordMarkers(b.Sheet, b.Session, b.Rows);
            }
            UndoManager.Push(undo);
        }

        // ---- dedup --------------------------------------------------------------------------------

        [ExcelCommand(Name = "PT_DedupRows", MenuName = "Power Tools", MenuText = "Remove duplicate rows (by key)")]
        public static void DedupRows()
        {
            var b = ReadBlock();
            if (b == null) return;

            List<int> keyCols = SelectedKeyColumns(b);
            string keyDesc = string.Join(", ", keyCols.Select(c => ColumnStatsForm.ColLetters(b.C0 + c)));

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var kept = new List<int>();
            int duplicates = 0;
            for (int r = b.DataStart; r < b.Rows; r++)
            {
                string key = string.Join("", keyCols.Select(c => b.Texts[r, c]));
                if (seen.Add(key)) kept.Add(r);
                else duplicates++;
            }

            if (duplicates == 0)
            {
                Dialogs.Info($"No duplicate rows (key: column(s) {keyDesc}).");
                return;
            }

            if (!Dialogs.Confirm(
                $"Remove {duplicates} duplicate row(s) of {b.Rows - b.DataStart} data row(s)?\n\n" +
                $"Key: column(s) {keyDesc}. The first occurrence is kept; rows shift up within the table.\n" +
                "Undoable (Ctrl+Z / Ctrl+Alt+Z)."))
                return;

            var undo = UndoManager.NewEntry("Remove duplicates", b.Sheet);
            using (new ExcelStateGuard(manualCalculation: true))
            {
                dynamic dataRange = b.DataRange();
                UndoManager.CaptureBlock(undo, dataRange);
                if (b.Session != null && b.RecordMarkerColumn > 0)
                    UndoManager.CaptureCell(undo, b.Sheet, b.R0 + b.Rows, b.RecordMarkerColumn);

                var outGrid = new object[b.Rows - b.DataStart, b.Cols];
                for (int i = 0; i < kept.Count; i++)
                    for (int c = 0; c < b.Cols; c++)
                        outGrid[i, c] = b.Formulas[kept[i], c];
                // rows below the kept ones stay in the grid as empty strings, then get truly cleared
                for (int i = kept.Count; i < b.Rows - b.DataStart; i++)
                    for (int c = 0; c < b.Cols; c++)
                        outGrid[i, c] = "";

                dataRange.FormulaR1C1 = outGrid;

                int firstEmptyRow = b.R0 + b.DataStart + kept.Count;
                int lastRow = b.R0 + b.Rows - 1;
                if (firstEmptyRow <= lastRow)
                    b.Sheet.Range(b.Sheet.Cells(firstEmptyRow, b.C0), b.Sheet.Cells(lastRow, b.C0 + b.Cols - 1)).ClearContents();

                if (b.Session != null)
                    CsvSheet.SyncRecordMarkers(b.Sheet, b.Session, b.DataStart + kept.Count);
            }
            UndoManager.Push(undo);

            Dialogs.Info($"Removed {duplicates} duplicate row(s).");
        }

        // ---- shared block reading -------------------------------------------------------------------

        private sealed class Block
        {
            public dynamic Sheet;
            public int R0, C0, Rows, Cols;
            public int DataCols;         // columns eligible for headers and sort/dedup keys
            public int DataStart;        // block-relative index of the first data row (after the header)
            public int ActiveColumn;     // absolute sheet column of the active cell
            public CsvSession Session;
            public int RecordMarkerColumn = -1;
            public string[,] Formulas;   // what to rewrite with (preserves formulas)
            public string[,] Texts;      // faithful display texts (sort keys, dedup keys, header detect)

            public dynamic DataRange() =>
                Sheet.Range(Sheet.Cells(R0 + DataStart, C0), Sheet.Cells(R0 + Rows - 1, C0 + Cols - 1));
        }

        private static Block ReadBlock()
        {
            dynamic app = ExcelApp.App;
            dynamic active = app.ActiveCell;
            if (active == null) { Dialogs.Info("Select a cell inside the data first."); return null; }

            dynamic sheet = active.Worksheet;
            dynamic region = active.CurrentRegion;
            int r0 = (int)region.Row;
            int c0 = (int)region.Column;
            int rows = (int)region.Rows.Count;
            int cols = (int)region.Columns.Count;
            int dataCols = cols;
            CsvSession csvSession = null;
            int recordMarkerColumn = -1;

            // A source worksheet has hidden logical-record columns outside CurrentRegion. Row operations
            // must move them (and later // helper columns) with the visible values; dedup must also clear
            // their tail, otherwise removed records survive as blank lines when the source is saved.
            int markerColumn, sourceColumns, logicalRows;
            string sourceMarkerToken;
            bool hasWorkspaceLayout = CsvSheet.TryGetWorkspaceLayout(
                (object)sheet, out markerColumn, out sourceColumns, out logicalRows, out sourceMarkerToken);
            csvSession = CsvSession.Get((object)sheet);
            if (csvSession != null && !hasWorkspaceLayout)
            {
                Dialogs.Info(CsvSheet.DamagedBoundaryRecovery);
                return null;
            }
            if (csvSession == null && (hasWorkspaceLayout || CsvSheet.HasWorkspaceMetadataCandidate((object)sheet)))
            {
                Dialogs.Info("This sheet contains CSV workspace metadata but is not attached to an open source session. Reopen the source file before using row tools.");
                return null;
            }
            if (csvSession != null && hasWorkspaceLayout)
            {
                if (logicalRows >= 1048576)
                {
                    Dialogs.Info("This source currently uses Excel's final row, so there is no room for its " +
                                 "record-boundary sentinel. Delete at least one worksheet row before using row tools.");
                    return null;
                }
                int lastColumn = (int)RangeHelpers.LastColumn(sheet);
                r0 = 1;
                c0 = 1;
                rows = logicalRows;
                dataCols = sourceColumns;
                cols = lastColumn;
                recordMarkerColumn = markerColumn;
                if (rows > 0)
                    region = sheet.Range(sheet.Cells(r0, c0), sheet.Cells(r0 + rows - 1, c0 + cols - 1));
            }
            if (rows < 2)
            {
                Dialogs.Info("No data block around the selection (need at least 2 rows).");
                return null;
            }

            var b = new Block
            {
                Sheet = sheet,
                R0 = r0,
                C0 = c0,
                Rows = rows,
                Cols = cols,
                DataCols = dataCols,
                ActiveColumn = (int)active.Column,
                Session = csvSession,
                RecordMarkerColumn = recordMarkerColumn,
            };

            // R1C1 keeps relative formulas relative to their row when records and helper cells move.
            object rawFormulas = region.FormulaR1C1;
            object rawValues = region.Value2;
            b.Formulas = ReadGrid(rawFormulas, rows, cols, null, null, 0, 0);
            b.Texts = ReadGrid(rawValues, rows, cols, new CellTextReader(CsvSheet.CellText), sheet, b.R0, b.C0);

            // Header detection over the block ('#'-comment headers included); data starts below it.
            var rowsList = new List<string[]>(rows);
            for (int r = 0; r < rows; r++)
            {
                var row = new string[dataCols];
                for (int c = 0; c < dataCols; c++) row[c] = b.Texts[r, c];
                rowsList.Add(row);
            }
            int header = CsvDetect.DetectHeaderRow(rowsList, "#");
            b.DataStart = header >= 0 ? header + 1 : 0;

            if (b.Rows - b.DataStart < 2)
            {
                Dialogs.Info("Not enough data rows below the header.");
                return null;
            }
            return b;
        }

        private delegate string CellTextReader(object v, dynamic sheet, int row, int col);

        private static string[,] ReadGrid(object raw, int rows, int cols, CellTextReader toText, dynamic sheet, int r0, int c0)
        {
            var outGrid = new string[rows, cols];
            if (raw is object[,] g)
            {
                int rb = g.GetLowerBound(0), cb = g.GetLowerBound(1);
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                    {
                        object v = g[rb + r, cb + c];
                        outGrid[r, c] = toText != null
                            ? toText(v, sheet, r0 + r, c0 + c)
                            : (v == null ? "" : Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture));
                    }
            }
            else
            {
                outGrid[0, 0] = toText != null
                    ? toText(raw, sheet, r0, c0)
                    : (raw == null ? "" : Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture));
            }
            return outGrid;
        }

        private static List<int> SelectedKeyColumns(Block b)
        {
            var cols = new SortedSet<int>();
            try
            {
                dynamic app = ExcelApp.App;
                dynamic areas = app.Selection.Areas;
                int n = (int)areas.Count;
                for (int a = 1; a <= n; a++)
                {
                    dynamic ar = areas(a);
                    int c0 = (int)ar.Column;
                    int cc = (int)ar.Columns.Count;
                    for (int c = c0; c < c0 + cc; c++)
                        if (c >= b.C0 && c < b.C0 + b.DataCols) cols.Add(c - b.C0);
                }
            }
            catch { }

            if (cols.Count == 0) cols.Add(Math.Max(0, Math.Min(b.ActiveColumn - b.C0, b.DataCols - 1)));
            return cols.ToList();
        }
    }
}
