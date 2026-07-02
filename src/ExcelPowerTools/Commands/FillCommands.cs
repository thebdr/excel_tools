using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Text;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Smart fill-down (Ctrl+Shift+D): the top cell of each selected column seeds the fill, its trailing
    /// number increments per visible row (zero-padding preserved: Z07→Z08…Z10; IP last octet →.6 .7);
    /// two seeds infer the step (10, 20 → 30, 40). Hidden/filtered rows are skipped without consuming a
    /// step. Excel's native Ctrl+D remains the plain-copy fill.
    /// </summary>
    public static class FillCommands
    {
        [ExcelCommand(Name = "PT_SmartFillDown", MenuName = "Power Tools", MenuText = "Smart Fill Down", ShortCut = "^+d")]
        public static void SmartFillDown()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            dynamic areas;
            try { areas = app.Selection.Areas; } catch { return; }   // not a cell range

            var undo = UndoManager.NewEntry("Smart Fill Down", sheet);
            using (new ExcelStateGuard(manualCalculation: true))
            {
                int areaCount = (int)areas.Count;
                for (int a = 1; a <= areaCount; a++)
                    FillArea(sheet, areas(a), undo);
            }
            UndoManager.Push(undo);
        }

        private static void FillArea(dynamic sheet, dynamic area, UndoEntry undo)
        {
            int r0 = (int)area.Row;
            int c0 = (int)area.Column;
            int rows = (int)area.Rows.Count;
            int cols = (int)area.Columns.Count;
            if (rows < 2) return;

            object raw = area.Value2;
            if (!(raw is object[,] grid)) return;   // single cell — nothing to fill
            int rBase = grid.GetLowerBound(0);
            int cBase = grid.GetLowerBound(1);

            // Visible row offsets within the area: hidden rows are neither written nor counted.
            var visible = new List<int>();
            for (int r = 0; r < rows; r++)
                if (!RangeHelpers.IsRowHidden(sheet, r0 + r)) visible.Add(r);
            if (visible.Count < 2) return;

            for (int c = 0; c < cols; c++)
            {
                int sheetCol = c0 + c;

                string seed1 = CsvSheet.CellText(grid[rBase + visible[0], cBase + c], sheet, r0 + visible[0], sheetCol);
                if (string.IsNullOrEmpty(seed1)) continue;   // no seed — leave the column alone

                string seed2 = CsvSheet.CellText(grid[rBase + visible[1], cBase + c], sheet, r0 + visible[1], sheetCol);
                var seq = FillSequence.FromSeeds(seed1, string.IsNullOrEmpty(seed2) ? null : seed2);

                int start = seq.TwoSeedsConsumed ? 2 : 1;
                for (int k = start; k < visible.Count; k++)
                {
                    int sheetRow = r0 + visible[k];
                    UndoManager.CaptureCell(undo, sheet, sheetRow, sheetCol);
                    sheet.Cells(sheetRow, sheetCol).Value = seq.ValueAt(k);
                }
            }
        }
    }
}
