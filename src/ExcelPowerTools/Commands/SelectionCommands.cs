using ExcelDna.Integration;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Commands
{
    /// <summary>Port of PasteMultiSelectionValues.KeepValuesOnlyV2 (Ctrl+Q).</summary>
    public static class SelectionCommands
    {
        /// <summary>
        /// Replaces formulas with their computed values, across every selected area, but only for
        /// visible (not filtered/hidden) cells — the same behaviour as the original Ctrl+Q macro.
        /// </summary>
        [ExcelCommand(Name = "PT_KeepValuesOnly", MenuName = "Power Tools", MenuText = "Keep Values Only", ShortCut = "^q")]
        public static void KeepValuesOnly()
        {
            dynamic app = ExcelApp.App;
            dynamic selection = app.Selection;

            dynamic selectionAreas;
            try { selectionAreas = selection.Areas; }
            catch { return; }   // selection is not a range (chart, picture, nothing)

            using (new ExcelStateGuard())
            {
                // Count all selected cells across areas (Selection.Cells.Count is unreliable for multi-area).
                long totalCells = 0;
                int selAreaCount = (int)selectionAreas.Count;
                for (int a = 1; a <= selAreaCount; a++)
                    totalCells += (long)(int)selectionAreas(a).Cells.Count;

                if (totalCells == 0) return;

                dynamic target = selection;
                if (totalCells > 1)
                {
                    try { target = selection.SpecialCells(Xl.xlCellTypeVisible); }
                    catch { return; }   // no visible cells
                }

                dynamic areas = target.Areas;
                int areaCount = (int)areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    dynamic area = areas(a);
                    // Reading Value2 returns the computed values; writing them back drops the formulas.
                    object values = area.Value2;
                    area.Value2 = values;
                }
            }
        }
    }
}
