using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Text;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;
using Microsoft.Win32;

namespace ExcelPowerTools.Commands
{
    /// <summary>Regex find &amp; replace over the selection or the whole sheet, with preview.</summary>
    public static class RegexReplaceCommands
    {
        private const string SettingsKey = @"Software\ExcelPowerTools\RegexReplace";

        [ExcelCommand(Name = "PT_RegexReplace", MenuName = "Power Tools", MenuText = "Regex Replace...")]
        public static void ShowRegexReplace()
        {
            dynamic app = ExcelApp.App;

            bool preferSelection = false;
            try { preferSelection = Convert.ToDouble(app.Selection.CountLarge) > 1.0; } catch { }

            using (var form = new RegexReplaceForm(LoadSpec(), preferSelection))
            {
                form.ComputeChanges = Compute;
                form.ApplyChanges = Apply;
                form.ShowDialog(ExcelWindow.Current);
                SaveSpec(form.CurrentSpec);
            }
        }

        // ---- scan -------------------------------------------------------------------------------

        private static List<RegexCellChange> Compute(RegexReplaceSpec spec, bool selectionScope, bool visibleOnly)
        {
            var changes = new List<RegexCellChange>();
            var seen = new HashSet<long>();   // dedupe overlapping selection areas

            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return changes;

            var regex = RegexReplacer.Build(spec);
            string replacement = RegexReplacer.EffectiveReplacement(spec);

            var areas = new List<dynamic>();
            if (selectionScope)
            {
                dynamic selection = app.Selection;
                dynamic selAreas;
                try { selAreas = selection.Areas; } catch { return changes; }

                int count = (int)selAreas.Count;
                for (int a = 1; a <= count; a++)
                {
                    // Bound whole-column/row selections to the used range so we never scan a million rows.
                    dynamic clipped = app.Intersect(selAreas(a), sheet.UsedRange);
                    if (clipped != null) areas.Add(clipped);
                }
            }
            else
            {
                areas.Add(sheet.UsedRange);
            }

            foreach (dynamic area in areas)
            {
                int r0 = (int)area.Row;
                int c0 = (int)area.Column;
                int rows = (int)area.Rows.Count;
                int cols = (int)area.Columns.Count;

                object raw = area.Value2;
                object[,] grid;
                int rBase, cBase;
                if (raw is object[,] g) { grid = g; rBase = g.GetLowerBound(0); cBase = g.GetLowerBound(1); }
                else { grid = new object[1, 1]; grid[0, 0] = raw; rBase = 0; cBase = 0; rows = 1; cols = 1; }

                for (int r = 0; r < rows; r++)
                {
                    int sheetRow = r0 + r;
                    if (visibleOnly && IsRowHidden(sheet, sheetRow)) continue;

                    for (int c = 0; c < cols; c++)
                    {
                        int sheetCol = c0 + c;
                        long key = (long)sheetRow << 20 | (uint)sheetCol;
                        if (!seen.Add(key)) continue;

                        string text = CsvSheet.CellText(grid[rBase + r, cBase + c], sheet, sheetRow, sheetCol);
                        if (RegexReplacer.TryReplace(regex, replacement, text, out string after))
                            changes.Add(new RegexCellChange { Row = sheetRow, Column = sheetCol, Before = text, After = after });
                    }
                }
            }

            return changes;
        }

        private static bool IsRowHidden(dynamic sheet, int row)
        {
            try { return (bool)sheet.Rows(row).Hidden; } catch { return false; }
        }

        // ---- apply ------------------------------------------------------------------------------

        private static void Apply(List<RegexCellChange> changes)
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            var undo = UndoManager.NewEntry("Regex Replace", sheet);
            using (new ExcelStateGuard(manualCalculation: true))
            {
                foreach (var ch in changes)
                {
                    UndoManager.CaptureCell(undo, sheet, ch.Row, ch.Column);
                    sheet.Cells(ch.Row, ch.Column).Value = ch.After;
                }
            }
            UndoManager.Push(undo);
        }

        // ---- settings ---------------------------------------------------------------------------

        private static RegexReplaceSpec LoadSpec()
        {
            var s = new RegexReplaceSpec();
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(SettingsKey))
            {
                if (k == null) return s;
                s.Pattern = (k.GetValue("Pattern") as string) ?? "";
                s.Replacement = (k.GetValue("Replacement") as string) ?? "";
                s.IgnoreCase = (k.GetValue("IgnoreCase") as string) == "1";
                s.LiteralPattern = (k.GetValue("Literal") as string) == "1";
                s.WholeCell = (k.GetValue("WholeCell") as string) == "1";
            }
            return s;
        }

        private static void SaveSpec(RegexReplaceSpec s)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(SettingsKey))
            {
                if (k == null) return;
                k.SetValue("Pattern", s.Pattern ?? "");
                k.SetValue("Replacement", s.Replacement ?? "");
                k.SetValue("IgnoreCase", s.IgnoreCase ? "1" : "0");
                k.SetValue("Literal", s.LiteralPattern ? "1" : "0");
                k.SetValue("WholeCell", s.WholeCell ? "1" : "0");
            }
        }
    }
}
