using System;
using System.Globalization;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core;
using ExcelPowerTools.Interop;
using ExcelPowerTools.Settings;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    /// <summary>Port of IoListAux.ChangeField: edit a delimited field inside each selected cell.</summary>
    public static class FieldEditorCommands
    {
        [ExcelCommand(Name = "PT_ChangeField", MenuName = "Power Tools", MenuText = "Change Field...")]
        public static void ChangeField()
        {
            var settings = FieldEditorSettings.Load();

            using (var form = new FieldEditorForm(settings))
            {
                if (form.ShowDialog(ExcelWindow.Current) != DialogResult.OK)
                    return;

                form.CopyInto(settings);
                settings.Save();

                string sep1 = form.Separator;
                if (string.IsNullOrEmpty(sep1))
                {
                    Dialogs.Info("Please provide at least the start separator.");
                    return;
                }

                string sep2 = form.Separator2;
                int fieldIndex = form.FieldIndex - 1;                 // dialog is 1-based
                var op = FieldStringEditor.ParseNewValue(form.NewValue);
                string fontColor = form.FontColor;
                string interiorColor = form.InteriorColor;

                Apply(sep1, sep2, fieldIndex, op, fontColor, interiorColor);
            }
        }

        private static void Apply(string sep1, string sep2, int fieldIndex,
                                  FieldChangeOperation op, string fontColor, string interiorColor)
        {
            dynamic app = ExcelApp.App;
            dynamic selection = app.Selection;

            dynamic areas;
            try { areas = selection.Areas; }
            catch
            {
                Dialogs.Info("Select a range of cells first.");
                return;
            }

            int changed = 0, outOfRange = 0, notNumeric = 0;
            var undo = UndoManager.NewEntry("Change Field", app.ActiveSheet);

            using (new ExcelStateGuard())
            {
                int areaCount = (int)areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    dynamic area = areas(a);
                    int rows = (int)area.Rows.Count;
                    int cols = (int)area.Columns.Count;

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            dynamic cell = area.Cells(r, c);

                            // Read .Value (not .Value2), matching the VBA: dates/currency come back as the
                            // displayed value and numbers use the user's locale, so text addresses round-trip 1:1.
                            object raw = cell.Value;
                            string text = raw == null ? "" : Convert.ToString(raw, CultureInfo.CurrentCulture);
                            if (text.Length == 0) continue;
                            if ((bool)cell.EntireRow.Hidden) continue;

                            FieldChangeResult result = FieldStringEditor.ChangeField(text, sep1, sep2, fieldIndex, op);

                            switch (result.Status)
                            {
                                case FieldChangeStatus.Changed:
                                    UndoManager.CaptureCell(undo, app.ActiveSheet, (int)cell.Row, (int)cell.Column);
                                    cell.Value = result.NewValue;
                                    ApplyColors(cell, fontColor, interiorColor);
                                    changed++;
                                    break;
                                case FieldChangeStatus.SkippedIndexOutOfRange:
                                    outOfRange++;
                                    break;
                                case FieldChangeStatus.SkippedNotNumeric:
                                    notNumeric++;
                                    break;
                            }
                        }
                    }
                }
            }

            UndoManager.Push(undo);

            if (outOfRange > 0 || notNumeric > 0)
            {
                Dialogs.Info(
                    $"Done. {changed} cell(s) changed." +
                    (outOfRange > 0 ? $"\n{outOfRange} skipped: field index out of range." : "") +
                    (notNumeric > 0 ? $"\n{notNumeric} skipped: field is not a whole number (needed for +/-)." : ""));
            }
        }

        private static void ApplyColors(dynamic cell, string fontColor, string interiorColor)
        {
            if (string.Equals(fontColor, "auto", StringComparison.OrdinalIgnoreCase))
                cell.Font.ColorIndex = Xl.xlColorIndexAutomatic;
            else if (TryParseOleColor(fontColor, out int fc))
                cell.Font.Color = fc;

            if (string.Equals(interiorColor, "none", StringComparison.OrdinalIgnoreCase))
                cell.Interior.ColorIndex = Xl.xlColorIndexNone;
            else if (TryParseOleColor(interiorColor, out int ic))
                cell.Interior.Color = ic;
        }

        /// <summary>Parses "RRGGBB" hex into an Excel OLE colour (BGR-packed). Returns false for blank/invalid.</summary>
        private static bool TryParseOleColor(string hex, out int oleColor)
        {
            oleColor = 0;
            if (string.IsNullOrEmpty(hex) || hex.Length != 6) return false;

            if (int.TryParse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int r) &&
                int.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int g) &&
                int.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int b))
            {
                oleColor = r | (g << 8) | (b << 16);   // Excel expects BGR, same as VBA RGB()
                return true;
            }
            return false;
        }
    }
}
