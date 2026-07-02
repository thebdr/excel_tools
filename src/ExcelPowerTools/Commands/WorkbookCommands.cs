using System;
using System.IO;
using ExcelDna.Integration;
using ExcelPowerTools.Interop;
using VB = Microsoft.VisualBasic;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Port of aPlayground.exportSheetsAsNewFile, generalised: split the active workbook into several
    /// files, each holding a fixed number of consecutive sheets. (The VBA version had a hard-coded path.)
    /// </summary>
    public static class WorkbookCommands
    {
        [ExcelCommand(Name = "PT_SplitWorkbook", MenuName = "Power Tools", MenuText = "Split workbook into files...")]
        public static void ExportSheetsAsFiles()
        {
            dynamic app = ExcelApp.App;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            string fullName = (string)wb.FullName;
            if (string.IsNullOrEmpty(fullName) || !File.Exists(fullName))
            {
                Dialogs.Info("Save the workbook to disk first, then split it.");
                return;
            }

            string answer = VB.Interaction.InputBox(
                "How many sheets per output file?", Dialogs.Title, "5");
            if (string.IsNullOrEmpty(answer)) return;
            if (!int.TryParse(answer.Trim(), out int perFile) || perFile < 1)
            {
                Dialogs.Info("Please enter a whole number of at least 1.");
                return;
            }

            int totalSheets = (int)wb.Sheets.Count;
            int fileCount = (int)Math.Ceiling(totalSheets / (double)perFile);
            string wbName = Path.GetFileName(fullName);
            string ext = Path.GetExtension(fullName);
            int created = 0;

            using (new ExcelStateGuard(suppressEvents: true))
            {
                for (int i = 0; i < fileCount; i++)
                {
                    int firstSheet = i * perFile + 1;
                    int lastSheet = firstSheet + perFile - 1;
                    if (lastSheet > totalSheets) lastSheet = totalSheets;

                    string chunkName = Path.GetFileNameWithoutExtension(wbName) +
                                       $" sheets_{firstSheet}-{lastSheet}{ext}";
                    string chunkPath = Path.Combine(Path.GetDirectoryName(fullName), chunkName);

                    try
                    {
                        File.Copy(fullName, chunkPath, overwrite: true);

                        dynamic dummy = app.Workbooks.Open(chunkPath);
                        bool prevAlerts = app.DisplayAlerts;
                        app.DisplayAlerts = false;
                        try
                        {
                            for (int s = (int)dummy.Sheets.Count; s >= 1; s--)
                            {
                                if (s < firstSheet || s > lastSheet)
                                    dummy.Sheets(s).Delete();
                            }
                        }
                        finally { app.DisplayAlerts = prevAlerts; }

                        dummy.Close(true);   // save changes
                        created++;
                    }
                    catch (Exception ex)
                    {
                        Dialogs.Info($"Failed writing '{chunkName}':\n{ex.Message}");
                    }
                }
            }

            Dialogs.Info($"Done. {created} file(s) created next to the original.");
        }
    }
}
