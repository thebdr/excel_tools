using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// CSV subsystem — built on the round-trip-faithful ExcelPowerTools.Core.Csv reader/writer. The add-in
    /// (not Excel) reads the raw bytes, so values are never coerced, and it writes them back byte-faithfully.
    /// </summary>
    public static class CsvCommands
    {
        private const string OpenFilter = "CSV / text files|*.csv;*.txt|CSV files|*.csv|Text files|*.txt|All files|*.*";
        private const string SaveFilter = "CSV files|*.csv|Text files|*.txt|All files|*.*";

        [ExcelCommand(Name = "PT_OpenCsv", MenuName = "Power Tools", MenuText = "Open CSV/TXT...")]
        public static void OpenCsv()
        {
            string path = Dialogs.OpenFile(OpenFilter);
            if (path == null) return;
            OpenCsvFromPath(path);
        }

        /// <summary>The open pipeline: raw bytes → fingerprint → Text grid → styling → session.</summary>
        public static void OpenCsvFromPath(string path)
        {
            byte[] bytes;
            try { bytes = File.ReadAllBytes(path); }
            catch (Exception ex) { Dialogs.Info("Couldn't read the file:\n" + ex.Message); return; }

            CsvDocument doc = CsvFile.Load(bytes);
            var ragged = CsvDetect.FindRaggedRows(doc.DataRows(), out int expected);

            dynamic app = ExcelApp.App;
            dynamic wb;
            using (new ExcelStateGuard(manualCalculation: true))
            {
                wb = app.Workbooks.Add();
                dynamic sheet = app.ActiveSheet;
                try { sheet.Name = SafeSheetName(Path.GetFileNameWithoutExtension(path)); } catch { }

                CsvSheet.WriteRecords(sheet, doc.Records, out int rows, out int dataWidth);
                int headerSheetRow = doc.HeaderRecordIndex >= 0 ? doc.HeaderRecordIndex + 1 : -1;
                CsvStyler.Apply(app, sheet, headerSheetRow, rows, dataWidth);

                CsvSession.Set((string)wb.Name, new CsvSession
                {
                    SourcePath = path,
                    Format = doc.Format,
                    HeaderSheetRow = headerSheetRow,
                    DataWidth = dataWidth,
                });
            }

            if (ragged.Count > 0)
                Dialogs.Info($"Loaded \"{Path.GetFileName(path)}\".\n\n{ragged.Count} row(s) don't have the expected " +
                             $"{expected} columns (first at data row {ragged[0] + 1}) — likely a stray delimiter or " +
                             "quote. They're loaded as-is so you can spot and fix them.");
        }

        [ExcelCommand(Name = "PT_SaveCsv", MenuName = "Power Tools", MenuText = "Save CSV")]
        public static void SaveCsv()
        {
            dynamic app = ExcelApp.App;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            var session = CsvSession.Get((string)wb.Name);
            if (session == null || string.IsNullOrEmpty(session.SourcePath))
            {
                // No remembered CSV for this workbook — behave like "Export as CSV".
                ExportSheetAsCsv();
                return;
            }

            dynamic sheet = app.ActiveSheet;
            var doc = new CsvDocument
            {
                Format = session.Format,
                Records = CsvSheet.ReadRecords(sheet, session.DataWidth),
            };

            try
            {
                WriteWithBackup(session.SourcePath, doc.Save());
                Dialogs.Info("Saved:\n" + session.SourcePath);
            }
            catch (Exception ex)
            {
                Dialogs.Info("Save failed:\n" + ex.Message);
            }
        }

        [ExcelCommand(Name = "PT_ExportCsv", MenuName = "Power Tools", MenuText = "Export sheet as CSV...")]
        public static void ExportSheetAsCsv()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            char separator;
            using (var form = new CsvSeparatorForm("Export sheet as CSV", LastSeparatorToken()))
            {
                if (form.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return;
                SetLastSeparatorToken(form.RawToken);
                separator = CsvSeparatorForm.Translate(form.RawToken);
            }

            string defaultName = SafeFileName((string)sheet.Name) + ".csv";
            string path = Dialogs.SaveFile(SaveFilter, defaultName);
            if (path == null) return;

            var fmt = new CsvFormat
            {
                Delimiter = separator,
                Encoding = new UTF8Encoding(false),
                HasBom = false,
                LineEnding = CsvLineEnding.Crlf,
                HasTrailingNewline = true,
            };
            var doc = new CsvDocument { Format = fmt, Records = CsvSheet.ReadRecords(sheet, 0) };

            try
            {
                WriteWithBackup(path, doc.Save());
                Dialogs.Info("Exported to:\n" + path);
            }
            catch (Exception ex)
            {
                Dialogs.Info("Export failed:\n" + ex.Message);
            }
        }

        [ExcelCommand(Name = "PT_RestyleCsv", MenuName = "Power Tools", MenuText = "Re-style CSV sheet")]
        public static void RestyleActiveSheet()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);

            var records = CsvSheet.ReadRecords(sheet, 0);
            var rows = new List<string[]>();
            foreach (var r in records) rows.Add(r.Fields);

            int headerIndex = CsvDetect.DetectHeaderRow(rows, "#");
            int headerSheetRow = headerIndex >= 0 ? headerIndex + 1 : -1;   // records map 1:1 to sheet rows

            using (new ExcelStateGuard())
                CsvStyler.Apply(app, sheet, headerSheetRow, lastRow, lastCol);
        }

        [ExcelCommand(Name = "PT_BackupCsv", MenuName = "Power Tools", MenuText = "Backup CSV file")]
        public static void BackupActiveFile()
        {
            dynamic app = ExcelApp.App;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            var session = CsvSession.Get((string)wb.Name);
            if (session == null || string.IsNullOrEmpty(session.SourcePath) || !File.Exists(session.SourcePath))
            {
                Dialogs.Info("No CSV file is associated with this workbook. Open one with \"Open CSV/TXT\".");
                return;
            }

            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string dir = Path.GetDirectoryName(session.SourcePath);
            string name = Path.GetFileNameWithoutExtension(session.SourcePath);
            string ext = Path.GetExtension(session.SourcePath);
            string backup = Path.Combine(dir, $"_{name} ({stamp}){ext}");

            try
            {
                File.Copy(session.SourcePath, backup, overwrite: false);
                Dialogs.Info("Backup created:\n" + backup);
            }
            catch (Exception ex)
            {
                Dialogs.Info("Backup failed:\n" + ex.Message);
            }
        }

        // ---- open interception (wired in AddIn.AutoOpen) -----------------------------------------

        /// <summary>
        /// Application.WorkbookOpen handler: when Excel natively opens a .csv (double-click in Explorer,
        /// File▸Open, ...), discard Excel's already-coerced parse and re-open the untouched disk file
        /// through our pipeline. Deferred via QueueAsMacro — closing a workbook inside its own Open event
        /// is not safe.
        /// </summary>
        internal static void OnWorkbookOpen(Excel.Workbook wb)
        {
            try
            {
                if (!InterceptOpenEnabled()) return;

                string path = wb?.FullName;
                if (string.IsNullOrEmpty(path) || !path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)) return;
                if (!File.Exists(path)) return;

                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        wb.Close(false);          // drop the native parse; the file on disk is untouched
                        OpenCsvFromPath(path);
                    }
                    catch (Exception ex)
                    {
                        Dialogs.Info("CSV open interception failed:\n" + ex.Message);
                    }
                });
            }
            catch { /* never let a handler exception break the user's open */ }
        }

        public static bool InterceptOpenEnabled()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(SettingsKey))
                return (k?.GetValue("InterceptOpen") as string) != "0";   // default: on
        }

        public static void SetInterceptOpen(bool enabled)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(SettingsKey))
                k?.SetValue("InterceptOpen", enabled ? "1" : "0");
        }

        // ---- helpers ----------------------------------------------------------------------------

        /// <summary>Writes bytes with a one-level ".bak" backup and an atomic replace (no partial file on crash).</summary>
        private static void WriteWithBackup(string path, byte[] bytes)
        {
            string tmp = path + ".tmp";
            try
            {
                File.WriteAllBytes(tmp, bytes);
                if (File.Exists(path))
                {
                    string bak = path + ".bak";
                    try { if (File.Exists(bak)) File.SetAttributes(bak, FileAttributes.Normal); } catch { }
                    File.Replace(tmp, path, bak);   // atomic; original stays intact on failure
                }
                else
                {
                    File.Move(tmp, path);
                }
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }   // never leave a stray .tmp
            }
        }

        private static string SafeSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "CSV";
            foreach (char c in new[] { '[', ']', ':', '*', '?', '/', '\\' }) name = name.Replace(c, '_');
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private static string SafeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "export";
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            return name;
        }

        private const string SettingsKey = @"Software\ExcelPowerTools\Csv";

        private static string LastSeparatorToken()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(SettingsKey))
                return (k?.GetValue("LastSeparator") as string) ?? ";";
        }

        private static void SetLastSeparatorToken(string token)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(SettingsKey))
                k?.SetValue("LastSeparator", token ?? ";");
        }
    }
}
