using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.Files;
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
        private const string OpenFilter = "Delimited files|*.csv;*.tsv;*.txt|CSV files|*.csv|TSV files|*.tsv|Text files|*.txt";
        private const string SaveFilter = "Delimited files|*.csv;*.tsv;*.txt|CSV files|*.csv|TSV files|*.tsv|Text files|*.txt|All files|*.*";
        private static readonly HashSet<string> PendingIndexRefreshes =
            new HashSet<string>(StringComparer.Ordinal);

        [ExcelCommand(Name = "PT_OpenCsv", MenuName = "Power Tools", MenuText = "Open delimited files...")]
        public static void OpenCsv()
        {
            string[] paths = Dialogs.OpenFiles(OpenFilter);
            if (paths == null || paths.Length == 0) return;
            OpenPaths(paths, showIndexWhenDone: paths.Length > 1);
        }

        [ExcelCommand(Name = "PT_OpenCsvFolder", MenuName = "Power Tools", MenuText = "Open delimited folder...")]
        public static void OpenCsvFolder()
        {
            string folder = Dialogs.SelectFolder("Open every CSV, TSV and separated-value TXT file in this folder.");
            if (folder == null) return;

            string[] paths;
            try
            {
                paths = Directory.EnumerateFiles(folder, "*", SearchOption.TopDirectoryOnly)
                    .Where(CsvWorkspaceLogic.IsSupportedPath)
                    .OrderBy(p => Path.GetFileName(p), StringComparer.OrdinalIgnoreCase)
                    .ThenBy(p => p, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            catch (Exception ex) { Dialogs.Info("Couldn't read the folder:\n" + ex.Message); return; }

            if (paths.Length == 0)
            {
                Dialogs.Info("No CSV, TSV or TXT files were found in:\n" + folder);
                return;
            }
            OpenPaths(paths, showIndexWhenDone: true);
        }

        /// <summary>Entry point used by native-open interception.</summary>
        public static bool OpenCsvFromPath(string path) => OpenPaths(new[] { path }, showIndexWhenDone: false);

        private static bool OpenPaths(IEnumerable<string> requestedPaths, bool showIndexWhenDone)
        {
            var paths = requestedPaths
                .Where(p => !string.IsNullOrWhiteSpace(p) && CsvWorkspaceLogic.IsSupportedPath(p))
                .Select(CsvWorkspaceLogic.CanonicalPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (paths.Count == 0) return false;

            dynamic app = ExcelApp.App;
            dynamic workbook = null;
            dynamic firstOpenedSheet = null;
            int opened = 0, alreadyOpen = 0;
            bool workspaceCreated = false;
            var warnings = new List<string>();

            try
            {
                workbook = CsvWorkspace.GetOrCreateWorkbook((object)app, out workspaceCreated);
                using (new ExcelStateGuard(manualCalculation: true, suppressEvents: true))
                {
                    for (int i = 0; i < paths.Count; i++)
                    {
                        string path = paths[i];
                        try { app.StatusBar = $"Opening delimited file {i + 1} of {paths.Count}: {Path.GetFileName(path)}"; } catch { }

                        CsvSession existing = CsvSession.FindByPath(path);
                        if (existing != null)
                        {
                            alreadyOpen++;
                            if (firstOpenedSheet == null) firstOpenedSheet = existing.ResolveSheet();
                            continue;
                        }

                        dynamic createdSheet = null;
                        try
                        {
                            byte[] bytes = ReadOpenSourceBytes(path);
                            CsvDocument doc = CsvFile.LoadForPath(bytes, path);
                            var ragged = CsvDetect.FindRaggedRows(doc.DataRows(), out int expected);
                            ValidateExcelSize(doc, path);

                            dynamic sheet = CsvWorkspace.AddSourceSheet(workbook, path);
                            createdSheet = sheet;
                            sheet.Activate(); // CsvStyler freeze panes are applied through ActiveWindow.
                            CsvSheet.WriteRecords(
                                sheet, doc.Records, out int rows, out int dataWidth, out string recordMarkerToken);
                            int headerSheetRow = doc.HeaderRecordIndex >= 0 ? doc.HeaderRecordIndex + 1 : -1;
                            CsvStyler.Apply(app, sheet, headerSheetRow, rows, dataWidth);

                            var session = new CsvSession
                            {
                                SourcePath = path,
                                Format = doc.Format,
                                HeaderSheetRow = headerSheetRow,
                                DataWidth = dataWidth,
                                RowCount = rows,
                                ColumnCount = dataWidth,
                                RecordMarkerToken = recordMarkerToken,
                                RecordMarkerColumn = recordMarkerToken == null ? 0 : dataWidth + 2,
                                OpenedAt = DateTime.Now,
                                Notes = ragged.Count == 0 ? "" :
                                    $"{ragged.Count} ragged row(s); expected {expected} columns; first data row {ragged[0] + 1}",
                            };
                            CsvSession.Set(sheet, session);
                            // Baseline the worksheet representation (ragged rows are padded in the grid),
                            // so an untouched ragged source is not immediately classified as locally edited.
                            session.UpdateBaseline(bytes, CsvSheet.ReadRecords(sheet, dataWidth));

                            if (firstOpenedSheet == null) firstOpenedSheet = sheet;
                            opened++;
                        }
                        catch (Exception ex)
                        {
                            if (createdSheet != null)
                            {
                                bool alerts = true;
                                try
                                {
                                    CsvSession.Remove(createdSheet);
                                    CsvWorkspace.ForgetSheetIdentity(createdSheet);
                                    alerts = (bool)app.DisplayAlerts;
                                    app.DisplayAlerts = false;
                                    createdSheet.Delete();
                                }
                                catch { }
                                finally { try { app.DisplayAlerts = alerts; } catch { } }
                            }
                            warnings.Add(path + ": " + ex.Message);
                        }
                    }
                }

                if (workspaceCreated && opened == 0 && alreadyOpen == 0)
                {
                    string unusedWorkspaceId = CsvWorkspace.TryGetWorkspaceId(workbook);
                    try { workbook.Close(false); } catch { }
                    CsvWorkspace.ForgetWorkspace(unusedWorkspaceId);
                    workbook = null;
                }
                else
                {
                    if (showIndexWhenDone || opened + alreadyOpen > 1) CsvWorkspace.ShowIndex();
                    else
                    {
                        RefreshIndexBestEffort(workbook);
                        try { firstOpenedSheet?.Activate(); } catch { }
                    }
                }
            }
            finally
            {
                try { app.StatusBar = false; } catch { }
            }

            if (warnings.Count > 0)
            {
                string shown = string.Join("\n", warnings.Take(12));
                if (warnings.Count > 12) shown += $"\n...and {warnings.Count - 12} more.";
                Dialogs.Info($"Opened {opened} file(s); {alreadyOpen} already open; {warnings.Count} failed.\n\n" + shown);
            }
            return opened + alreadyOpen > 0;
        }

        private enum SaveResult { Saved, Unchanged, Reloaded, Skipped, Failed }

        [ExcelCommand(Name = "PT_SaveCsv", MenuName = "Power Tools", MenuText = "Save source file")]
        public static void SaveCsv()
        {
            dynamic app = ExcelApp.App;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) return;

            dynamic sheet = app.ActiveSheet;
            var session = CsvSession.Get(sheet);
            if (session == null)
            {
                if (CsvSession.IsWorkspaceWorkbook(wb))
                {
                    Dialogs.Info("Select a source-file sheet to save. The Index sheet is generated metadata.");
                    return;
                }
                ExportSheetAsCsv();
                return;
            }

            SaveSession(session, showSuccess: true, confirmNormalWrite: true);
        }

        [ExcelCommand(Name = "PT_SaveAllCsv", MenuName = "Power Tools", MenuText = "Save all source files")]
        public static void SaveAllCsv()
        {
            dynamic app = ExcelApp.App;
            dynamic workbook = app.ActiveWorkbook;
            SaveAllWorkbook((object)workbook, askConfirmation: true);
        }

        private static void SaveAllWorkbook(object workbookObject, bool askConfirmation)
        {
            dynamic workbook = workbookObject;
            if (workbook == null || !CsvSession.IsWorkspaceWorkbook(workbook))
            {
                Dialogs.Info("No delimited-file workspace is active.");
                return;
            }

            var sheetMap = CsvWorkspace.SheetMap(workbook);
            var edited = new List<CsvSession>();
            foreach (CsvSession session in CsvSession.ForWorkbook((object)workbook))
            {
                sheetMap.TryGetValue(session.SheetCodeName, out dynamic sourceSheet);
                if (session.HasLocalChanges(sourceSheet, force: true)) edited.Add(session);
            }
            if (edited.Count == 0)
            {
                RefreshIndexBestEffort(workbook);
                Dialogs.Info("No source sheets have value changes to save.");
                return;
            }

            if (askConfirmation && !Dialogs.Confirm($"Save {edited.Count} edited source file(s)?\n\n" +
                                                     "Files changed outside Excel will be stopped for review."))
                return;

            int saved = 0, unchanged = 0, conflicts = 0, failed = 0;
            foreach (CsvSession session in edited)
            {
                sheetMap.TryGetValue(session.SheetCodeName, out dynamic sourceSheet);
                switch (SaveSession(session, showSuccess: false, confirmNormalWrite: false,
                    knownSheet: sourceSheet, refreshIndex: false))
                {
                    case SaveResult.Saved: saved++; break;
                    case SaveResult.Unchanged:
                    case SaveResult.Reloaded: unchanged++; break;
                    case SaveResult.Skipped: conflicts++; break;
                    case SaveResult.Failed: failed++; break;
                }
            }

            RefreshIndexBestEffort(workbook);
            Dialogs.Info($"Save All finished.\n\nSaved: {saved}\nUnchanged/reloaded: {unchanged}" +
                         $"\nSkipped conflicts: {conflicts}\nFailed: {failed}");
        }

        [ExcelCommand(Name = "PT_CsvIndex", MenuName = "Power Tools", MenuText = "Show CSV workspace Index")]
        public static void ShowIndex() => CsvWorkspace.ShowIndex();

        private static SaveResult SaveSession(
            CsvSession session, bool showSuccess, bool confirmNormalWrite,
            dynamic knownSheet = null, bool refreshIndex = true)
        {
            try
            {
                dynamic sheet = knownSheet ?? session.ResolveSheet();
                if (sheet == null) return SaveResult.Failed;
                var doc = new CsvDocument
                {
                    Format = session.Format,
                    Records = CsvSheet.ReadRecords(sheet, session.DataWidth),
                };
                ValidateExcelSize(doc, session.SourcePath);

                byte[] outBytes = doc.Save();
                string mismatch = CsvDiff.VerifyRoundTrip(outBytes, session.Format, doc.Records);
                if (mismatch != null)
                {
                    Dialogs.Info("Save BLOCKED by the round-trip self-check:\n" + mismatch +
                                 "\n\nThe file was not written.");
                    return SaveResult.Failed;
                }

                // Normalize internal row metadata before any save decision. This keeps inserted and trailing
                // blank records stable by moving the end sentinel immediately after the logical source rows.
                CsvSheet.SyncRecordMarkers(sheet, session, doc.Records.Count);

                bool localChanges = session.HasLocalChanges(doc.Records);
                bool hasViewOnlyContent = CsvSheet.HasViewOnlyContent(sheet);
                SourceReadState sourceState = ReadSource(session.SourcePath, out byte[] diskBytes, out string readError);
                if (sourceState == SourceReadState.Unavailable)
                {
                    Dialogs.Info("Save BLOCKED because the current source could not be verified:\n" +
                                 session.SourcePath + "\n\n" + readError);
                    return SaveResult.Failed;
                }
                bool sourceExists = sourceState == SourceReadState.Present;

                CsvSourceFingerprint observed = sourceExists ? CsvSourceFingerprint.FromBytes(diskBytes) : null;
                CsvSourceFingerprint proposed = CsvSourceFingerprint.FromBytes(outBytes);
                CsvSaveDisposition disposition = CsvSaveSafety.Assess(
                    session.SourceFingerprint, observed, proposed, localChanges);

                // Another process may already have written exactly what this sheet contains.
                if (disposition == CsvSaveDisposition.AlreadySynchronized)
                {
                    session.UpdateBaseline(diskBytes, doc.Records);
                    if (refreshIndex) RefreshIndexBestEffort(sheet.Parent);
                    if (showSuccess) Dialogs.Info("The disk file already matches this sheet. Baseline refreshed.\n" + session.SourcePath);
                    return SaveResult.Unchanged;
                }

                if (disposition == CsvSaveDisposition.ExternalChange ||
                    disposition == CsvSaveDisposition.Conflict ||
                    disposition == CsvSaveDisposition.SourceMissing ||
                    disposition == CsvSaveDisposition.SourceMissingWithLocalChanges)
                {
                    DateTime openedModified = session.SourceModifiedUtc == DateTime.MinValue
                        ? DateTime.MinValue : session.SourceModifiedUtc.ToLocalTime();
                    DateTime? currentModified = sourceExists ? (DateTime?)File.GetLastWriteTime(session.SourcePath) : null;
                    CsvConflictAction action = CsvConflictForm.Choose(
                        (string)sheet.Name, session.SourcePath, openedModified, currentModified,
                        sourceExists, localChanges, hasViewOnlyContent);

                    if (action == CsvConflictAction.Cancel) return SaveResult.Skipped;
                    if (action == CsvConflictAction.Reload)
                    {
                        if (localChanges || hasViewOnlyContent)
                        {
                            string discarded = localChanges && hasViewOnlyContent
                                ? "local source edits and view-only helper content"
                                : localChanges ? "local source edits" : "view-only helper content";
                            if (!Dialogs.Confirm("Reloading rebuilds this source sheet and will discard " +
                                                 discarded + ". Continue?"))
                                return SaveResult.Skipped;
                        }
                        SourceReadState latestState = ReadSource(session.SourcePath, out byte[] latestBytes, out string latestError);
                        if (latestState != SourceReadState.Present)
                        {
                            Dialogs.Info("Reload BLOCKED because the source is no longer readable:\n" +
                                         session.SourcePath +
                                         (string.IsNullOrEmpty(latestError) ? "" : "\n\n" + latestError));
                            return SaveResult.Skipped;
                        }
                        ReloadSession(session, latestBytes, refreshIndex);
                        if (showSuccess) Dialogs.Info("Reloaded from disk:\n" + session.SourcePath);
                        return SaveResult.Reloaded;
                    }
                    if (action == CsvConflictAction.SaveCopy)
                        return SaveCopy(session, doc, outBytes, showSuccess, refreshIndex);

                    string verb = sourceExists ? "overwrite the externally changed file" : "recreate the missing source file";
                    if (!Dialogs.Confirm("Really " + verb + "?\n\n" + session.SourcePath +
                                         (sourceExists ? "\n\nThe current disk version will be kept as .bak." : "")))
                        return SaveResult.Skipped;

                    if (!DiskStillMatches(session.SourcePath, sourceExists, diskBytes))
                    {
                        Dialogs.Info("Save BLOCKED: the source changed again while the decision dialog was open. Review it again.");
                        return SaveResult.Skipped;
                    }

                    WriteWithBackup(session.SourcePath, outBytes);
                    session.UpdateBaseline(outBytes, doc.Records);
                    if (refreshIndex) RefreshIndexBestEffort(sheet.Parent);
                    if (showSuccess) Dialogs.Info("Saved:\n" + session.SourcePath);
                    return SaveResult.Saved;
                }

                // If the worksheet values never changed, preserve the exact original bytes even if serializing
                // the grid would normalize mixed line endings or unnecessary quotes.
                if (!localChanges)
                {
                    session.IsDirty = false;
                    if (showSuccess) Dialogs.Info("No value changes to save.\n" + session.SourcePath);
                    return SaveResult.Unchanged;
                }

                if (BytesEqual(diskBytes, outBytes))
                {
                    session.UpdateBaseline(diskBytes, doc.Records);
                    if (refreshIndex) RefreshIndexBestEffort(sheet.Parent);
                    if (showSuccess) Dialogs.Info("File already matches the sheet — nothing to write.\n" + session.SourcePath);
                    return SaveResult.Unchanged;
                }

                var diff = CsvDiff.Summarize(CsvFile.LoadForPath(diskBytes, session.SourcePath).Records, doc.Records);
                string summary = diff.HasChanges ? diff.Describe() : "formatting/quoting normalization only";
                if (confirmNormalWrite && !Dialogs.Confirm("Save over:\n" + session.SourcePath +
                                                           "\n\nChanges: " + summary +
                                                           "\n\n(a .bak of the current file is kept)"))
                    return SaveResult.Skipped;

                if (!DiskStillMatches(session.SourcePath, expectedExists: true, expectedBytes: diskBytes))
                {
                    Dialogs.Info("Save BLOCKED: the source changed while the save was being confirmed. Review it again.");
                    return SaveResult.Skipped;
                }

                WriteWithBackup(session.SourcePath, outBytes);
                session.UpdateBaseline(outBytes, doc.Records);
                if (refreshIndex) RefreshIndexBestEffort(sheet.Parent);
                if (showSuccess) Dialogs.Info("Saved:\n" + session.SourcePath);
                return SaveResult.Saved;
            }
            catch (Exception ex)
            {
                Dialogs.Info("Save failed:\n" + ex.Message);
                return SaveResult.Failed;
            }
        }

        private static SaveResult SaveCopy(
            CsvSession session, CsvDocument doc, byte[] outBytes, bool showSuccess, bool refreshIndex)
        {
            string defaultName = SafeFileName(Path.GetFileName(session.SourcePath));
            string newPath = Dialogs.SaveFile(SaveFilter, defaultName);
            if (newPath == null) return SaveResult.Skipped;
            newPath = CsvWorkspaceLogic.CanonicalPath(newPath);
            if (CsvWorkspaceLogic.SameSourcePath(newPath, session.SourcePath))
            {
                Dialogs.Info("Choose a different path for Save a copy, or use Overwrite anyway.");
                return SaveResult.Skipped;
            }
            CsvSession alreadyBound = CsvSession.FindByPath(newPath);
            if (alreadyBound != null && !ReferenceEquals(alreadyBound, session))
            {
                Dialogs.Info("That path is already open on another source sheet:\n" + newPath);
                return SaveResult.Skipped;
            }

            WriteWithBackup(newPath, outBytes);
            session.SourcePath = newPath;
            session.UpdateBaseline(outBytes, doc.Records);
            dynamic sheet = session.ResolveSheet();
            if (refreshIndex && sheet != null) RefreshIndexBestEffort(sheet.Parent);
            if (showSuccess) Dialogs.Info("Saved copy and rebound this sheet to:\n" + newPath);
            return SaveResult.Saved;
        }

        private static void ReloadSession(CsvSession session, byte[] diskBytes, bool refreshIndex)
        {
            CsvDocument doc = CsvFile.LoadForPath(diskBytes, session.SourcePath);
            ValidateExcelSize(doc, session.SourcePath);
            var ragged = CsvDetect.FindRaggedRows(doc.DataRows(), out int expected);
            dynamic app = ExcelApp.App;
            dynamic sheet = session.ResolveSheet();
            using (new ExcelStateGuard(manualCalculation: true, suppressEvents: true))
            {
                sheet.Activate();
                try { if ((bool)sheet.AutoFilterMode) sheet.AutoFilterMode = false; } catch { }
                try
                {
                    int oldMarker = CsvSheet.FindRecordMarkerColumn(sheet, (int)RangeHelpers.LastColumn(sheet));
                    if (oldMarker < 2) oldMarker = session.DataWidth + 2;
                    sheet.Columns(oldMarker - 1).Hidden = false;
                    sheet.Columns(oldMarker).Hidden = false;
                }
                catch { }
                sheet.Cells.Clear();
                CsvSheet.WriteRecords(
                    sheet, doc.Records, out int rows, out int width, out string recordMarkerToken,
                    session.RecordMarkerToken);
                string sheetIdentity = CsvWorkspace.EnsureSheetIdentity(
                    sheet, recordMarkerToken, width + 2);
                if (!string.Equals(sheetIdentity, session.SheetCodeName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Reload changed the source worksheet identity.");
                int header = doc.HeaderRecordIndex >= 0 ? doc.HeaderRecordIndex + 1 : -1;
                CsvStyler.Apply(app, sheet, header, rows, width);
                session.Format = doc.Format;
                session.HeaderSheetRow = header;
                session.DataWidth = width;
                session.RowCount = rows;
                session.ColumnCount = width;
                session.RecordMarkerToken = recordMarkerToken;
                session.RecordMarkerColumn = recordMarkerToken == null ? 0 : width + 2;
                session.Notes = ragged.Count == 0 ? "" :
                    $"{ragged.Count} ragged row(s); expected {expected} columns; first data row {ragged[0] + 1}";
                session.UpdateBaseline(diskBytes, CsvSheet.ReadRecords(sheet, width));
            }
            if (refreshIndex) RefreshIndexBestEffort(sheet.Parent);
        }

        [ExcelCommand(Name = "PT_ExportCsv", MenuName = "Power Tools", MenuText = "Export sheet as CSV...")]
        public static void ExportSheetAsCsv()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;
            if (CsvWorkspace.IsIndexSheet(sheet))
            {
                Dialogs.Info("The Index sheet is generated metadata and is not exported as a source file.");
                return;
            }

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
                byte[] outBytes = doc.Save();

                string mismatch = CsvDiff.VerifyRoundTrip(outBytes, fmt, doc.Records);
                if (mismatch != null)
                {
                    Dialogs.Info("Export BLOCKED by the round-trip self-check:\n" + mismatch +
                                 "\n\nThe file was not written.");
                    return;
                }

                int collisions = CountFieldsContaining(doc.Records, separator);

                WriteWithBackup(path, outBytes);
                Dialogs.Info("Exported to:\n" + path +
                             (collisions > 0
                                 ? $"\n\n{collisions} field(s) contain '{separator}' and were quoted to stay intact."
                                 : ""));
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
            if (CsvWorkspace.IsIndexSheet(sheet))
            {
                Dialogs.Info("The Index sheet is generated metadata and is not restyled as a source file.");
                return;
            }

            CsvSheet.Measure(sheet, out int measuredRows, out int measuredColumns);
            if (measuredRows == 0 || measuredColumns == 0) return;
            long lastRow = measuredRows;
            long lastCol = measuredColumns;

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
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            var session = CsvSession.Get(sheet);
            if (session == null || string.IsNullOrEmpty(session.SourcePath) || !File.Exists(session.SourcePath))
            {
                Dialogs.Info("No source file is associated with the active sheet.");
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
        /// Application.WorkbookOpen handler: when Excel natively opens a CSV/TSV/TXT (double-click in Explorer,
        /// File▸Open, ...), open a safe workspace replacement from the untouched disk bytes and only then
        /// discard Excel's already-coerced parse. Deferred via QueueAsMacro — workbook changes are not safe
        /// inside the Open event itself.
        /// </summary>
        internal static void OnWorkbookOpen(Excel.Workbook wb)
        {
            try
            {
                CsvWorkspace.RekeyOpenedWorkbook(wb);
                if (!InterceptOpenEnabled()) return;

                string path = wb?.FullName;
                if (string.IsNullOrEmpty(path) || !CsvWorkspaceLogic.IsSupportedPath(path)) return;
                if (!File.Exists(path)) return;

                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        // Never discard the native workbook until the safe replacement is fully open. If the
                        // user managed to edit it before this queued macro ran, leave those edits untouched.
                        bool unchanged = true;
                        try { unchanged = (bool)wb.Saved; } catch { }
                        if (!unchanged)
                        {
                            Dialogs.Info("Safe open interception was skipped because the native workbook was edited.\n" +
                                         "Its in-memory changes were left open.");
                            return;
                        }

                        if (OpenCsvFromPath(path))
                            wb.Close(false);       // replacement exists; drop only Excel's coerced native parse
                    }
                    catch (Exception ex)
                    {
                        Dialogs.Info("CSV open interception failed:\n" + ex.Message);
                    }
                });
            }
            catch { /* never let a handler exception break the user's open */ }
        }

        /// <summary>Marks only the source sheet dirty; the generated Index is rebuilt on activation/command.</summary>
        internal static void OnSheetChange(object sh, Excel.Range target)
        {
            try
            {
                CsvSession session = CsvSession.Get(sh);
                if (session != null)
                {
                    session.IsDirty = true;
                    if (UndoManager.IsApplying) return;

                    // A user may append by typing directly into the row that currently holds the hidden end
                    // sentinel. Move it immediately, before a later row deletion or native Excel sort can
                    // delete/misplace it and make a valid blank record disappear.
                    if (CsvSheet.TryGetWorkspaceLayout(
                        sh, out int markerColumn, out int sourceColumns,
                        out int logicalRows, out string ignoredToken, out int sentinelRow) &&
                        logicalRows < 1048576 && sentinelRow != logicalRows + 1)
                    {
                        UndoEntry manualEdit = ManualEditTracker.CurrentChange ??
                            CreateInferredAppendUndo(sh, target, sentinelRow, sourceColumns);
                        if (manualEdit != null)
                        {
                            UndoManager.CaptureCellIfNotCovered(
                                manualEdit, (dynamic)sh, sentinelRow, markerColumn);
                            UndoManager.CaptureCellIfNotCovered(
                                manualEdit, (dynamic)sh, logicalRows + 1, markerColumn);
                        }

                        CsvSheet.SyncRecordMarkers(
                            sh, session, logicalRows, preserveManualEditTrackingShadow: true);
                        UndoManager.PromoteLatestManualUndo(manualEdit);
                        if (manualEdit == null)
                        {
                            UndoManager.ClearHistory();
                            try
                            {
                                ExcelApp.App.StatusBar =
                                    "CSV row boundary preserved; Undo history was cleared because this mixed/large append exceeded the tracked selection.";
                            }
                            catch { }
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Manual tracking may miss a paste/fill larger than the selected destination. When the changed range
        /// is wholly below the old source boundary, its prior source values are known to be blank, so create a
        /// safe replacement undo entry before moving the sentinel (which clears Excel's native undo stack).
        /// </summary>
        private static UndoEntry CreateInferredAppendUndo(
            object sheetObject, Excel.Range target, int firstNewRow, int dataColumns)
        {
            if (sheetObject == null || target == null || firstNewRow < 1 || dataColumns < 1) return null;
            try
            {
                dynamic sheet = sheetObject;
                dynamic changed = target;
                long cellCount = Convert.ToInt64(changed.CountLarge);
                if (cellCount < 1 || cellCount > 50000) return null;

                dynamic areas = changed.Areas;
                int count = (int)areas.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic area = areas(i);
                    int row = (int)area.Row;
                    int column = (int)area.Column;
                    int rows = (int)area.Rows.Count;
                    int columns = (int)area.Columns.Count;
                    if (row < firstNewRow || column < 1 || column + columns - 1 > dataColumns)
                        return null;
                }

                var entry = UndoManager.NewEntry("Append delimited rows", sheet);
                entry.IsManualEdit = true;
                for (int i = 1; i <= count; i++)
                {
                    dynamic area = areas(i);
                    int row = (int)area.Row;
                    int column = (int)area.Column;
                    int rows = (int)area.Rows.Count;
                    int columns = (int)area.Columns.Count;
                    var blanks = new string[rows, columns];
                    for (int r = 0; r < rows; r++)
                        for (int c = 0; c < columns; c++) blanks[r, c] = "";
                    entry.Blocks.Add(new UndoBlock
                    {
                        Row = row,
                        Col = column,
                        Formulas = blanks,
                    });
                    entry.CellCount += rows * columns;
                }
                UndoManager.PushManual(entry);
                return entry;
            }
            catch { return null; }
        }

        internal static void OnSheetActivate(object sh)
        {
            try
            {
                if (CsvWorkspace.SuppressIndexActivationRefresh || !CsvWorkspace.IsIndexSheet(sh)) return;

                dynamic sheet = sh;
                string workspaceId = CsvWorkspace.TryGetWorkspaceId(sheet.Parent);
                if (string.IsNullOrEmpty(workspaceId)) return;

                string key = workspaceId;
                if (!PendingIndexRefreshes.Add(key)) return;

                try
                {
                    // Excel raises SheetActivate synchronously, including from Worksheets.Add. Defer all
                    // workbook mutation until Excel has unwound the event and coalesce repeated activations.
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            dynamic workbook = CsvWorkspace.FindWorkbook(workspaceId);
                            dynamic index = CsvWorkspace.FindIndexSheet(workbook);
                            if (workbook != null && index != null && CsvWorkspace.IsIndexSheet(index))
                                CsvWorkspace.RefreshIndex(workbook);
                        }
                        catch { }
                        finally { PendingIndexRefreshes.Remove(key); }
                    });
                }
                catch { PendingIndexRefreshes.Remove(key); }
            }
            catch { }
        }

        internal static void OnSheetBeforeDelete(object sh)
        {
            try
            {
                CsvWorkspace.ForgetIndexSheet(sh);
                CsvSession.Remove(sh);
                CsvWorkspace.ForgetSheetIdentity(sh);
            }
            catch { }
        }

        internal static void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                string closingWorkbookName = workbook.Name;
                IReadOnlyList<CsvSession> sessions = CsvSession.ForClosingWorkbook(workbook);
                if (sessions.Count == 0) return;

                var edited = new List<CsvSession>();
                foreach (CsvSession session in sessions)
                {
                    dynamic sourceSheet = CsvWorkspace.FindSheetInWorkbook(
                        workbook, session.SheetCodeName);
                    if (session.HasLocalChanges(sourceSheet, force: true)) edited.Add(session);
                }
                if (edited.Count == 0)
                {
                    QueueWorkspaceCleanup(closingWorkbookName, sessions);
                    return;
                }

                DialogResult choice = Dialogs.YesNoCancel(
                    $"{edited.Count} source sheet(s) have unsaved value changes.\n\n" +
                    "Yes: keep the workspace open and run Save All\n" +
                    "No: close without saving the source files\n" +
                    "Cancel: keep the workspace open");
                if (choice == DialogResult.No)
                {
                    QueueWorkspaceCleanup(closingWorkbookName, sessions);
                    return;
                }

                cancel = true;
                if (choice == DialogResult.Yes)
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            workbook.Activate();
                            SaveAllWorkbook((object)workbook, askConfirmation: false);
                        }
                        catch (Exception ex) { Dialogs.Info("Save All could not start:\n" + ex.Message); }
                    });
                }
            }
            catch (Exception ex)
            {
                // An uncertain close check must fail closed. Otherwise a transient COM/defined-name error can
                // silently discard source edits by making a live workspace look like an ordinary workbook.
                if (!CsvSession.HasAnySessions) return;
                cancel = true;
                try
                {
                    Dialogs.Info("Excel Power Tools could not verify whether source sheets have unsaved " +
                                 "changes, so the workbook was kept open.\n\n" + ex.Message);
                }
                catch { }
            }
        }

        private static void QueueWorkspaceCleanup(
            string closingWorkbookName, IEnumerable<CsvSession> closingSessions)
        {
            CsvSession[] sessions = (closingSessions ?? Enumerable.Empty<CsvSession>()).ToArray();
            string[] workspaceIds = sessions
                .Select(session => session.WorkspaceId)
                .Where(id => !string.IsNullOrEmpty(id))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            if (workspaceIds.Length == 0) return;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    // Another BeforeClose handler may have cancelled after ours. Checking the workbook name
                    // still works when its hidden workspace ID is the damaged item that triggered this path.
                    if (CsvWorkspace.IsWorkbookOpenByName(closingWorkbookName)) return;
                    foreach (string workspaceId in workspaceIds)
                    {
                        // A copied workbook can carry historical anchors. Never purge the original session
                        // while a real workbook still claims its runtime workspace identity.
                        bool workspaceIdOwnerIsOpen = CsvWorkspace.FindWorkbook(workspaceId) != null;
                        bool rememberedNameOwnerIsOpen = sessions
                            .Where(session => string.Equals(
                                session.WorkspaceId, workspaceId, StringComparison.Ordinal))
                            .Any(session => CsvWorkspace.FindWorkbookByName(
                                session.RuntimeWorkbookName) != null);
                        if (CsvWorkspaceLogic.KeepRuntimeSessionDuringCleanup(
                            workspaceIdOwnerIsOpen, rememberedNameOwnerIsOpen)) continue;
                        CsvSession.RemoveWorkspace(workspaceId);
                        CsvWorkspace.ForgetWorkspace(workspaceId);
                    }
                }
                catch { }
            });
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

        private static void RefreshIndexBestEffort(dynamic workbook)
        {
            try { CsvWorkspace.RefreshIndex(workbook); }
            catch { /* index metadata must never turn a committed source operation into a reported failure */ }
        }

        // Byte-level file primitives live in Core (AtomicFileStore); thin delegates keep call sites unchanged.
        private static bool BytesEqual(byte[] a, byte[] b) => AtomicFileStore.BytesEqual(a, b);

        /// <summary>
        /// Excel opens a native CSV with write access before WorkbookOpen is raised. File.ReadAllBytes uses
        /// FileShare.Read, which rejects that existing writer even though Excel permits another reader. The
        /// interception path must therefore share both reads and writes while it captures the untouched bytes.
        /// </summary>
        private static byte[] ReadOpenSourceBytes(string path)
        {
            using (var input = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                if (input.Length > int.MaxValue)
                    throw new IOException("The delimited file is too large to open in Excel.");

                using (var output = new MemoryStream((int)input.Length))
                {
                    input.CopyTo(output);
                    return output.ToArray();
                }
            }
        }

        private static SourceReadState ReadSource(string path, out byte[] bytes, out string error) =>
            AtomicFileStore.ReadSource(path, out bytes, out error);

        private static bool DiskStillMatches(string path, bool expectedExists, byte[] expectedBytes) =>
            AtomicFileStore.DiskStillMatches(path, expectedExists, expectedBytes);

        private static int CountFieldsContaining(IReadOnlyList<CsvRecord> records, char c)
        {
            int n = 0;
            foreach (var r in records)
                foreach (string f in r.Fields)
                    if (f != null && f.IndexOf(c) >= 0) n++;
            return n;
        }

        /// <summary>Writes bytes with a one-level ".bak" backup and an atomic replace (no partial file on crash).</summary>
        private static void WriteWithBackup(string path, byte[] bytes) =>
            AtomicFileStore.WriteWithBackup(path, bytes);

        private static void ValidateExcelSize(CsvDocument doc, string path)
        {
            const int maxRows = 1048575; // the following Excel row holds the movable end sentinel
            const int maxColumns = 16382; // two hidden workspace columns preserve blank record boundaries
            const int maxCellChars = 32767;

            if (doc.Records.Count > maxRows)
                throw new InvalidDataException(
                    $"{Path.GetFileName(path)} has {doc.Records.Count} rows; this workspace supports {maxRows} " +
                    "because one Excel row is reserved for its record-boundary sentinel.");

            int widest = 0;
            for (int r = 0; r < doc.Records.Count; r++)
            {
                CsvRecord record = doc.Records[r];
                if (record.FieldCount > widest) widest = record.FieldCount;
                foreach (string field in record.Fields)
                    if (field != null && field.Length > maxCellChars)
                        throw new InvalidDataException(
                            $"{Path.GetFileName(path)} row {r + 1} contains a field longer than Excel's {maxCellChars}-character cell limit.");
            }
            if (widest > maxColumns)
                throw new InvalidDataException(
                    $"{Path.GetFileName(path)} has {widest} columns; this workspace supports {maxColumns} data columns " +
                    "because two Excel columns are reserved internally.");
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
