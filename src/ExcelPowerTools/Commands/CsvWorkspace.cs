using System;
using System.Collections.Generic;
using System.IO;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools.Commands
{
    /// <summary>Workbook/sheet identity and the generated Index sheet for the delimited-file workspace.</summary>
    internal static class CsvWorkspace
    {
        private const string WorkspaceDefinedName = "__PT_CSV_WORKSPACE_ID";
        private const string IndexDefinedName = "__PT_CSV_INDEX_SHEET";
        private const string IndexTitle = "Excel Power Tools CSV Workspace";
        private const int HeaderRow = 3;
        internal const int IndexHeaderRow = HeaderRow;   // the Index table: sessions from the row below, in ForWorkbook order
        internal const int IndexReloadColumn = 1;        // "Status" — the blue-bordered cell whose double-click reloads from disk

        private static readonly HashSet<string> RuntimeWorkspaceIds = new HashSet<string>(StringComparer.Ordinal);
        private static readonly IndexRefreshGate IndexOperations = new IndexRefreshGate();
        private static bool _showingIndex;

        public static bool SuppressIndexActivationRefresh => _showingIndex || IndexOperations.IsEntered;

        public static string EnsureWorkspaceId(dynamic workbook)
        {
            string existing = TryGetWorkspaceId(workbook);
            if (!string.IsNullOrEmpty(existing) && RuntimeWorkspaceIds.Contains(existing)) return existing;

            string id = Guid.NewGuid().ToString("N");
            try
            {
                try { workbook.Names.Item(WorkspaceDefinedName).Delete(); } catch { }
                dynamic name = workbook.Names.Add(WorkspaceDefinedName, "=\"" + id + "\"");
                try { name.Visible = false; } catch { }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not initialize the CSV workspace identity.", ex);
            }
            RuntimeWorkspaceIds.Add(id);
            return id;
        }

        /// <summary>
        /// A workspace workbook saved to disk carries its hidden ID. On a later/opened copy that ID is only
        /// historical and must be replaced so it cannot alias a currently-open runtime workspace.
        /// </summary>
        public static void RekeyOpenedWorkbook(dynamic workbook)
        {
            string persisted = TryGetWorkspaceId(workbook);
            if (string.IsNullOrEmpty(persisted)) return;
            bool wasSaved = true;
            try { wasSaved = (bool)workbook.Saved; } catch { }
            AssignNewWorkspaceId(workbook);
            // The rekey is runtime safety metadata, not a user edit that should cause a close prompt.
            if (wasSaved) try { workbook.Saved = true; } catch { }
        }

        public static string TryGetWorkspaceId(dynamic workbook)
        {
            if (workbook == null) return null;
            try
            {
                dynamic name = workbook.Names.Item(WorkspaceDefinedName);
                string value = ParseDefinedNameValue((string)name.RefersTo);
                return CsvWorkspaceLogic.IsWorkspaceId(value) ? value : null;
            }
            catch { return null; }
        }

        internal static string RequireRuntimeWorkspaceId(dynamic workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            // Unlike TryGetWorkspaceId, this deliberately lets COM failures propagate. A busy Excel
            // object must abort an index refresh; it must never be mistaken for a missing identity and
            // cause a live workbook to be re-keyed.
            dynamic name = workbook.Names.Item(WorkspaceDefinedName);
            string value = ParseDefinedNameValue((string)name.RefersTo);
            if (!CsvWorkspaceLogic.IsWorkspaceId(value) || !RuntimeWorkspaceIds.Contains(value))
                throw new InvalidOperationException("The workbook is not registered as a live CSV workspace.");
            return value;
        }

        /// <summary>Returns the GUID shared by the sheet's unique hidden marker and workbook anchor.</summary>
        public static string SheetCodeName(dynamic sheet)
        {
            try
            {
                if (!CsvSheet.TryReadSingleInternalRecordMarker(
                    sheet, out string markerToken, out int ignoredColumn)) return null;
                string id = CsvSheet.TryGetRecordMarkerId(markerToken);
                if (string.IsNullOrEmpty(id)) return null;
                dynamic resolved = ResolveSheetAnchorStrict(sheet.Parent, id);
                return SameWorksheet(resolved, sheet) ? id : null;
            }
            catch { return null; }
        }

        internal static string EnsureSheetIdentity(dynamic sheet, string markerToken, int markerColumn)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            string id = CsvSheet.TryGetRecordMarkerId(markerToken);
            int actualColumn = -1;
            if (string.IsNullOrEmpty(id) || markerColumn < 3 ||
                !CsvSheet.TryValidateRecordMarker(sheet, markerToken, out actualColumn) ||
                actualColumn != markerColumn)
                throw new InvalidOperationException("The source sheet does not have one valid hidden record marker.");

            dynamic workbook = sheet.Parent;
            string anchorName = CsvWorkspaceLogic.SheetAnchorName(id);
            Excel.Name anchor = null;
            try
            {
                Excel.Name existing = FindWorkbookScopedNameStrict(workbook, anchorName);
                if (existing != null)
                {
                    dynamic existingTarget = ResolveAnchorTargetStrict(workbook, existing);
                    if (SameWorksheet(existingTarget, sheet))
                    {
                        // Reload can change the source width and therefore move the hidden marker column.
                        // Keep the GUID, but move its anchor off any column that has become source data.
                        existing.RefersTo = CsvWorkspaceLogic.SheetColumnAnchorFormula(
                            (string)sheet.Name, markerColumn);
                        if (!AnchorTargetsColumnStrict(workbook, existing, sheet, markerColumn))
                            throw new InvalidOperationException(
                                "The source worksheet anchor did not move with its hidden marker.");
                        return id;
                    }
                    throw new InvalidOperationException("The source marker identity is already anchored to another sheet.");
                }

                var typedWorkbook = (Excel.Workbook)workbook;
                anchor = typedWorkbook.Names.Add(
                    Name: anchorName,
                    RefersTo: CsvWorkspaceLogic.SheetColumnAnchorFormula((string)sheet.Name, markerColumn),
                    Visible: false);
                anchor.Visible = false;

                dynamic resolved = ResolveSheetAnchorStrict(workbook, id);
                int resolvedColumn = -1;
                if (!SameWorksheet(resolved, sheet) ||
                    !AnchorTargetsColumnStrict(workbook, anchor, sheet, markerColumn) ||
                    !CsvSheet.TryValidateRecordMarker(resolved, markerToken, out resolvedColumn) ||
                    resolvedColumn != markerColumn)
                    throw new InvalidOperationException("The worksheet anchor did not resolve to the created sheet.");
                return id;
            }
            catch (Exception ex)
            {
                try { anchor?.Delete(); } catch { }
                throw new InvalidOperationException("Could not initialize a stable worksheet identity.", ex);
            }
        }

        public static dynamic FindWorkbook(string workspaceId)
        {
            if (string.IsNullOrEmpty(workspaceId)) return null;
            try
            {
                dynamic app = ExcelApp.App;
                int count = (int)app.Workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = app.Workbooks(i);
                    if (string.Equals(TryGetWorkspaceId(wb), workspaceId, StringComparison.Ordinal)) return wb;
                }
            }
            catch { }
            return null;
        }

        internal static dynamic FindWorkbookByName(string workbookName)
        {
            if (string.IsNullOrEmpty(workbookName)) return null;
            try
            {
                dynamic app = ExcelApp.App;
                int count = (int)app.Workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic workbook = app.Workbooks(i);
                    if (string.Equals((string)workbook.Name, workbookName, StringComparison.OrdinalIgnoreCase))
                        return workbook;
                }
            }
            catch { }
            return null;
        }

        public static dynamic FindSheet(string workspaceId, string sheetCodeName)
        {
            if (string.IsNullOrEmpty(sheetCodeName)) return null;
            dynamic workbook = FindWorkbook(workspaceId);
            if (workbook == null) return null;
            try { return FindSheetInWorkbookStrict(workbook, sheetCodeName); }
            catch { }
            return null;
        }

        internal static dynamic FindSheetInWorkbook(dynamic workbook, string sheetCodeName)
        {
            try { return FindSheetInWorkbookStrict(workbook, sheetCodeName); }
            catch { return null; }
        }

        /// <summary>
        /// Returns every runtime source identity physically represented in a workbook. Workbook-scoped
        /// anchors and sheet markers are independent fallbacks, so damage to either one cannot hide a live
        /// session from the workbook-close safety check.
        /// </summary>
        internal static HashSet<string> WorkbookSheetIdentityIds(dynamic workbook)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (workbook == null) return result;

            int nameCount = (int)workbook.Names.Count;
            for (int i = 1; i <= nameCount; i++)
            {
                Excel.Name candidate = (Excel.Name)workbook.Names.Item(i);
                Excel.Workbook owner = candidate.Parent as Excel.Workbook;
                if (owner == null || !SameWorkbook(owner, workbook)) continue;
                string id = CsvWorkspaceLogic.TryParseSheetAnchorId(candidate.Name);
                if (!string.IsNullOrEmpty(id)) result.Add(id);
            }

            int sheetCount = (int)workbook.Worksheets.Count;
            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic sheet = workbook.Worksheets(i);
                if (!CsvSheet.TryReadSingleInternalRecordMarker(
                    sheet, out string markerToken, out int ignoredColumn)) continue;
                string id = CsvSheet.TryGetRecordMarkerId(markerToken);
                if (!string.IsNullOrEmpty(id)) result.Add(id);
            }
            return result;
        }

        /// <summary>Workbook-scoped source anchors that currently target this exact worksheet.</summary>
        internal static HashSet<string> SheetIdentityIdsTargeting(dynamic sheet)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (sheet == null) return result;
            dynamic workbook = sheet.Parent;
            int count = (int)workbook.Names.Count;
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    Excel.Name candidate = (Excel.Name)workbook.Names.Item(i);
                    Excel.Workbook owner = candidate.Parent as Excel.Workbook;
                    if (owner == null || !SameWorkbook(owner, workbook)) continue;
                    string id = CsvWorkspaceLogic.TryParseSheetAnchorId(candidate.Name);
                    if (string.IsNullOrEmpty(id)) continue;
                    string refersTo = (string)candidate.RefersTo;
                    if (CsvWorkspaceLogic.BrokenSheetAnchorTargetsSheet(
                            refersTo, (string)sheet.Name) ||
                        SameWorksheet(ResolveAnchorTargetStrict(workbook, candidate), sheet))
                        result.Add(id);
                }
                catch { }
            }
            return result;
        }

        internal static bool IsWorkbookOpenByName(string workbookName)
        {
            return FindWorkbookByName(workbookName) != null;
        }

        public static Dictionary<string, dynamic> SheetMap(dynamic workbook)
        {
            var result = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            if (workbook == null) return result;
            try
            {
                foreach (CsvSession session in CsvSession.ForWorkbook((object)workbook))
                {
                    dynamic sheet = FindSheetInWorkbookStrict(workbook, session.SheetCodeName);
                    int markerColumn = -1;
                    if (sheet == null || !CsvSheet.TryValidateRecordMarker(
                        sheet, session.RecordMarkerToken, out markerColumn)) continue;
                    session.RecordMarkerColumn = markerColumn;
                    result[session.SheetCodeName] = result.ContainsKey(session.SheetCodeName) ? null : sheet;
                }
            }
            catch { }
            return result;
        }

        public static dynamic GetOrCreateWorkbook(dynamic app, out bool created)
        {
            IDisposable operation = IndexOperations.TryEnter();
            if (operation == null)
                throw new InvalidOperationException("The CSV workspace is already being initialized.");

            using (operation)
            using (new ExcelStateGuard(suppressEvents: true))
            {
                dynamic workbook = CsvSession.WorkspaceWorkbook();
                if (workbook != null)
                {
                    created = false;
                    workbook.Activate();
                    EnsureIndexSheetCore(workbook);
                    return workbook;
                }

                created = true;
                workbook = app.Workbooks.Add(Xl.xlWBATWorksheet);
                // The window title instead of "Book1" — the caption is free to set on an
                // unsaved workbook (the workbook NAME is not; it would need a save).
                try { workbook.Windows(1).Caption = "PT_CSV"; } catch { }
                string id = null;
                try
                {
                    id = EnsureWorkspaceId(workbook);
                    dynamic index = workbook.Worksheets(1);
                    index.Name = "Index";
                    CreateIndexAnchorStrict(workbook, index);
                    InitializeIndex(index);
                    return workbook;
                }
                catch
                {
                    if (!string.IsNullOrEmpty(id)) ForgetWorkspace(id);
                    try { workbook.Close(false); } catch { }
                    throw;
                }
            }
        }

        public static dynamic AddSourceSheet(dynamic workbook, string sourcePath)
        {
            using (new ExcelStateGuard(suppressEvents: true))
            {
                var names = new List<string>();
                int sheetCount = (int)workbook.Sheets.Count;
                for (int i = 1; i <= sheetCount; i++)
                    names.Add((string)workbook.Sheets(i).Name);

                string name = CsvWorkspaceLogic.UniqueSheetName(sourcePath, names);
                var typedWorkbook = (Excel.Workbook)workbook;
                var after = (Excel.Worksheet)typedWorkbook.Worksheets[typedWorkbook.Worksheets.Count];
                Excel.Worksheet sheet = null;
                try
                {
                    sheet = (Excel.Worksheet)typedWorkbook.Worksheets.Add(After: after);
                    sheet.Name = name;
                    return sheet;
                }
                catch
                {
                    if (sheet != null)
                    {
                        Excel.Application app = typedWorkbook.Application;
                        bool alerts = true;
                        try
                        {
                            alerts = app.DisplayAlerts;
                            app.DisplayAlerts = false;
                            sheet.Delete();
                        }
                        catch { }
                        finally { try { app.DisplayAlerts = alerts; } catch { } }
                    }
                    throw;
                }
            }
        }

        public static dynamic EnsureIndexSheet(dynamic workbook)
        {
            IDisposable operation = IndexOperations.TryEnter();
            if (operation == null)
                throw new InvalidOperationException("The CSV workspace index is already being updated.");

            using (operation)
            using (new ExcelStateGuard(suppressEvents: true))
                return EnsureIndexSheetCore(workbook);
        }

        private static dynamic EnsureIndexSheetCore(dynamic workbook)
        {
            RequireRuntimeWorkspaceId(workbook);
            Excel.Name roleAnchor = FindWorkbookScopedNameStrict(workbook, IndexDefinedName);
            if (roleAnchor != null)
            {
                string refersTo = (string)roleAnchor.RefersTo;
                if (!ContainsBrokenReference(refersTo))
                {
                    dynamic known = ResolveAnchorTargetStrict(workbook, roleAnchor);
                    if (known == null)
                        throw new InvalidOperationException("The Index worksheet anchor is invalid.");
                    return known;
                }
                roleAnchor.Delete(); // a clearly broken target may be replaced; uncertainty throws above
            }

            var names = new List<string>();
            int allCount = (int)workbook.Sheets.Count;
            for (int i = 1; i <= allCount; i++)
                names.Add((string)workbook.Sheets(i).Name);

            var typedWorkbook = (Excel.Workbook)workbook;
            var first = (Excel.Worksheet)typedWorkbook.Worksheets[1];
            var index = (Excel.Worksheet)typedWorkbook.Worksheets.Add(Before: first);
            Excel.Name createdAnchor = null;
            try
            {
                index.Name = CsvWorkspaceLogic.UniqueSheetName("Index.csv", names);
                createdAnchor = CreateIndexAnchorStrict(workbook, index);
                InitializeIndex(index);
                return index;
            }
            catch
            {
                bool deleted = TryDeleteWorksheet(index);
                if (deleted) try { createdAnchor?.Delete(); } catch { }
                throw;
            }
        }

        private static dynamic FindSheetInWorkbookStrict(dynamic workbook, string sheetCodeName)
        {
            if (string.IsNullOrEmpty(sheetCodeName)) return null;
            return ResolveSheetAnchorStrict(workbook, sheetCodeName);
        }

        private static dynamic ResolveSheetAnchorStrict(dynamic workbook, string sheetId)
        {
            if (!CsvWorkspaceLogic.IsWorkspaceId(sheetId)) return null;
            Excel.Name anchor = FindWorkbookScopedNameStrict(
                workbook, CsvWorkspaceLogic.SheetAnchorName(sheetId));
            return anchor == null ? null : ResolveAnchorTargetStrict(workbook, anchor);
        }

        private static dynamic ResolveAnchorTargetStrict(dynamic workbook, dynamic anchor)
        {
            if (anchor == null) return null;
            string refersTo = (string)anchor.RefersTo;
            if (ContainsBrokenReference(refersTo)) return null;

            try
            {
                Excel.Range range = (Excel.Range)anchor.RefersToRange;
                Excel.Worksheet sheet = range?.Parent as Excel.Worksheet;
                if (sheet == null || !SameWorkbook(sheet.Parent, workbook))
                    throw new InvalidOperationException(
                        "The worksheet identity anchor points outside its owning workbook.");
                return sheet;
            }
            catch (Exception ex)
            {
                // A clearly broken reference is missing. Any other COM failure is uncertainty and must abort
                // the operation rather than creating a second anchor for a live worksheet.
                string current = (string)anchor.RefersTo;
                if (ContainsBrokenReference(current)) return null;
                if (ex is InvalidOperationException) throw;
                throw new InvalidOperationException("Could not resolve the worksheet identity anchor.", ex);
            }
        }

        private static bool AnchorTargetsColumnStrict(
            dynamic workbook, dynamic anchor, dynamic expectedSheet, int expectedColumn)
        {
            if (!SameWorksheet(ResolveAnchorTargetStrict(workbook, anchor), expectedSheet)) return false;
            Excel.Range range = (Excel.Range)anchor.RefersToRange;
            return range != null && range.Column == expectedColumn && range.Columns.Count == 1;
        }

        internal static bool SameWorksheet(dynamic left, dynamic right)
        {
            if (left == null || right == null) return false;
            return SameWorkbook(left.Parent, right.Parent) &&
                string.Equals((string)left.Name, (string)right.Name, StringComparison.OrdinalIgnoreCase);
        }

        internal static bool SameWorkbook(dynamic left, dynamic right)
        {
            if (left == null || right == null) return false;
            return string.Equals((string)left.Name, (string)right.Name, StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsBrokenReference(string refersTo) =>
            !string.IsNullOrEmpty(refersTo) &&
            refersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0;

        private static Excel.Name FindWorkbookScopedNameStrict(dynamic workbook, string shortName)
        {
            if (workbook == null || string.IsNullOrEmpty(shortName)) return null;
            Excel.Name match = null;
            int count = (int)workbook.Names.Count;
            for (int i = 1; i <= count; i++)
            {
                Excel.Name candidate = (Excel.Name)workbook.Names.Item(i);
                Excel.Workbook owner = candidate.Parent as Excel.Workbook;
                if (owner == null || !SameWorkbook(owner, workbook) ||
                    !string.Equals(candidate.Name, shortName, StringComparison.OrdinalIgnoreCase)) continue;
                if (match != null)
                    throw new InvalidOperationException("More than one workbook-scoped name has the same identity.");
                match = candidate;
            }
            return match;
        }

        private static Excel.Name CreateIndexAnchorStrict(dynamic workbook, dynamic sheet)
        {
            if (FindWorkbookScopedNameStrict(workbook, IndexDefinedName) != null)
                throw new InvalidOperationException("The workbook already contains an Index worksheet anchor.");

            Excel.Name anchor = null;
            try
            {
                var typedWorkbook = (Excel.Workbook)workbook;
                anchor = typedWorkbook.Names.Add(
                    Name: IndexDefinedName,
                    RefersTo: CsvWorkspaceLogic.SheetAnchorFormula((string)sheet.Name),
                    Visible: false);
                anchor.Visible = false;
                if (!SameWorksheet(ResolveAnchorTargetStrict(workbook, anchor), sheet))
                    throw new InvalidOperationException("The Index worksheet anchor did not resolve correctly.");
                return anchor;
            }
            catch
            {
                try { anchor?.Delete(); } catch { }
                throw;
            }
        }

        internal static dynamic FindIndexSheet(dynamic workbook)
        {
            if (workbook == null) return null;
            try
            {
                Excel.Name anchor = FindWorkbookScopedNameStrict(workbook, IndexDefinedName);
                return anchor == null ? null : ResolveAnchorTargetStrict(workbook, anchor);
            }
            catch { return null; }
        }

        internal static void ForgetSheetIdentity(dynamic sheet)
        {
            if (sheet == null) return;
            foreach (string sheetId in SheetIdentityIdsTargeting(sheet))
            {
                try { DeleteSheetIdentity(sheet.Parent, sheetId); }
                catch { }
            }
        }

        private static void DeleteSheetIdentity(dynamic workbook, string sheetId)
        {
            if (workbook == null || !CsvWorkspaceLogic.IsWorkspaceId(sheetId)) return;
            Excel.Name anchor = FindWorkbookScopedNameStrict(
                workbook, CsvWorkspaceLogic.SheetAnchorName(sheetId));
            anchor?.Delete();
        }

        private static bool TryDeleteWorksheet(Excel.Worksheet sheet)
        {
            if (sheet == null) return true;
            Excel.Application app = sheet.Application;
            bool alerts = true;
            try
            {
                alerts = app.DisplayAlerts;
                app.DisplayAlerts = false;
                sheet.Delete();
                return true;
            }
            catch { return false; }
            finally { try { app.DisplayAlerts = alerts; } catch { } }
        }

        public static bool IsIndexSheet(dynamic sheet)
        {
            if (sheet == null) return false;
            try
            {
                string id = TryGetWorkspaceId(sheet.Parent);
                if (string.IsNullOrEmpty(id) || !RuntimeWorkspaceIds.Contains(id)) return false;
                return SameWorksheet(FindIndexSheet(sheet.Parent), sheet);
            }
            catch { }
            return false;
        }

        public static void RefreshIndex(dynamic workbook)
        {
            if (workbook == null) return;
            IDisposable operation = IndexOperations.TryEnter();
            if (operation == null) return;

            using (operation)
            using (new ExcelStateGuard(suppressEvents: true))
                RefreshIndexCore(workbook);
        }

        private static void RefreshIndexCore(dynamic workbook)
        {
            dynamic index = EnsureIndexSheetCore(workbook);
            IReadOnlyList<CsvSession> sessions = CsvSession.ForWorkbook(workbook);
            Dictionary<string, dynamic> sheets = SheetMap(workbook);

            string[] headers =
            {
                "Status", "Sheet", "File", "Full path", "Created", "Modified", "Opened",
                "Size (bytes)", "Rows", "Columns", "Separator", "Encoding", "BOM", "Line endings", "Notes"
            };
            var data = new object[Math.Max(1, sessions.Count), headers.Length];

            for (int i = 0; i < sessions.Count; i++)
            {
                CsvSession s = sessions[i];
                sheets.TryGetValue(s.SheetCodeName, out dynamic sourceSheet);
                string sheetName = sourceSheet == null ? "(sheet deleted)" : (string)sourceSheet.Name;
                RefreshDimensions(s, sourceSheet);

                DateTime? created = null, modified = null;
                long? size = null;
                try
                {
                    if (File.Exists(s.SourcePath))
                    {
                        var info = new FileInfo(s.SourcePath);
                        created = info.CreationTime;
                        modified = info.LastWriteTime;
                        size = info.Length;
                    }
                }
                catch { }

                data[i, 0] = s.StatusText(sourceSheet);
                data[i, 1] = sheetName;
                data[i, 2] = SafeFileName(s.SourcePath);
                data[i, 3] = s.SourcePath;
                data[i, 4] = created;
                data[i, 5] = modified;
                data[i, 6] = s.OpenedAt;
                data[i, 7] = size;
                data[i, 8] = s.RowCount;
                data[i, 9] = s.ColumnCount;
                data[i, 10] = DelimiterLabel(s.Format?.Delimiter ?? ';');
                data[i, 11] = s.Format?.Encoding?.WebName ?? "";
                data[i, 12] = s.Format?.HasBom == true ? "Yes" : "No";
                data[i, 13] = LineEndingLabel(s.Format?.LineEnding ?? CsvLineEnding.Crlf);
                data[i, 14] = s.Notes ?? "";
            }

            try { index.Hyperlinks.Delete(); } catch { }
            index.Cells.Clear();
            index.Cells.NumberFormat = "General";

            // Treat all user/file-derived strings literally (filenames such as "=report.csv" are valid).
            foreach (int c in new[] { 1, 2, 3, 4, 11, 12, 13, 14, 15 })
                index.Columns(c).NumberFormat = "@";

            index.Cells(1, 1).Value2 = IndexTitle;
            index.Cells(1, 2).Value2 = "Updated";
            index.Cells(1, 3).Value = DateTime.Now;

            var headerGrid = new object[1, headers.Length];
            for (int c = 0; c < headers.Length; c++) headerGrid[0, c] = headers[c];
            index.Range(index.Cells(HeaderRow, 1), index.Cells(HeaderRow, headers.Length)).Value = headerGrid;

            if (sessions.Count > 0)
                index.Range(index.Cells(HeaderRow + 1, 1), index.Cells(HeaderRow + sessions.Count, headers.Length)).Value = data;
            else
                index.Cells(HeaderRow + 1, 1).Value2 = "No source files open";

            StyleIndex(index, sessions.Count, headers.Length);

            for (int i = 0; i < sessions.Count; i++)
            {
                sheets.TryGetValue(sessions[i].SheetCodeName, out dynamic sourceSheet);
                if (sourceSheet == null) continue;
                string sheetName = (string)sourceSheet.Name;
                try
                {
                    index.Hyperlinks.Add(index.Cells(HeaderRow + 1 + i, 2), "",
                        CsvWorkspaceLogic.EscapeSheetNameForSubAddress(sheetName), Type.Missing, sheetName);
                }
                catch { }
            }
        }

        public static void ShowIndex()
        {
            dynamic workbook = null;
            try
            {
                dynamic active = ExcelApp.App.ActiveWorkbook;
                if (active != null && CsvSession.IsWorkspaceWorkbook(active)) workbook = active;
            }
            catch { }
            if (workbook == null) workbook = CsvSession.WorkspaceWorkbook();
            if (workbook == null)
            {
                Dialogs.Info("No delimited-file workspace is open.");
                return;
            }

            EnsureIndexSheet(workbook);
            string workspaceId = RequireRuntimeWorkspaceId(workbook);
            try
            {
                // A newly created workbook is not always ready for the destructive Clear/format/write
                // sequence during the file-open callback. Run the refresh and activation in Excel's
                // queued macro context, retaining only the stable id because the workbook may close
                // before the queued work executes.
                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() => RefreshAndActivateIndex(workspaceId));
            }
            catch { RefreshAndActivateIndex(workspaceId); }
        }

        private static void RefreshAndActivateIndex(string workspaceId)
        {
            dynamic workbook = FindWorkbook(workspaceId);
            if (workbook == null) return;

            // Index metadata is best effort and must never prevent navigation to the sheet.
            try { RefreshIndex(workbook); } catch { }
            ActivateIndex(workspaceId);
        }

        private static void ActivateIndex(string workspaceId)
        {
            dynamic workbook = FindWorkbook(workspaceId);
            dynamic index = FindIndexSheet(workbook);
            if (workbook == null || index == null) return;

            _showingIndex = true;
            try
            {
                workbook.Activate();
                index.Activate();
                dynamic window = ExcelApp.App.ActiveWindow;
                window.FreezePanes = false;
                window.SplitRow = HeaderRow;
                window.SplitColumn = 0;
                window.FreezePanes = true;
            }
            catch { }
            finally { _showingIndex = false; }
        }

        public static void ForgetWorkspace(dynamic workbook)
        {
            string id = TryGetWorkspaceId(workbook);
            if (string.IsNullOrEmpty(id)) return;
            RuntimeWorkspaceIds.Remove(id);
        }

        public static void ForgetIndexSheet(dynamic sheet)
        {
            if (sheet == null) return;
            try
            {
                dynamic workbook = sheet.Parent;
                Excel.Name anchor = FindWorkbookScopedNameStrict(workbook, IndexDefinedName);
                if (anchor != null && SameWorksheet(ResolveAnchorTargetStrict(workbook, anchor), sheet))
                    anchor.Delete();
            }
            catch { }
        }

        public static void ForgetWorkspace(string workspaceId)
        {
            if (string.IsNullOrEmpty(workspaceId)) return;
            RuntimeWorkspaceIds.Remove(workspaceId);
        }

        private static void InitializeIndex(dynamic index)
        {
            using (new ExcelStateGuard(suppressEvents: true))
            {
                index.Cells.Clear();
                index.Cells(1, 1).Value2 = IndexTitle;
            }
        }

        private static void RefreshDimensions(CsvSession session, dynamic sheet)
        {
            if (sheet == null) return;
            try
            {
                CsvSheet.Measure(sheet, out int rows, out int columns);
                session.RowCount = rows;
                session.ColumnCount = columns;
            }
            catch { }
        }

        private static void StyleIndex(dynamic sheet, int rows, int columns)
        {
            int lastRow = HeaderRow + Math.Max(rows, 1);
            dynamic title = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, columns));
            title.Font.Name = "Segoe UI";
            title.Font.Bold = true;
            title.Font.Size = 14;

            dynamic header = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, columns));
            header.Font.Name = "Segoe UI";
            header.Font.Bold = true;
            header.Interior.Color = Rgb(180, 210, 240);

            dynamic table = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(lastRow, columns));
            table.Font.Name = "Segoe UI";
            try
            {
                if ((bool)sheet.AutoFilterMode) sheet.AutoFilterMode = false;
                table.AutoFilter();
            }
            catch { }

            try
            {
                sheet.Columns(1).ColumnWidth = 30;
                sheet.Columns(2).ColumnWidth = 24;
                sheet.Columns(3).ColumnWidth = 30;
                sheet.Columns(4).ColumnWidth = 70;
                sheet.Columns(5).ColumnWidth = 20;
                sheet.Columns(6).ColumnWidth = 20;
                sheet.Columns(7).ColumnWidth = 20;
                sheet.Columns(8).ColumnWidth = 14;
                sheet.Columns(9).ColumnWidth = 10;
                sheet.Columns(10).ColumnWidth = 10;
                sheet.Columns(11).ColumnWidth = 16;
                sheet.Columns(12).ColumnWidth = 18;
                sheet.Columns(13).ColumnWidth = 8;
                sheet.Columns(14).ColumnWidth = 13;
                sheet.Columns(15).ColumnWidth = 45;
                sheet.Range(sheet.Cells(HeaderRow + 1, 5), sheet.Cells(lastRow, 7)).NumberFormat = "yyyy-mm-dd hh:mm:ss";
                sheet.Columns(4).NumberFormat = "@";
                // The Status cells react to a double-click (reload from disk): the repo-wide blue
                // border says so — the gesture is restricted to them (CsvCommands.OnSheetBeforeDoubleClick).
                if (rows > 0)
                {
                    dynamic reloadCells = sheet.Range(
                        sheet.Cells(HeaderRow + 1, IndexReloadColumn), sheet.Cells(HeaderRow + rows, IndexReloadColumn));
                    reloadCells.Borders.LineStyle = 1;   // xlContinuous
                    reloadCells.Borders.Weight = 2;      // xlThin
                    reloadCells.Borders.Color = Rgb(0, 112, 192);
                }
            }
            catch { }

            for (int i = 0; i < rows; i++)
            {
                try
                {
                    dynamic cell = sheet.Cells(HeaderRow + 1 + i, 1);
                    string status = cell.Value2 as string ?? "";
                    if (status.StartsWith("Conflict", StringComparison.Ordinal)) cell.Interior.Color = Rgb(255, 199, 206);
                    else if (status == "Edited") cell.Interior.Color = Rgb(255, 235, 156);
                    else if (status == "Changed externally") cell.Interior.Color = Rgb(255, 217, 102);
                    else if (status.StartsWith("Missing", StringComparison.Ordinal)) cell.Interior.Color = Rgb(217, 217, 217);
                    else cell.Interior.Color = Rgb(198, 239, 206);
                }
                catch { }
            }
        }

        private static string ParseDefinedNameValue(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo)) return null;
            string value = refersTo.Trim();
            if (value.StartsWith("=", StringComparison.Ordinal)) value = value.Substring(1);
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"')
                value = value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
            return value;
        }

        private static string AssignNewWorkspaceId(dynamic workbook)
        {
            string id = Guid.NewGuid().ToString("N");
            try
            {
                try { workbook.Names.Item(WorkspaceDefinedName).Delete(); } catch { }
                dynamic name = workbook.Names.Add(WorkspaceDefinedName, "=\"" + id + "\"");
                try { name.Visible = false; } catch { }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not initialize the CSV workspace identity.", ex);
            }
            RuntimeWorkspaceIds.Add(id);
            return id;
        }

        private static string SafeFileName(string path)
        {
            try { return Path.GetFileName(path ?? ""); } catch { return path ?? ""; }
        }

        private static string DelimiterLabel(char delimiter)
        {
            switch (delimiter)
            {
                case '\t': return "Tab";
                case ',': return "Comma (,)";
                case ';': return "Semicolon (;)";
                case '|': return "Pipe (|)";
                default: return delimiter.ToString();
            }
        }

        private static string LineEndingLabel(CsvLineEnding ending) =>
            ending == CsvLineEnding.Lf ? "LF" : ending == CsvLineEnding.Cr ? "CR" : "CRLF";

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
    }
}
