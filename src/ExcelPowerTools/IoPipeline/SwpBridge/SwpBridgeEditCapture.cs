using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The SWP bridge's application-event subscribers. The spine is generated content with five
    /// exceptions — the override columns, whose edits persist to the sidecar and recompute the
    /// diff in memory; every other cell reverts. Double-clicking a Verdict/Coordinate cell opens
    /// the SWP block preview; the _Queue Include glyph toggles by double-click; _Dictionary and
    /// _StrobeTemplates are free-form and parse back on every change.
    /// </summary>
    internal static class SwpBridgeEditCapture
    {
        private const int MaxCapturedCells = 2000;

        [ThreadStatic] private static bool _reentry;

        public static void OnSheetChange(object sh, Excel.Range target)
        {
            if (_reentry) return;
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null || target == null || !SwpBridgeWorkspace.IsBridgeSheet(sh)) return;

            _reentry = true;
            try
            {
                dynamic sheet = sh;
                string name = (string)sheet.Name;

                if (name == SwpBridgeWorkspace.SpineSheetName) HandleSpineEdit(sheet, target, session);
                else if (name == SwpBridgeWorkspace.QueueSheetName) HandleQueueSheetEdit(sheet, target, session);
                else if (name == SwpBridgeWorkspace.DictionarySheetName) HandleDictionaryEdit(sheet, session);
                else if (name == SwpBridgeWorkspace.GlossarySheetName) HandleGlossaryEdit(sheet, session);
                else if (name == SwpBridgeWorkspace.OffsetsSheetName) HandleOffsetsEdit(sheet, session);
                else if (name == SwpBridgeWorkspace.StrobeSheetName) HandleStrobeEdit(sheet, session);
                else if (name == SwpBridgeWorkspace.PreviewSheetName) SwpBridgePreview.Rerender(session);
                // Home and _Log revert on the next refresh; stray edits there are harmless.
            }
            catch { }
            finally { _reentry = false; }
        }

        public static void OnSheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null || target == null || !SwpBridgeWorkspace.IsBridgeSheet(sh)) return;

            try
            {
                dynamic sheet = sh;
                string name = (string)sheet.Name;

                if (name == SwpBridgeWorkspace.QueueSheetName)
                {
                    if ((int)target.Column != SwpBridgeQueueSheet.IncludeColumn) return;
                    SwpAction action = SwpBridgeQueueSheet.ActionAtRow(session, (int)target.Row);
                    if (action == null || action.Status != SwpActionStatus.Pending) return;

                    cancel = true;
                    session.Queue.SetInclude(action.ActionId, !action.Include);
                    session.SaveSidecarNow();
                    RunSuppressed(() =>
                        sheet.Cells((int)target.Row, SwpBridgeQueueSheet.IncludeColumn).Value =
                            action.Include ? SwpBridgeQueueSheet.IncludedGlyph : SwpBridgeQueueSheet.ExcludedGlyph);
                    return;
                }

                if (name == SwpBridgeWorkspace.SpineSheetName)
                {
                    int column = (int)target.Column;
                    if (!session.ModelRowByWorksheetRow.TryGetValue((int)target.Row, out int modelRow)) return;

                    if (column == SwpBridgeSpineSheet.ColVerdict || column == SwpBridgeSpineSheet.ColCoordinate)
                    {
                        cancel = true;
                        SwpBridgePreview.Open(session, modelRow);
                    }
                    else if (column == SwpBridgeSpineSheet.ColRow || column == SwpBridgeSpineSheet.ColSwpRow)
                    {
                        cancel = true;
                        OpenSourceRow(session, modelRow, swpSide: column == SwpBridgeSpineSheet.ColSwpRow);
                    }
                    else if (column == SwpBridgeSpineSheet.ColTsNow || column == SwpBridgeSpineSheet.ColTsNew)
                    {
                        cancel = true;
                        string card = "";
                        try { card = ((string)target.Text ?? "").Trim(); } catch { }
                        SwpBridgeAuxSheets.OpenTsRefLookup(sheet.Parent, session, card);
                    }
                    return;
                }

                if (name == SwpBridgeWorkspace.PreviewSheetName)
                {
                    if ((int)target.Row == SwpBridgePreview.ToggleRow)
                    {
                        cancel = true;
                        SwpBridgePreview.ToggleFullBlock(session);
                    }
                    else if ((int)target.Column == 1 && (int)target.Row > SwpBridgePreview.HeaderRow &&
                             session.PreviewContext != null && session.Model != null &&
                             session.PreviewContext.ModelRow < session.Model.Rows.Count)
                    {
                        // The "SWP row" column: open the SWP workbook at that row (the sheet is
                        // the previewed coordinate's — every shown row belongs to it).
                        int rowNumber = 0;
                        try { rowNumber = (int)double.Parse((string)target.Text, System.Globalization.CultureInfo.InvariantCulture); }
                        catch { }
                        SwpSpineRow context = session.Model.Rows[session.PreviewContext.ModelRow];
                        if (rowNumber > 0 && context.Mapped != null && context.Mapped.Ok)
                        {
                            cancel = true;
                            OpenWorkbookAt(session, session.SwpPath,
                                context.Mapped.Coordinate.SheetName, rowNumber);
                        }
                    }
                    return;
                }

                if (name == SwpBridgeWorkspace.GlossarySheetName &&
                    (int)target.Column == SwpBridgeAuxSheets.GlossaryUseColumn)
                {
                    int index = (int)target.Row - SwpBridgeAuxSheets.GlossaryHeaderRow - 1;
                    if (index < 0 || index >= session.GlossaryRelevant.Count) return;

                    cancel = true;
                    Core.IoPipeline.SwpBridge.SwpGlossaryPair pair = session.GlossaryRelevant[index];
                    SwpGlossaryTweak tweak = null;
                    foreach (SwpGlossaryTweak candidate in session.Sidecar.GlossaryTweaks)
                        if (candidate.En.Equals(pair.En, StringComparison.OrdinalIgnoreCase)) { tweak = candidate; break; }
                    if (tweak == null)
                    {
                        tweak = new SwpGlossaryTweak { En = pair.En };
                        session.Sidecar.GlossaryTweaks.Add(tweak);
                    }
                    tweak.Enabled = !pair.Enabled;
                    _reentry = true;
                    try { SwpBridgeBuilder.RecomputeAndRerender(session); }
                    finally { _reentry = false; }
                }
            }
            catch { }
        }

        /// <summary>
        /// Opens the source workbook of a spine row — the IO list at its card-sheet row, or the
        /// SWP at the matched row — reusing an instance already open in this Excel. Opening the
        /// SWP here is for INSPECTION; apply still refuses while it stays open.
        /// </summary>
        private static void OpenSourceRow(SwpBridgeSession session, int modelRow, bool swpSide)
        {
            try
            {
                SwpSpineRow row = session.Model.Rows[modelRow];
                string path, sheetName;
                int rowNumber;
                if (swpSide)
                {
                    if (row.Snapshot == null || row.Mapped == null || !row.Mapped.Ok)
                    {
                        try { ExcelApp.App.StatusBar = "SWP Bridge: this row has no matched SWP row to open."; } catch { }
                        return;
                    }
                    path = session.SwpPath;
                    sheetName = row.Mapped.Coordinate.SheetName;
                    rowNumber = row.Snapshot.RowNumber;
                }
                else
                {
                    path = session.Settings.IoListPath;
                    sheetName = row.Entry.CardSheet;
                    rowNumber = Math.Max(1, row.Entry.SourceRow);
                }

                OpenWorkbookAt(session, path, sheetName, rowNumber);
            }
            catch (Exception ex)
            {
                try { ExcelApp.App.StatusBar = "SWP Bridge: could not open the source row — " + ex.Message; } catch { }
            }
        }

        /// <summary>Opens (or fronts) a workbook and jumps to a sheet row.</summary>
        internal static void OpenWorkbookAt(SwpBridgeSession session, string path,
            string sheetName, int rowNumber)
        {
            try
            {
                dynamic app = ExcelApp.App;
                dynamic workbook = SwpBridgeCardReader.FindOpenWorkbook(app, path);
                if (workbook == null)
                {
                    try { app.StatusBar = "SWP Bridge: opening " + System.IO.Path.GetFileName(path) + "…"; } catch { }
                    workbook = app.Workbooks.Open(path);
                }
                dynamic worksheet = workbook.Worksheets(sheetName);
                app.Goto(worksheet.Cells(rowNumber, 1), true);
                try { app.StatusBar = false; } catch { }
            }
            catch (Exception ex)
            {
                try { ExcelApp.App.StatusBar = "SWP Bridge: could not open the source row — " + ex.Message; } catch { }
            }
        }

        /// <summary>The bridge workbook closing ends the session; everything is already on disk.</summary>
        public static void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                if (SwpBridgeSession.Current != null && SwpBridgeWorkspace.IsLiveBridgeWorkbook(workbook))
                    SwpBridgeSession.Teardown();
            }
            catch { }
        }

        // ---- the spine ----------------------------------------------------------------------

        private static void HandleSpineEdit(dynamic sheet, Excel.Range target, SwpBridgeSession session)
        {
            List<CellRef> cells = EnumerateCells(target);
            bool overridesChanged = false;

            foreach (CellRef cell in cells)
            {
                bool isOverrideColumn =
                    cell.Column >= SwpBridgeSpineSheet.ColSkipLabel &&
                    cell.Column <= SwpBridgeSpineSheet.ColDesc2Override;

                if (!isOverrideColumn || !session.ModelRowByWorksheetRow.TryGetValue(cell.Row, out int modelRow))
                {
                    RestoreSpineCell(sheet, session, cell.Row, cell.Column);
                    continue;
                }

                overridesChanged |= ApplyOverrideEdit(sheet, session, modelRow, cell.Row, cell.Column);
            }

            if (overridesChanged)
            {
                // In-memory recompute: descriptions, coordinates and the diff move with the
                // overrides; the re-render also normalizes every glyph the user typed.
                SwpBridgeBuilder.RecomputeAndRerender(session);
            }
        }

        private static bool ApplyOverrideEdit(dynamic sheet, SwpBridgeSession session,
            int modelRow, int row, int column)
        {
            SwpSpineRow spineRow = session.Model.Rows[modelRow];
            string mnemonic = spineRow.Entry.Mnemonic;
            string typed = CellTextOf(sheet, row, column).Trim();

            SwpRowOverride overrides = session.OverrideFor(mnemonic, create: true);
            switch (column)
            {
                case SwpBridgeSpineSheet.ColSkipLabel:
                    overrides.SkipLabelCheck = IsChecked(typed);
                    return true;
                case SwpBridgeSpineSheet.ColForceIds:
                    overrides.ForceElectricalIds = IsChecked(typed);
                    return true;
                case SwpBridgeSpineSheet.ColDbIndex:
                    if (typed.Length == 0) { overrides.SwpDbIndex = null; return true; }
                    if (int.TryParse(typed, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
                    {
                        overrides.SwpDbIndex = index;
                        return true;
                    }
                    RestoreSpineCell(sheet, session, row, column);
                    try { ExcelApp.App.StatusBar = "SWP Bridge: the DB index must be a whole number (empty = AUTO)."; } catch { }
                    return false;
                case SwpBridgeSpineSheet.ColDesc1Override:
                    overrides.Desc1Override = typed;
                    return true;
                case SwpBridgeSpineSheet.ColDesc2Override:
                    overrides.Desc2Override = typed;
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsChecked(string typed)
        {
            string t = typed.ToLowerInvariant();
            return t == "☑" || t == "x" || t == "1" || t == "yes" || t == "true" || t == "si" || t == "sì";
        }

        private static void RestoreSpineCell(dynamic sheet, SwpBridgeSession session, int row, int column)
        {
            RunSuppressed(() =>
            {
                try
                {
                    object value = "";
                    if (session.SpineData != null &&
                        row >= SwpBridgeSpineSheet.FirstDataRow &&
                        row < SwpBridgeSpineSheet.FirstDataRow + session.SpineData.GetLength(0) &&
                        column >= 1 && column <= session.SpineData.GetLength(1))
                        value = session.SpineData[row - SwpBridgeSpineSheet.FirstDataRow, column - 1] ?? "";
                    else if (row == SwpBridgeSpineSheet.HeaderRow || row == SwpBridgeSpineSheet.MetaRow)
                    {
                        // Header/meta rows are cheap to rebuild wholesale on the next refresh;
                        // just blank the stray edit so nothing false remains.
                    }
                    sheet.Cells(row, column).Value = value;
                }
                catch { }
            });
        }

        // ---- the queue sheet ----------------------------------------------------------------

        private static void HandleQueueSheetEdit(dynamic sheet, Excel.Range target, SwpBridgeSession session)
        {
            bool queueChanged = false;
            foreach (CellRef cell in EnumerateCells(target))
            {
                if (cell.Column != SwpBridgeQueueSheet.IncludeColumn) continue;
                SwpAction action = SwpBridgeQueueSheet.ActionAtRow(session, cell.Row);
                if (action == null || action.Status != SwpActionStatus.Pending) continue;

                string typed = CellTextOf(sheet, cell.Row, cell.Column).Trim().ToLowerInvariant();
                bool include = typed == "x" || typed == "1" || typed == "☑" || typed == "yes" || typed == "true";
                if (action.Include != include)
                {
                    session.Queue.SetInclude(action.ActionId, include);
                    queueChanged = true;
                }
            }

            if (queueChanged) session.SaveSidecarNow();
            SwpBridgeQueueSheet.Refresh(session);   // everything here is generated — re-render restores
        }

        // ---- the editable config sheets -----------------------------------------------------

        private static void HandleDictionaryEdit(dynamic sheet, SwpBridgeSession session)
        {
            if (!SwpBridgeAuxSheets.ParseDictionary(sheet, session)) return;
            SwpBridgeBuilder.RecomputeAndRerender(session);
            try { ExcelApp.App.StatusBar = "SWP Bridge: dictionary saved and applied."; } catch { }
        }

        private static void HandleGlossaryEdit(dynamic sheet, SwpBridgeSession session)
        {
            if (!SwpBridgeAuxSheets.ParseGlossary(sheet, session)) return;
            SwpBridgeBuilder.RecomputeAndRerender(session);
            try { ExcelApp.App.StatusBar = "SWP Bridge: glossary decisions saved and applied."; } catch { }
        }

        private static void HandleOffsetsEdit(dynamic sheet, SwpBridgeSession session)
        {
            if (!SwpBridgeAuxSheets.ParseOffsets(sheet)) return;
            SwpBridgeBuilder.RecomputeAndRerender(session);
            try { ExcelApp.App.StatusBar = "SWP Bridge: index rules saved (all sessions) and applied."; } catch { }
        }

        private static void HandleStrobeEdit(dynamic sheet, SwpBridgeSession session)
        {
            if (!SwpBridgeAuxSheets.ParseStrobeTemplates(sheet, session)) return;
            SwpBridgeBuilder.RecomputeAndRerender(session);
            try { ExcelApp.App.StatusBar = "SWP Bridge: strobe templates saved and applied."; } catch { }
        }

        // ---- plumbing -----------------------------------------------------------------------

        private struct CellRef { public int Row, Column; }

        private static List<CellRef> EnumerateCells(Excel.Range target)
        {
            var cells = new List<CellRef>();
            try
            {
                dynamic tgt = target;
                dynamic areas = tgt.Areas;
                int areaCount = (int)areas.Count;
                for (int a = 1; a <= areaCount && cells.Count <= MaxCapturedCells; a++)
                {
                    dynamic area = areas(a);
                    int r0 = (int)area.Row, c0 = (int)area.Column;
                    int rows = (int)area.Rows.Count, cols = (int)area.Columns.Count;
                    for (int r = 0; r < rows && cells.Count <= MaxCapturedCells; r++)
                        for (int c = 0; c < cols && cells.Count <= MaxCapturedCells; c++)
                            cells.Add(new CellRef { Row = r0 + r, Column = c0 + c });
                }
            }
            catch { }
            return cells;
        }

        private static string CellTextOf(dynamic sheet, int row, int column)
        {
            object raw = sheet.Cells(row, column).Value2;
            return (string)CsvSheet.CellText(raw, sheet, row, column);
        }

        /// <summary>Programmatic bridge-sheet writes with this handler standing down.</summary>
        internal static void RunSuppressed(Action action)
        {
            bool previous = _reentry;
            _reentry = true;
            try { action(); }
            finally { _reentry = previous; }
        }
    }
}
