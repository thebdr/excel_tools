using System;
using System.Collections.Generic;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Experimental.IoBridge;
using ExcelPowerTools.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools.Experimental.IoBridge
{
    /// <summary>
    /// The bridge's application-event subscribers. Typing into a card sheet's data cell queues an
    /// action (the cell keeps the typed value, a mirror flag lights the pending highlight); typing
    /// the old value back unqueues it. Everything outside the editable data region is reverted —
    /// generated content has exactly one source of truth, the model. On the _Queue sheet the Include
    /// glyph toggles by double-click, and every other cell is generated and reverts.
    /// </summary>
    internal static class IoBridgeEditCapture
    {
        private const int MaxCapturedCells = 2000;

        [ThreadStatic] private static bool _reentry;

        public static void OnSheetChange(object sh, Excel.Range target)
        {
            if (_reentry) return;
            IoBridgeSession session = IoBridgeSession.Current;
            if (session == null || target == null || !IoBridgeWorkspace.IsBridgeSheet(sh)) return;

            _reentry = true;
            try
            {
                dynamic sheet = sh;
                string name = (string)sheet.Name;

                if (name == IoBridgeWorkspace.QueueSheetName)
                {
                    HandleQueueSheetEdit(sheet, target, session);
                    return;
                }
                if (name == IoBridgeWorkspace.ColumnsSheetName) return;   // the config editor is free-form

                if (session.SheetMaps.TryGetValue(name, out BridgeSheetMap map))
                    HandleCardSheetEdit(sheet, target, session, map);
            }
            catch { }
            finally { _reentry = false; }
        }

        public static void OnSheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            IoBridgeSession session = IoBridgeSession.Current;
            if (session == null || target == null || !IoBridgeWorkspace.IsBridgeSheet(sh)) return;

            try
            {
                dynamic sheet = sh;
                if ((string)sheet.Name != IoBridgeWorkspace.QueueSheetName) return;
                if ((int)target.Column != IoBridgeQueueSheet.IncludeColumn) return;

                BridgeAction action = IoBridgeQueueSheet.ActionAtRow(session, (int)target.Row);
                if (action == null || action.Status != BridgeActionStatus.Pending) return;

                cancel = true;   // suppress in-cell editing — the double-click IS the gesture
                session.Queue.SetInclude(action.ActionId, !action.Include);
                session.SaveSidecarNow();

                RunSuppressed(() =>
                    sheet.Cells((int)target.Row, IoBridgeQueueSheet.IncludeColumn).Value =
                        action.Include ? IoBridgeQueueSheet.IncludedGlyph : IoBridgeQueueSheet.ExcludedGlyph);
            }
            catch { }
        }

        /// <summary>The bridge workbook closing ends the session; the queue is already on disk.</summary>
        public static void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            try
            {
                if (IoBridgeSession.Current != null && IoBridgeWorkspace.IsLiveBridgeWorkbook(workbook))
                    IoBridgeSession.Teardown();
            }
            catch { }
        }

        // ---- queue sheet -----------------------------------------------------------------------------

        private static void HandleQueueSheetEdit(dynamic sheet, Excel.Range target, IoBridgeSession session)
        {
            bool queueChanged = false;
            foreach (var cell in EnumerateCells(target))
            {
                if (cell.Column == IoBridgeQueueSheet.IncludeColumn)
                {
                    BridgeAction action = IoBridgeQueueSheet.ActionAtRow(session, cell.Row);
                    if (action != null && action.Status == BridgeActionStatus.Pending)
                    {
                        // Typed values count too: x/1/☑ include, blank/0/☐ exclude.
                        string typed = CellTextOf(sheet, cell.Row, cell.Column).Trim().ToLowerInvariant();
                        bool include = typed == "x" || typed == "1" || typed == "☑" ||
                                       typed == "yes" || typed == "true";
                        if (action.Include != include)
                        {
                            session.Queue.SetInclude(action.ActionId, include);
                            queueChanged = true;
                        }
                    }
                }
            }

            if (queueChanged) session.SaveSidecarNow();
            // Everything on this sheet is generated — re-render restores glyphs and reverts stray edits.
            IoBridgeQueueSheet.Refresh(session);
        }

        // ---- card sheets -----------------------------------------------------------------------------

        private static void HandleCardSheetEdit(
            dynamic sheet, Excel.Range target, IoBridgeSession session, BridgeSheetMap map)
        {
            var cells = EnumerateCells(target);
            if (cells.Count > MaxCapturedCells)
            {
                RestoreCells(sheet, session, map, cells);
                try { ExcelApp.App.StatusBar = "IO Bridge: paste too large to queue — the cells were restored."; } catch { }
                return;
            }

            bool queueChanged = false;
            foreach (var cell in cells)
                queueChanged |= CaptureCell(sheet, session, map, cell.Row, cell.Column);

            if (queueChanged)
            {
                session.SaveSidecarNow();
                IoBridgeQueueSheet.Refresh(session);
            }
        }

        /// <summary>Captures one changed cell. True when the queue changed.</summary>
        private static bool CaptureCell(dynamic sheet, IoBridgeSession session, BridgeSheetMap map, int row, int column)
        {
            int layoutIndex = session.LayoutIndexOfSheetColumn(column);

            // Outside the data region (meta/group/header rows, separator, spacers, right of the grid).
            if (layoutIndex < 0 || !map.RowByWorksheetRow.TryGetValue(row, out BridgeRow modelRow))
            {
                RestoreCell(sheet, session, map, row, column);
                return false;
            }

            string oldValue = session.ModelValueAt(map, modelRow, layoutIndex);
            if (oldValue == null)
            {
                // Not editable here: status columns, IO cells of CSV-only rows, CSV cells of rows
                // without a record (integration happens through the row action, not by typing).
                RestoreCell(sheet, session, map, row, column);
                string side = session.Layout.Columns[layoutIndex].Def.Side;
                if (side == "CSV" && modelRow.Csv == null)
                    try { ExcelApp.App.StatusBar = "IO Bridge: no CSV record here — use the row action \"Integrate into CSV\"."; } catch { }
                else if (side == "CSV" && modelRow.Csv != null && modelRow.Csv.IsMulti)
                    try { ExcelApp.App.StatusBar = "IO Bridge: " + modelRow.Csv.RecordIndexes.Count + " CSV rows share this point — edit them in the CSV directly."; } catch { }
                return false;
            }

            string typed = CellTextOf(sheet, row, column);
            LayoutColumn layoutColumn = session.Layout.Columns[layoutIndex];

            BridgeTarget targetDoc;
            string fieldHeader;
            if (layoutColumn.Def.Side == "IO")
            {
                targetDoc = BridgeTarget.IoList;
                fieldHeader = layoutColumn.Def.NormalizedHeader;
            }
            else
            {
                targetDoc = modelRow.Csv.File;
                fieldHeader = session.CsvFieldHeaderFor(
                    modelRow, IoBridgeSession.CanonicalIndexOf(layoutColumn.Def.NormalizedHeader));
                if (fieldHeader.Length == 0)
                {
                    RestoreCell(sheet, session, map, row, column);
                    return false;
                }
            }

            if (typed == oldValue)
            {
                bool removed = session.Queue.Unqueue(targetDoc, modelRow.Signal.IdentityKey(), fieldHeader);
                Paint(sheet, session, map, row, layoutIndex, null);   // back to the source fill (or none)
                return removed;
            }

            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = targetDoc,
                Signal = CloneSignal(modelRow.Signal),
                Linked = LinkedFieldOf(layoutColumn, fieldHeader),
                FieldHeader = fieldHeader,
                OldValue = oldValue,
                NewValue = typed,
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
            Paint(sheet, session, map, row, layoutIndex, BridgeActionStatus.Pending);
            return true;
        }

        internal static SignalRef CloneSignal(SignalRef signal) => new SignalRef
        {
            CardKey = signal.CardKey,
            Address = signal.Address,
            Mnemonic = signal.Mnemonic,
        };

        internal static LinkedField LinkedFieldOf(LayoutColumn column, string fieldHeader)
        {
            string key = column.Def.Side == "CSV" ? column.Def.NormalizedHeader : fieldHeader;
            if (column.Def.Side == "CSV")
            {
                if (key == CsvCanonical.LayoutKey(CsvCanonical.Name)) return LinkedField.Mnemonic;
                if (key == CsvCanonical.LayoutKey(CsvCanonical.Logic)) return LinkedField.NormalCondition;
                if (key == CsvCanonical.LayoutKey(CsvCanonical.Description)) return LinkedField.Description;
                if (key == CsvCanonical.LayoutKey(CsvCanonical.Node) ||
                    key == CsvCanonical.LayoutKey(CsvCanonical.Offset) ||
                    key == CsvCanonical.LayoutKey(CsvCanonical.Bit)) return LinkedField.Address;
                return LinkedField.None;
            }

            if (key == "mnemonic") return LinkedField.Mnemonic;
            if (key == "bit") return LinkedField.Address;
            if (key == "normalcondition") return LinkedField.NormalCondition;
            if (key.StartsWith("descriptionlanguage1", StringComparison.Ordinal) ||
                key.StartsWith("descriptionenglish", StringComparison.Ordinal)) return LinkedField.Description;
            return LinkedField.None;
        }

        // ---- plumbing --------------------------------------------------------------------------------

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

        private static void RestoreCells(dynamic sheet, IoBridgeSession session, BridgeSheetMap map, List<CellRef> cells)
        {
            RunSuppressed(() =>
            {
                foreach (CellRef cell in cells) RestoreCellCore(sheet, session, map, cell.Row, cell.Column);
            });
        }

        private static void RestoreCell(dynamic sheet, IoBridgeSession session, BridgeSheetMap map, int row, int column)
        {
            RunSuppressed(() => RestoreCellCore(sheet, session, map, row, column));
        }

        /// <summary>Puts a cell back to its rendered value (data, status, or blank). Guarded by callers.</summary>
        private static void RestoreCellCore(dynamic sheet, IoBridgeSession session, BridgeSheetMap map, int row, int column)
        {
            try
            {
                int layoutIndex = session.LayoutIndexOfSheetColumn(column);
                string value = "";

                if (map.RowByWorksheetRow.TryGetValue(row, out BridgeRow modelRow) &&
                    layoutIndex >= 0 && layoutIndex < session.Layout.Columns.Count)
                {
                    LayoutColumn layoutColumn = session.Layout.Columns[layoutIndex];
                    switch (layoutColumn.Def.Side)
                    {
                        case "Status":
                            value = layoutColumn.Def.NormalizedHeader == "status.match"
                                ? modelRow.MatchVerdict : modelRow.RowLabel;
                            break;
                        default:
                            value = session.ModelValueAt(map, modelRow, layoutIndex) ?? "";
                            break;
                    }
                }

                sheet.Cells(row, column).Value = value;
            }
            catch { }
        }

        /// <summary>
        /// Immediate highlight feedback on queue add/remove — the cell lights up (or reverts) NOW.
        /// Deliberately NO ExcelStateGuard: interior-color writes fire no events, and toggling
        /// ScreenUpdating inside a SheetChange handler is precisely what deferred the repaint until
        /// the next full refresh.
        /// </summary>
        private static void Paint(
            dynamic sheet, IoBridgeSession session, BridgeSheetMap map, int row, int layoutIndex,
            BridgeActionStatus? status)
        {
            IoBridgeBuilder.PaintActionCell(sheet, session, map, row, layoutIndex, status);
        }

        /// <summary>
        /// Runs programmatic bridge-sheet writes with THIS handler standing down (the events still
        /// fire; they just exit at the reentry gate). No screen-state manipulation, so Excel repaints
        /// the changes immediately — the row-action and expand paths use this.
        /// </summary>
        internal static void RunSuppressed(Action action)
        {
            bool previous = _reentry;
            _reentry = true;
            try { action(); }
            finally { _reentry = previous; }
        }
    }
}
