using System;
using ExcelDna.Integration;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// Feeds the user's own edits into the undo timeline. Mechanism: every selection change refreshes a
    /// "shadow" snapshot of the selected cells; when SheetChange fires, the shadow holds the BEFORE
    /// values of exactly the edited cells → pushed as a "Manual edit" entry. Programmatic writes are
    /// suppressed automatically (ExcelStateGuard). Blind spots by design: edits larger than the shadowed
    /// selection (paste over a single cell, fill-handle drags), formatting, and row/column moves.
    /// </summary>
    internal static class ManualEditTracker
    {
        private const string SettingsKey = @"Software\ExcelPowerTools\Timeline";
        private const long MaxShadowCells = 50000;   // whole-column selections aren't shadowed

        private static int _suppress;
        private static bool? _enabled;
        private static UndoEntry _shadow;

        // ---- enable/disable (ribbon toggle, persisted) --------------------------------------------

        public static bool Enabled
        {
            get
            {
                if (_enabled == null)
                {
                    using (RegistryKey k = Registry.CurrentUser.OpenSubKey(SettingsKey))
                        _enabled = (k?.GetValue("TrackEdits") as string) != "0";   // default: on
                }
                return _enabled.Value;
            }
        }

        public static void SetEnabled(bool on)
        {
            _enabled = on;
            if (!on) _shadow = null;
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(SettingsKey))
                k?.SetValue("TrackEdits", on ? "1" : "0");
        }

        /// <summary>Programmatic-write suppression (taken by ExcelStateGuard for every PT operation).</summary>
        public static IDisposable Suppress() => new SuppressScope();

        private sealed class SuppressScope : IDisposable
        {
            public SuppressScope() { _suppress++; }
            public void Dispose() { _suppress--; _shadow = null; }   // shadow is stale after programmatic writes
        }

        // ---- application events (wired in AddIn.AutoOpen) -----------------------------------------

        public static void OnSelectionChange(object sh, Excel.Range target) => Snapshot(target);

        public static void OnSheetActivate(object sh)
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                Snapshot(app.Selection);
            }
            catch { _shadow = null; }
        }

        public static void OnSheetChange(object sh, Excel.Range target)
        {
            if (!Enabled || _suppress > 0) return;

            UndoEntry shadow = _shadow;
            try
            {
                if (shadow == null || target == null) return;

                dynamic tgt = target;
                dynamic sheet = tgt.Worksheet;
                string wb = (string)sheet.Parent.Name;
                string ws = (string)sheet.Name;
                if (wb != shadow.WorkbookName || ws != shadow.SheetName) return;

                var entry = new UndoEntry
                {
                    WorkbookName = wb,
                    SheetName = ws,
                    IsManualEdit = true,
                };

                string addr = "";
                try { addr = (string)tgt.Address(false, false); } catch { }
                entry.Description = "Edit " + (addr.Length > 48 ? addr.Substring(0, 48) + "…" : addr);

                dynamic areas = tgt.Areas;
                int n = (int)areas.Count;
                for (int a = 1; a <= n; a++)
                {
                    dynamic ar = areas(a);
                    int r0 = (int)ar.Row, c0 = (int)ar.Column;
                    int rows = (int)ar.Rows.Count, cols = (int)ar.Columns.Count;

                    // Only record edits fully covered by the shadow — a partial before-state would
                    // corrupt on undo. Uncovered edits (big paste, fill drag) are simply not tracked.
                    if (!TrySlice(shadow, r0, c0, rows, cols, out UndoBlock block)) return;

                    entry.Blocks.Add(block);
                    entry.CellCount += rows * cols;
                }

                UndoManager.PushManual(entry);
            }
            catch { }
            finally
            {
                // The just-changed cells are the new "before" for the next edit (e.g. repeated F2 edits
                // without moving the selection) — refresh the shadow from the current selection.
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    Snapshot(app.Selection);
                }
                catch { _shadow = null; }
            }
        }

        // ---- internals -----------------------------------------------------------------------------

        private static void Snapshot(dynamic selection)
        {
            _shadow = null;
            if (!Enabled || _suppress > 0 || selection == null) return;

            try
            {
                long cellCount = Convert.ToInt64(selection.CountLarge);
                if (cellCount < 1 || cellCount > MaxShadowCells) return;

                var e = UndoManager.NewEntry("shadow", selection.Worksheet);
                dynamic areas = selection.Areas;
                int n = (int)areas.Count;
                for (int a = 1; a <= n; a++)
                    UndoManager.CaptureBlock(e, areas(a));

                _shadow = e;
            }
            catch { _shadow = null; }   // selection isn't a cell range (shape, chart, ...)
        }

        private static bool TrySlice(UndoEntry shadow, int r0, int c0, int rows, int cols, out UndoBlock block)
        {
            block = null;
            foreach (UndoBlock b in shadow.Blocks)
            {
                int bRows = b.Formulas.GetLength(0), bCols = b.Formulas.GetLength(1);
                if (r0 >= b.Row && c0 >= b.Col
                    && r0 + rows <= b.Row + bRows && c0 + cols <= b.Col + bCols)
                {
                    var f = new string[rows, cols];
                    for (int r = 0; r < rows; r++)
                        for (int c = 0; c < cols; c++)
                            f[r, c] = b.Formulas[r0 - b.Row + r, c0 - b.Col + c];

                    block = new UndoBlock { Row = r0, Col = c0, Formulas = f };
                    return true;
                }
            }
            return false;
        }
    }
}
