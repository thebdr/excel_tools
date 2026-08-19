using System;
using System.Collections.Generic;
using ExcelDna.Integration;

namespace ExcelPowerTools.Interop
{
    /// <summary>A rectangular snapshot of cell formulas/values (top-left 1-based).</summary>
    internal sealed class UndoBlock
    {
        public int Row;
        public int Col;
        public string[,] Formulas;   // [rows, cols], 0-based; "" = empty cell
    }

    /// <summary>One timeline step: which cells an operation touched and what they held before it.</summary>
    internal sealed class UndoEntry
    {
        public string Description;
        public string WorkbookName;
        public string SheetName;
        public bool IsManualEdit;
        public DateTime Timestamp = DateTime.Now;
        public List<UndoBlock> Blocks = new List<UndoBlock>();
        public int CellCount;
    }

    /// <summary>
    /// The unified undo timeline — the thing VBA never gave us. Power-Tools commands snapshot the cells
    /// they change; the ManualEditTracker feeds the user's own edits into the same stack. Entries store
    /// .Formula strings, so real formulas survive. Application.OnUndo is registered after each PT
    /// operation, so NATIVE Ctrl+Z undoes it immediately afterwards (then toggles to Redo). The full
    /// timeline (50 deep) is reachable via PT_UndoLast/PT_RedoLast (Ctrl+Alt+Z/Y) and the Timeline window.
    /// Cell contents are restored; formatting is not tracked.
    /// </summary>
    internal static class UndoManager
    {
        public const int MaxDepth = 50;

        private static readonly List<UndoEntry> _undo = new List<UndoEntry>();
        private static readonly List<UndoEntry> _redo = new List<UndoEntry>();
        private static int _applyDepth;

        /// <summary>True while saved cell blocks are being restored; event handlers must not rewrite them.</summary>
        internal static bool IsApplying => _applyDepth > 0;

        /// <summary>Raised whenever the stacks change (push, undo, redo) — the Timeline window listens.</summary>
        public static event Action Changed;

        /// <summary>Oldest → newest. The entry at the END is the next to be undone.</summary>
        public static IReadOnlyList<UndoEntry> UndoEntries => _undo;

        /// <summary>Oldest → newest. The entry at the END is the next to be redone.</summary>
        public static IReadOnlyList<UndoEntry> RedoEntries => _redo;

        // ---- capture ------------------------------------------------------------------------------

        public static UndoEntry NewEntry(string description, dynamic sheet)
        {
            var e = new UndoEntry { Description = description };
            try
            {
                e.SheetName = (string)sheet.Name;
                e.WorkbookName = (string)sheet.Parent.Name;
            }
            catch { }
            return e;
        }

        /// <summary>Snapshots a single cell (call BEFORE writing it).</summary>
        public static void CaptureCell(UndoEntry entry, dynamic sheet, int row, int col)
        {
            var f = new string[1, 1];
            f[0, 0] = FormulaString(sheet.Cells(row, col).Formula);
            entry.Blocks.Add(new UndoBlock { Row = row, Col = col, Formulas = f });
            entry.CellCount++;
        }

        /// <summary>Snapshots a cell only when an earlier block does not already contain its pre-action state.</summary>
        public static void CaptureCellIfNotCovered(UndoEntry entry, dynamic sheet, int row, int col)
        {
            if (entry == null) return;
            foreach (UndoBlock block in entry.Blocks)
            {
                int rows = block.Formulas.GetLength(0), cols = block.Formulas.GetLength(1);
                if (row >= block.Row && row < block.Row + rows &&
                    col >= block.Col && col < block.Col + cols) return;
            }
            CaptureCell(entry, sheet, row, col);
        }

        /// <summary>Snapshots a rectangular range in one COM call (call BEFORE writing it).</summary>
        public static void CaptureBlock(UndoEntry entry, dynamic range)
        {
            int r0 = (int)range.Row, c0 = (int)range.Column;
            int rows = (int)range.Rows.Count, cols = (int)range.Columns.Count;

            var f = new string[rows, cols];
            object raw = range.Formula;
            if (raw is object[,] g)
            {
                int rb = g.GetLowerBound(0), cb = g.GetLowerBound(1);
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        f[r, c] = FormulaString(g[rb + r, cb + c]);
            }
            else
            {
                f[0, 0] = FormulaString(raw);
            }

            entry.Blocks.Add(new UndoBlock { Row = r0, Col = c0, Formulas = f });
            entry.CellCount += rows * cols;
        }

        /// <summary>Publishes a PT operation: stacks it and makes native Ctrl+Z undo it (one-shot).</summary>
        public static void Push(UndoEntry entry)
        {
            if (entry == null || entry.CellCount == 0) return;

            AddToUndo(entry);
            _redo.Clear();
            RegisterOnUndo("Undo " + entry.Description + " (Power Tools)", "PT_UndoLast");
            RaiseChanged();
        }

        /// <summary>Publishes a tracked MANUAL edit: stacked, but Excel's own Ctrl+Z keeps handling it natively.</summary>
        public static void PushManual(UndoEntry entry)
        {
            if (entry == null || entry.CellCount == 0) return;

            AddToUndo(entry);
            _redo.Clear();
            RaiseChanged();
        }

        /// <summary>
        /// A companion programmatic write extended the latest manual edit (for example moving CSV metadata),
        /// which clears Excel's native undo stack. Register the already-stacked entry as the replacement.
        /// </summary>
        public static void PromoteLatestManualUndo(UndoEntry entry)
        {
            if (entry == null || !entry.IsManualEdit || _undo.Count == 0 ||
                !ReferenceEquals(_undo[_undo.Count - 1], entry)) return;
            _redo.Clear();
            RegisterOnUndo("Undo " + entry.Description + " (Power Tools)", "PT_UndoLast");
            RaiseChanged();
        }

        /// <summary>Creates a safety barrier after an untracked edit so older steps cannot overwrite it.</summary>
        public static void ClearHistory()
        {
            _undo.Clear();
            _redo.Clear();
            RaiseChanged();
        }

        // ---- undo / redo --------------------------------------------------------------------------

        public static void UndoLast() => UndoSteps(1);
        public static void RedoLast() => RedoSteps(1);

        /// <summary>Walks the timeline back <paramref name="count"/> steps.</summary>
        public static void UndoSteps(int count)
        {
            if (_undo.Count == 0) { Dialogs.Info("Nothing to undo."); return; }

            int done = 0;
            string lastDesc = null;
            while (done < count && _undo.Count > 0)
            {
                UndoEntry e = _undo[_undo.Count - 1];
                _undo.RemoveAt(_undo.Count - 1);

                UndoEntry inverse = SwapApply(e);
                if (inverse == null) continue;   // sheet gone — step dropped

                _redo.Add(inverse);
                if (_redo.Count > MaxDepth) _redo.RemoveAt(0);
                lastDesc = e.Description;
                done++;
            }

            if (lastDesc != null)
                RegisterOnUndo("Redo " + lastDesc + " (Power Tools)", "PT_RedoLast");
            RaiseChanged();
        }

        /// <summary>Replays <paramref name="count"/> undone steps.</summary>
        public static void RedoSteps(int count)
        {
            if (_redo.Count == 0) { Dialogs.Info("Nothing to redo."); return; }

            int done = 0;
            string lastDesc = null;
            while (done < count && _redo.Count > 0)
            {
                UndoEntry e = _redo[_redo.Count - 1];
                _redo.RemoveAt(_redo.Count - 1);

                UndoEntry inverse = SwapApply(e);
                if (inverse == null) continue;

                AddToUndo(inverse);
                lastDesc = e.Description;
                done++;
            }

            if (lastDesc != null)
                RegisterOnUndo("Undo " + lastDesc + " (Power Tools)", "PT_UndoLast");
            RaiseChanged();
        }

        /// <summary>Captures the current state of the entry's cells (for the inverse op), then restores the entry.</summary>
        private static UndoEntry SwapApply(UndoEntry e)
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic sheet = FindSheet(app, e);
            if (sheet == null)
            {
                Dialogs.Info($"Can't undo \"{e.Description}\" — workbook/sheet " +
                             $"'{e.WorkbookName}' / '{e.SheetName}' is no longer open. Step skipped.");
                return null;
            }

            var inverse = new UndoEntry
            {
                Description = e.Description,
                WorkbookName = e.WorkbookName,
                SheetName = e.SheetName,
                IsManualEdit = e.IsManualEdit,
            };

            _applyDepth++;
            try
            {
                using (new ExcelStateGuard(manualCalculation: true))
                {
                    foreach (UndoBlock b in e.Blocks)
                    {
                        int rows = b.Formulas.GetLength(0), cols = b.Formulas.GetLength(1);
                        dynamic range = sheet.Range(sheet.Cells(b.Row, b.Col), sheet.Cells(b.Row + rows - 1, b.Col + cols - 1));

                        CaptureBlock(inverse, range);   // current content, so the step can be re-done

                        if (rows == 1 && cols == 1)
                        {
                            SetCell(sheet.Cells(b.Row, b.Col), b.Formulas[0, 0]);
                        }
                        else
                        {
                            var data = new object[rows, cols];
                            for (int r = 0; r < rows; r++)
                                for (int c = 0; c < cols; c++)
                                    data[r, c] = b.Formulas[r, c];
                            range.Formula = data;
                        }
                    }
                }
            }
            finally { _applyDepth--; }

            try { sheet.Activate(); } catch { }
            return inverse;
        }

        // ---- plumbing -----------------------------------------------------------------------------

        private static void AddToUndo(UndoEntry entry)
        {
            _undo.Add(entry);
            if (_undo.Count > MaxDepth) _undo.RemoveAt(0);
        }

        private static void RaiseChanged()
        {
            try { Changed?.Invoke(); } catch { }
        }

        private static void SetCell(dynamic cell, string formula)
        {
            if (string.IsNullOrEmpty(formula)) cell.ClearContents();
            else cell.Formula = formula;   // restores a real formula, or literal text on Text-formatted cells
        }

        private static string FormulaString(object v) =>
            v == null ? "" : Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? "";

        private static dynamic FindSheet(dynamic app, UndoEntry e)
        {
            try { return app.Workbooks(e.WorkbookName).Worksheets(e.SheetName); }
            catch { return null; }
        }

        private static void RegisterOnUndo(string text, string macroName)
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                app.OnUndo(text, macroName);
            }
            catch { /* not fatal — the stack still works via the commands */ }
        }
    }
}
