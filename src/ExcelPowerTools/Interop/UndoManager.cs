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

    /// <summary>One undoable Power-Tools operation: which cells it touched and what they held before.</summary>
    internal sealed class UndoEntry
    {
        public string Description;
        public string WorkbookName;
        public string SheetName;
        public List<UndoBlock> Blocks = new List<UndoBlock>();
        public int CellCount;
    }

    /// <summary>
    /// Undo for add-in operations — the thing VBA never gave us. Mutating commands snapshot the cells
    /// they change (as .Formula strings, so real formulas survive) and push here. Application.OnUndo is
    /// registered after each push, so NATIVE Ctrl+Z undoes the operation immediately after it runs (and
    /// then toggles to Redo). Older operations stay reachable via PT_UndoLast/PT_RedoLast (Ctrl+Alt+Z/Y),
    /// up to 20 deep. Cell contents are restored; formatting changes are not tracked.
    /// </summary>
    internal static class UndoManager
    {
        private const int MaxDepth = 20;
        private static readonly List<UndoEntry> _undo = new List<UndoEntry>();
        private static readonly List<UndoEntry> _redo = new List<UndoEntry>();

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

        /// <summary>Publishes a completed operation: makes native Ctrl+Z undo it (one-shot) and stacks it.</summary>
        public static void Push(UndoEntry entry)
        {
            if (entry == null || entry.CellCount == 0) return;

            _undo.Add(entry);
            if (_undo.Count > MaxDepth) _undo.RemoveAt(0);
            _redo.Clear();

            RegisterOnUndo("Undo " + entry.Description + " (Power Tools)", "PT_UndoLast");
        }

        // ---- undo / redo --------------------------------------------------------------------------

        public static void UndoLast()
        {
            if (_undo.Count == 0) { Dialogs.Info("Nothing to undo."); return; }

            UndoEntry e = _undo[_undo.Count - 1];
            _undo.RemoveAt(_undo.Count - 1);

            UndoEntry inverse = SwapApply(e);
            if (inverse == null) return;   // sheet gone — step dropped

            _redo.Add(inverse);
            if (_redo.Count > MaxDepth) _redo.RemoveAt(0);
            RegisterOnUndo("Redo " + e.Description + " (Power Tools)", "PT_RedoLast");
        }

        public static void RedoLast()
        {
            if (_redo.Count == 0) { Dialogs.Info("Nothing to redo."); return; }

            UndoEntry e = _redo[_redo.Count - 1];
            _redo.RemoveAt(_redo.Count - 1);

            UndoEntry inverse = SwapApply(e);
            if (inverse == null) return;

            _undo.Add(inverse);
            if (_undo.Count > MaxDepth) _undo.RemoveAt(0);
            RegisterOnUndo("Undo " + e.Description + " (Power Tools)", "PT_UndoLast");
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
            };

            using (new ExcelStateGuard(manualCalculation: true))
            {
                foreach (UndoBlock b in e.Blocks)
                {
                    int rows = b.Formulas.GetLength(0), cols = b.Formulas.GetLength(1);
                    dynamic range = sheet.Range(sheet.Cells(b.Row, b.Col), sheet.Cells(b.Row + rows - 1, b.Col + cols - 1));

                    CaptureBlock(inverse, range);   // current content, so the operation can be re-done

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

            try { sheet.Activate(); } catch { }
            return inverse;
        }

        // ---- plumbing -----------------------------------------------------------------------------

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
