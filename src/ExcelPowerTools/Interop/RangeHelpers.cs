using System;

namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// Port of the reusable range helpers from the VBA xlReferences module.
    /// All calls are late-bound so they work against any installed Excel version.
    /// </summary>
    internal static class RangeHelpers
    {
        /// <summary>Last used row on a sheet (searching upward from the bottom), like LastCellOccupied "ROW".</summary>
        public static long LastRow(dynamic sheet)
        {
            dynamic found = FindLast(sheet, Xl.xlByRows);
            return found == null ? 1 : (long)(int)found.Row;
        }

        /// <summary>Last used column on a sheet (searching left from the right edge), like LastCellOccupied "COLUMN".</summary>
        public static long LastColumn(dynamic sheet)
        {
            dynamic found = FindLast(sheet, Xl.xlByColumns);
            return found == null ? 1 : (long)(int)found.Column;
        }

        /// <summary>Bottom-most non-empty row within a single column, like the VBA LastCellInColumn.</summary>
        public static long LastRowInColumn(dynamic sheet, int column)
        {
            try
            {
                dynamic col = sheet.Cells(1, column).EntireColumn;
                dynamic found = col.Find(
                    "*", Type.Missing, Xl.xlFormulas, Xl.xlPart, Xl.xlByRows, Xl.xlPrevious);
                return found == null ? 1 : (long)(int)found.Row;
            }
            catch
            {
                return 1;
            }
        }

        /// <summary>True when the row is hidden (filtered or manually); false on any COM error.</summary>
        public static bool IsRowHidden(dynamic sheet, int row)
        {
            try { return (bool)sheet.Rows(row).Hidden; } catch { return false; }
        }

        /// <summary>True if a worksheet with the given name exists in the workbook.</summary>
        public static bool WorksheetExists(dynamic workbook, string name)
        {
            try
            {
                dynamic ws = workbook.Worksheets[name];
                return ws != null;
            }
            catch
            {
                return false;
            }
        }

        private static dynamic FindLast(dynamic sheet, int searchOrder)
        {
            // Cells.Find("*", After, LookIn, LookAt, SearchOrder, SearchDirection:=xlPrevious)
            return sheet.Cells.Find(
                "*",
                Type.Missing,
                Xl.xlFormulas,
                Xl.xlPart,
                searchOrder,
                Xl.xlPrevious);
        }
    }
}
