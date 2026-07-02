using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;

namespace ExcelPowerTools.Functions
{
    /// <summary>Worksheet UDFs exposed by the add-in.</summary>
    public static class WorksheetFunctions
    {
        /// <summary>
        /// Port of the VBA vbaIndexMultiValues UDF: returns the values of a single-column (or single-row)
        /// range picked out by a list of 1-based positions.
        /// e.g. =PT.INDEXMULTI(A1:A10, {1;3;5}) returns the 1st, 3rd and 5th values.
        /// Deliberate deviation from the VBA: an out-of-range position yields #N/A for that one cell (like INDEX),
        /// instead of failing the whole formula with #VALUE! as the unguarded VBA UDF did.
        /// </summary>
        [ExcelFunction(Name = "PT.INDEXMULTI",
            Description = "Returns values from a single-column/row range at the given 1-based positions.")]
        public static object[,] IndexMulti(
            [ExcelArgument(Name = "source", Description = "single column or row of values")] object[,] source,
            [ExcelArgument(Name = "positions", Description = "1-based positions to return")] object positions)
        {
            List<object> flat = Flatten(source);
            List<int> indices = ReadPositions(positions);

            var result = new object[Math.Max(indices.Count, 1), 1];
            for (int k = 0; k < indices.Count; k++)
            {
                int idx = indices[k] - 1;   // 1-based -> 0-based
                result[k, 0] = (idx >= 0 && idx < flat.Count) ? flat[idx] : ExcelError.ExcelErrorNA;
            }
            if (indices.Count == 0) result[0, 0] = ExcelError.ExcelErrorValue;
            return result;
        }

        private static List<object> Flatten(object[,] source)
        {
            int rows = source.GetLength(0);
            int cols = source.GetLength(1);
            int rowBase = source.GetLowerBound(0);
            int colBase = source.GetLowerBound(1);

            var flat = new List<object>(rows * cols);
            if (cols == 1)
            {
                for (int r = 0; r < rows; r++) flat.Add(source[rowBase + r, colBase]);
            }
            else if (rows == 1)
            {
                for (int c = 0; c < cols; c++) flat.Add(source[rowBase, colBase + c]);
            }
            else
            {
                // fall back to the first column for a rectangular range
                for (int r = 0; r < rows; r++) flat.Add(source[rowBase + r, colBase]);
            }
            return flat;
        }

        private static List<int> ReadPositions(object positions)
        {
            var result = new List<int>();

            switch (positions)
            {
                case object[,] grid:
                    foreach (object o in grid)
                        if (TryToInt(o, out int v)) result.Add(v);
                    break;
                case object[] arr:
                    foreach (object o in arr)
                        if (TryToInt(o, out int v)) result.Add(v);
                    break;
                default:
                    if (TryToInt(positions, out int single)) result.Add(single);
                    break;
            }
            return result;
        }

        private static bool TryToInt(object o, out int value)
        {
            value = 0;
            if (o == null || o is ExcelEmpty || o is ExcelMissing || o is ExcelError) return false;
            try
            {
                value = (int)Math.Round(Convert.ToDouble(o, CultureInfo.InvariantCulture));
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
