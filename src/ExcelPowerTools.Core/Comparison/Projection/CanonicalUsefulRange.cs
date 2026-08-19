using System;
using System.Globalization;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Projection
{
    /// <summary>
    /// Content-aware rectangular boundary for one sparse canonical worksheet snapshot. Formula cells and typed
    /// constants establish the boundary; physically stored blank cells are retained only when they fall inside
    /// that rectangle. This prevents formatting-only tails from expanding an alignment grid while preserving
    /// blank-versus-missing evidence in the comparable workspace.
    /// </summary>
    public sealed class CanonicalUsefulRange
    {
        internal CanonicalUsefulRange(
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn,
            long totalStoredCellCount,
            long parsedStoredCellCount,
            long contentCellCount,
            long retainedStoredCellCount,
            long ignoredOutsideRangeCellCount,
            long invalidCoordinateCellCount,
            string firstInvalidSubjectKey)
        {
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            TotalStoredCellCount = totalStoredCellCount;
            ParsedStoredCellCount = parsedStoredCellCount;
            ContentCellCount = contentCellCount;
            RetainedStoredCellCount = retainedStoredCellCount;
            IgnoredOutsideRangeCellCount = ignoredOutsideRangeCellCount;
            InvalidCoordinateCellCount = invalidCoordinateCellCount;
            FirstInvalidSubjectKey = firstInvalidSubjectKey;
        }

        public bool IsEmpty => ContentCellCount == 0;
        public int FirstRow { get; }
        public int LastRow { get; }
        public int FirstColumn { get; }
        public int LastColumn { get; }
        public long TotalStoredCellCount { get; }
        public long ParsedStoredCellCount { get; }
        public long ContentCellCount { get; }
        public long RetainedStoredCellCount { get; }
        public long IgnoredOutsideRangeCellCount { get; }
        public long InvalidCoordinateCellCount { get; }
        public string FirstInvalidSubjectKey { get; }
        public long RowCount => IsEmpty ? 0L : checked((long)LastRow - FirstRow + 1L);
        public long ColumnCount => IsEmpty ? 0L : checked((long)LastColumn - FirstColumn + 1L);
        public long GridCellCount => checked(RowCount * ColumnCount);

        public string DisplayAddress => IsEmpty
            ? "(no formula/value cells)"
            : ToColumnName(FirstColumn) + FirstRow.ToString(CultureInfo.InvariantCulture) + ":" +
              ToColumnName(LastColumn) + LastRow.ToString(CultureInfo.InvariantCulture);

        public bool Contains(int oneBasedRow, int oneBasedColumn) =>
            !IsEmpty &&
            oneBasedRow >= FirstRow && oneBasedRow <= LastRow &&
            oneBasedColumn >= FirstColumn && oneBasedColumn <= LastColumn;

        private static string ToColumnName(int oneBasedColumn)
        {
            if (oneBasedColumn <= 0 || oneBasedColumn > 16384)
                throw new ArgumentOutOfRangeException(nameof(oneBasedColumn));

            var characters = new char[3];
            int position = characters.Length;
            int remaining = oneBasedColumn;
            while (remaining > 0)
            {
                remaining--;
                characters[--position] = (char)('A' + remaining % 26);
                remaining /= 26;
            }
            return new string(characters, position, characters.Length - position);
        }
    }

    public static class CanonicalUsefulRangeAnalyzer
    {
        public static CanonicalUsefulRange Analyze(
            RevisionSnapshot snapshot,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));

            int firstRow = int.MaxValue;
            int lastRow = 0;
            int firstColumn = int.MaxValue;
            int lastColumn = 0;
            long parsedStoredCells = 0;
            long contentCells = 0;
            long invalidCoordinateCells = 0;
            string firstInvalidSubjectKey = null;
            int processed = 0;

            foreach (CanonicalCellState cell in snapshot.Cells.Values)
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(
                        cell.SubjectKey,
                        out int row,
                        out int column))
                {
                    invalidCoordinateCells++;
                    if (firstInvalidSubjectKey == null) firstInvalidSubjectKey = cell.SubjectKey;
                    continue;
                }

                parsedStoredCells++;
                if (!ContainsFormulaOrValue(cell)) continue;

                contentCells++;
                firstRow = Math.Min(firstRow, row);
                lastRow = Math.Max(lastRow, row);
                firstColumn = Math.Min(firstColumn, column);
                lastColumn = Math.Max(lastColumn, column);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (contentCells == 0)
            {
                return new CanonicalUsefulRange(
                    0,
                    0,
                    0,
                    0,
                    snapshot.Cells.Count,
                    parsedStoredCells,
                    0,
                    0,
                    parsedStoredCells,
                    invalidCoordinateCells,
                    firstInvalidSubjectKey);
            }

            long retainedStoredCells = 0;
            long ignoredOutsideRangeCells = 0;
            processed = 0;
            foreach (CanonicalCellState cell in snapshot.Cells.Values)
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(
                        cell.SubjectKey,
                        out int row,
                        out int column))
                    continue;

                if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn)
                    retainedStoredCells++;
                else
                    ignoredOutsideRangeCells++;
            }

            cancellationToken.ThrowIfCancellationRequested();
            return new CanonicalUsefulRange(
                firstRow,
                lastRow,
                firstColumn,
                lastColumn,
                snapshot.Cells.Count,
                parsedStoredCells,
                contentCells,
                retainedStoredCells,
                ignoredOutsideRangeCells,
                invalidCoordinateCells,
                firstInvalidSubjectKey);
        }

        public static bool ContainsFormulaOrValue(CanonicalCellState cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            return cell.HasFormula || cell.ConstantValue.Kind != CanonicalValueKind.Blank;
        }
    }
}
