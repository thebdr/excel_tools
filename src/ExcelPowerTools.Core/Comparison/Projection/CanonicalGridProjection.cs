using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Projection
{
    public sealed class CanonicalGridProjectionOptions
    {
        public const int DefaultMaxMaterializedRows = 100000;
        public const int DefaultMaxMaterializedColumns = 4096;
        public const long DefaultMaxMaterializedCells = 5000000;

        public int MaxMaterializedRows { get; set; } = DefaultMaxMaterializedRows;
        public int MaxMaterializedColumns { get; set; } = DefaultMaxMaterializedColumns;
        public long MaxMaterializedCells { get; set; } = DefaultMaxMaterializedCells;
    }

    /// <summary>
    /// Bounded rectangular alignment view over an immutable sparse snapshot. The view retains the physical
    /// row/column offsets and original canonical cells, so inference strings can never replace exact evidence.
    /// </summary>
    public sealed class CanonicalGridProjection
    {
        private readonly IReadOnlyDictionary<long, CanonicalCellState> _physicalCells;

        internal CanonicalGridProjection(
            RevisionSnapshot source,
            AlignmentTable table,
            int firstPhysicalRow,
            int firstPhysicalColumn,
            IDictionary<long, CanonicalCellState> physicalCells)
        {
            Source = source;
            Table = table;
            FirstPhysicalRow = firstPhysicalRow;
            FirstPhysicalColumn = firstPhysicalColumn;
            _physicalCells = new ReadOnlyDictionary<long, CanonicalCellState>(
                new Dictionary<long, CanonicalCellState>(physicalCells));
        }

        public RevisionSnapshot Source { get; }
        public AlignmentTable Table { get; }
        /// <summary>One-based source row represented by table row zero; zero for an empty projection.</summary>
        public int FirstPhysicalRow { get; }
        /// <summary>One-based source column represented by table column zero; zero for an empty projection.</summary>
        public int FirstPhysicalColumn { get; }

        public int PhysicalRowForTableIndex(int tableRowIndex)
        {
            if (tableRowIndex < 0 || tableRowIndex >= Table.RowCount)
                throw new ArgumentOutOfRangeException(nameof(tableRowIndex));
            return checked(FirstPhysicalRow + tableRowIndex);
        }

        public int PhysicalColumnForTableIndex(int tableColumnIndex)
        {
            if (tableColumnIndex < 0 || tableColumnIndex >= Table.ColumnCount)
                throw new ArgumentOutOfRangeException(nameof(tableColumnIndex));
            return checked(FirstPhysicalColumn + tableColumnIndex);
        }

        public bool TryGetPhysicalCell(int oneBasedRow, int oneBasedColumn, out CanonicalCellState cell)
        {
            if (oneBasedRow <= 0) throw new ArgumentOutOfRangeException(nameof(oneBasedRow));
            if (oneBasedColumn <= 0) throw new ArgumentOutOfRangeException(nameof(oneBasedColumn));
            return _physicalCells.TryGetValue(CoordinateKey(oneBasedRow, oneBasedColumn), out cell);
        }

        /// <summary>
        /// Re-keys every physically stored canonical cell through confirmed logical row/column identities.
        /// Missing mappings are blocked rather than falling back to a volatile physical address.
        /// </summary>
        public RevisionSnapshot CreateAlignedSnapshot(
            string logicalSheetId,
            IReadOnlyDictionary<int, string> rowIdentities,
            IReadOnlyDictionary<int, string> columnIdentities,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            IReadOnlyList<AlignedCanonicalCellProjection> cells = CreateAlignedCells(
                logicalSheetId,
                rowIdentities,
                columnIdentities,
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            return new RevisionSnapshot(Source.RevisionId, cells.Select(item => item.Cell));
        }

        public IReadOnlyList<AlignedCanonicalCellProjection> CreateAlignedCells(
            string logicalSheetId,
            IReadOnlyDictionary<int, string> rowIdentities,
            IReadOnlyDictionary<int, string> columnIdentities,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (string.IsNullOrWhiteSpace(logicalSheetId))
                throw new ArgumentException("A logical sheet ID is required.", nameof(logicalSheetId));
            if (rowIdentities == null) throw new ArgumentNullException(nameof(rowIdentities));
            if (columnIdentities == null) throw new ArgumentNullException(nameof(columnIdentities));

            var aligned = new List<AlignedCanonicalCellProjection>(_physicalCells.Count);
            int processed = 0;
            foreach (KeyValuePair<long, CanonicalCellState> pair in _physicalCells.OrderBy(item => item.Key))
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                DecodeCoordinateKey(pair.Key, out int row, out int column);
                if (!rowIdentities.TryGetValue(row, out string rowIdentity) ||
                    string.IsNullOrWhiteSpace(rowIdentity))
                    throw new InvalidOperationException(
                        "No confirmed logical row identity exists for physical row " +
                        row.ToString(CultureInfo.InvariantCulture) + ".");
                if (!columnIdentities.TryGetValue(column, out string columnIdentity) ||
                    string.IsNullOrWhiteSpace(columnIdentity))
                    throw new InvalidOperationException(
                        "No confirmed logical column identity exists for physical column " +
                        column.ToString(CultureInfo.InvariantCulture) + ".");

                string subjectKey = StableId.Compute(
                    "aligned-cell-subject-v1",
                    logicalSheetId,
                    rowIdentity,
                    columnIdentity);
                aligned.Add(new AlignedCanonicalCellProjection(
                    CloneWithSubjectKey(pair.Value, subjectKey),
                    pair.Value.SubjectKey,
                    row,
                    column,
                    rowIdentity,
                    columnIdentity));
            }
            cancellationToken.ThrowIfCancellationRequested();
            return new ReadOnlyCollection<AlignedCanonicalCellProjection>(aligned);
        }

        private static CanonicalCellState CloneWithSubjectKey(CanonicalCellState source, string subjectKey)
        {
            if (source.HasFormula)
                return CanonicalCellState.Formula(
                    subjectKey,
                    source.FormulaText,
                    source.CachedResult,
                    source.DisplayText,
                    source.FormulaMetadata);
            if (source.ConstantValue.Kind == CanonicalValueKind.Blank)
                return CanonicalCellState.Blank(subjectKey);
            return CanonicalCellState.Constant(subjectKey, source.ConstantValue, source.DisplayText);
        }

        internal static long CoordinateKey(int row, int column) => ((long)row << 20) | (uint)column;

        private static void DecodeCoordinateKey(long key, out int row, out int column)
        {
            row = checked((int)(key >> 20));
            column = checked((int)(key & 0xFFFFF));
        }
    }

    public sealed class AlignedCanonicalCellProjection
    {
        internal AlignedCanonicalCellProjection(
            CanonicalCellState cell,
            string originalSubjectKey,
            int physicalRow,
            int physicalColumn,
            string rowIdentity,
            string columnIdentity)
        {
            Cell = cell;
            OriginalSubjectKey = originalSubjectKey;
            PhysicalRow = physicalRow;
            PhysicalColumn = physicalColumn;
            RowIdentity = rowIdentity;
            ColumnIdentity = columnIdentity;
        }

        public CanonicalCellState Cell { get; }
        public string OriginalSubjectKey { get; }
        public int PhysicalRow { get; }
        public int PhysicalColumn { get; }
        public string RowIdentity { get; }
        public string ColumnIdentity { get; }
    }

    public static class CanonicalGridProjectionBuilder
    {
        private const int ExcelMaxRows = 1048576;
        private const int ExcelMaxColumns = 16384;

        public static CanonicalGridProjection Build(
            RevisionSnapshot snapshot,
            CanonicalGridProjectionOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            options = options ?? new CanonicalGridProjectionOptions();
            ValidateOptions(options);

            CanonicalUsefulRange usefulRange = CanonicalUsefulRangeAnalyzer.Analyze(
                snapshot,
                cancellationToken);
            if (usefulRange.InvalidCoordinateCellCount > 0)
                throw new ArgumentException(
                    "The snapshot contains a non-positional subject key that cannot form an alignment grid: " +
                    usefulRange.FirstInvalidSubjectKey,
                    nameof(snapshot));

            var cells = new Dictionary<long, CanonicalCellState>();
            int processed = 0;
            foreach (CanonicalCellState cell in snapshot.Cells.Values)
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!TryParseSubjectCoordinate(cell.SubjectKey, out int row, out int column))
                    throw new ArgumentException(
                        "The snapshot contains a non-positional subject key that cannot form an alignment grid: " +
                        cell.SubjectKey,
                        nameof(snapshot));
                if (!usefulRange.Contains(row, column)) continue;
                long key = CanonicalGridProjection.CoordinateKey(row, column);
                if (cells.ContainsKey(key))
                    throw new ArgumentException("Two subject keys resolve to the same physical coordinate.", nameof(snapshot));
                cells.Add(key, cell);
            }

            if (usefulRange.IsEmpty)
                return new CanonicalGridProjection(
                    snapshot,
                    new AlignmentTable(Array.Empty<string[]>()),
                    0,
                    0,
                    cells);

            int rowCount = checked((int)usefulRange.RowCount);
            int columnCount = checked((int)usefulRange.ColumnCount);
            long materializedCells = usefulRange.GridCellCount;
            if (rowCount > options.MaxMaterializedRows ||
                columnCount > options.MaxMaterializedColumns ||
                materializedCells > options.MaxMaterializedCells)
                throw new InvalidOperationException(
                    "The content-bounded alignment view for useful range " + usefulRange.DisplayAddress +
                    " is too large (" +
                    rowCount.ToString("N0", CultureInfo.InvariantCulture) + " rows x " +
                    columnCount.ToString("N0", CultureInfo.InvariantCulture) + " columns = " +
                    materializedCells.ToString("N0", CultureInfo.InvariantCulture) + " cells). Per-sheet limits: " +
                    options.MaxMaterializedRows.ToString("N0", CultureInfo.InvariantCulture) + " rows, " +
                    options.MaxMaterializedColumns.ToString("N0", CultureInfo.InvariantCulture) + " columns, and " +
                    options.MaxMaterializedCells.ToString("N0", CultureInfo.InvariantCulture) +
                    " cells. Blank/formatting-only cells outside the useful range were already excluded. " +
                    "Inspect the range for a stray nonblank cell or split the comparison for now.");

            var rows = new List<string[]>(rowCount);
            for (int tableRow = 0; tableRow < rowCount; tableRow++)
            {
                if ((tableRow & 0x1F) == 0) cancellationToken.ThrowIfCancellationRequested();
                var values = new string[columnCount];
                int physicalRow = usefulRange.FirstRow + tableRow;
                for (int tableColumn = 0; tableColumn < columnCount; tableColumn++)
                {
                    int physicalColumn = usefulRange.FirstColumn + tableColumn;
                    if (cells.TryGetValue(
                        CanonicalGridProjection.CoordinateKey(physicalRow, physicalColumn),
                        out CanonicalCellState cell))
                        values[tableColumn] = AlignmentText(cell);
                }
                rows.Add(values);
            }

            cancellationToken.ThrowIfCancellationRequested();

            return new CanonicalGridProjection(
                snapshot,
                new AlignmentTable(rows),
                usefulRange.FirstRow,
                usefulRange.FirstColumn,
                cells);
        }

        public static bool TryParseSubjectCoordinate(string subjectKey, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrEmpty(subjectKey)) return false;

            if ((subjectKey[0] == 'R' || subjectKey[0] == 'r') &&
                TryParseR1C1(subjectKey, out row, out column))
                return IsExcelCoordinate(row, column);

            int letterCount = 0;
            while (letterCount < subjectKey.Length &&
                   ((subjectKey[letterCount] >= 'A' && subjectKey[letterCount] <= 'Z') ||
                    (subjectKey[letterCount] >= 'a' && subjectKey[letterCount] <= 'z')))
                letterCount++;
            if (letterCount == 0 || letterCount == subjectKey.Length) return false;

            long parsedColumn = 0;
            for (int index = 0; index < letterCount; index++)
            {
                char normalized = char.ToUpperInvariant(subjectKey[index]);
                parsedColumn = checked(parsedColumn * 26 + (normalized - 'A' + 1));
                if (parsedColumn > ExcelMaxColumns) return false;
            }
            if (!int.TryParse(
                    subjectKey.Substring(letterCount),
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out row))
                return false;
            column = (int)parsedColumn;
            return IsExcelCoordinate(row, column);
        }

        private static bool TryParseR1C1(string subjectKey, out int row, out int column)
        {
            row = 0;
            column = 0;
            int c = subjectKey.IndexOf('C', 1);
            if (c < 0) c = subjectKey.IndexOf('c', 1);
            if (c <= 1 || c == subjectKey.Length - 1) return false;
            return int.TryParse(
                       subjectKey.Substring(1, c - 1),
                       NumberStyles.None,
                       CultureInfo.InvariantCulture,
                       out row) &&
                   int.TryParse(
                       subjectKey.Substring(c + 1),
                       NumberStyles.None,
                       CultureInfo.InvariantCulture,
                       out column);
        }

        private static bool IsExcelCoordinate(int row, int column) =>
            row > 0 && row <= ExcelMaxRows && column > 0 && column <= ExcelMaxColumns;

        private static string AlignmentText(CanonicalCellState cell)
        {
            if (cell.HasFormula)
                return cell.FormulaText.StartsWith("=", StringComparison.Ordinal)
                    ? cell.FormulaText
                    : "=" + cell.FormulaText;
            return cell.ConstantValue.ExactText;
        }

        private static void ValidateOptions(CanonicalGridProjectionOptions options)
        {
            if (options.MaxMaterializedRows <= 0 || options.MaxMaterializedRows > ExcelMaxRows)
                throw new ArgumentOutOfRangeException(nameof(options.MaxMaterializedRows));
            if (options.MaxMaterializedColumns <= 0 || options.MaxMaterializedColumns > ExcelMaxColumns)
                throw new ArgumentOutOfRangeException(nameof(options.MaxMaterializedColumns));
            if (options.MaxMaterializedCells <= 0)
                throw new ArgumentOutOfRangeException(nameof(options.MaxMaterializedCells));
        }
    }
}
