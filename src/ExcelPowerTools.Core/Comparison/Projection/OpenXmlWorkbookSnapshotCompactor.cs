using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Projection
{
    /// <summary>
    /// Original physical-cell evidence and the content-bounded range retained for one worksheet.
    /// The counts describe the untrimmed adapter snapshot even when the returned workbook snapshot is compacted.
    /// </summary>
    public sealed class OpenXmlWorksheetCompaction
    {
        internal OpenXmlWorksheetCompaction(string sheetName, string partUri, CanonicalUsefulRange range)
        {
            SheetName = sheetName ?? "";
            PartUri = partUri ?? "";
            Range = range ?? throw new ArgumentNullException(nameof(range));
        }

        public string SheetName { get; }
        public string PartUri { get; }
        public CanonicalUsefulRange Range { get; }
    }

    /// <summary>
    /// A workbook snapshot whose formula/value evidence and interior blank evidence are unchanged, but whose
    /// formatting-only blank tails have been released from memory after they were parsed and audited.
    /// </summary>
    public sealed class OpenXmlWorkbookCompactionResult
    {
        private readonly ReadOnlyCollection<OpenXmlWorksheetCompaction> _worksheets;

        internal OpenXmlWorkbookCompactionResult(
            OpenXmlWorkbookSnapshot snapshot,
            IList<OpenXmlWorksheetCompaction> worksheets)
        {
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            _worksheets = new ReadOnlyCollection<OpenXmlWorksheetCompaction>(
                new List<OpenXmlWorksheetCompaction>(
                    worksheets ?? Array.Empty<OpenXmlWorksheetCompaction>()));
        }

        public OpenXmlWorkbookSnapshot Snapshot { get; }
        public IReadOnlyList<OpenXmlWorksheetCompaction> Worksheets => _worksheets;
    }

    public static class OpenXmlWorkbookSnapshotCompactor
    {
        /// <summary>
        /// Retains every formula/value cell, every physically stored blank inside the first-to-last content
        /// rectangle, and every malformed/non-positional cell so unsupported evidence cannot be hidden. Only
        /// valid, physically stored blank cells outside the useful range are discarded.
        /// </summary>
        public static OpenXmlWorkbookCompactionResult CompactToUsefulRanges(
            OpenXmlWorkbookSnapshot source,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            var compactedSheets = new List<OpenXmlWorksheetSnapshot>(source.Sheets.Count);
            var worksheetResults = new List<OpenXmlWorksheetCompaction>(source.Sheets.Count);
            foreach (OpenXmlWorksheetSnapshot sheet in source.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CanonicalUsefulRange range = CanonicalUsefulRangeAnalyzer.Analyze(
                    sheet.Snapshot,
                    cancellationToken);
                worksheetResults.Add(new OpenXmlWorksheetCompaction(
                    sheet.Name,
                    sheet.PartUri,
                    range));

                if (range.IgnoredOutsideRangeCellCount == 0)
                {
                    compactedSheets.Add(sheet);
                    continue;
                }

                long expected = range.RetainedStoredCellCount + range.InvalidCoordinateCellCount;
                int capacity = expected <= 0
                    ? 0
                    : (int)Math.Min(expected, int.MaxValue);
                var retainedCells = new List<CanonicalCellState>(capacity);
                int processed = 0;
                foreach (CanonicalCellState cell in sheet.Snapshot.Cells.Values)
                {
                    if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(
                            cell.SubjectKey,
                            out int row,
                            out int column) ||
                        range.Contains(row, column))
                        retainedCells.Add(cell);
                }

                cancellationToken.ThrowIfCancellationRequested();
                compactedSheets.Add(new OpenXmlWorksheetSnapshot(
                    sheet.Name,
                    sheet.SheetId,
                    sheet.RelationshipId,
                    sheet.PartUri,
                    sheet.CodeName,
                    sheet.Visibility,
                    sheet.Kind,
                    new RevisionSnapshot(sheet.Snapshot.RevisionId, retainedCells),
                    new List<OpenXmlFormulaObservation>(sheet.FormulaObservations),
                    new List<OpenXmlTableFormulaObservation>(sheet.TableFormulaObservations),
                    new List<OpenXmlCoverageMessage>(sheet.CoverageMessages)));
            }

            cancellationToken.ThrowIfCancellationRequested();
            var compacted = new OpenXmlWorkbookSnapshot(
                source.RevisionId,
                source.ByteContentId,
                source.DateSystem,
                source.HasVbaProject,
                compactedSheets,
                new List<OpenXmlCoverageMessage>(source.CoverageMessages));
            return new OpenXmlWorkbookCompactionResult(compacted, worksheetResults);
        }
    }
}
