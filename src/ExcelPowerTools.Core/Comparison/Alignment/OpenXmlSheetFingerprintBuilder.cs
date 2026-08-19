using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    /// <summary>
    /// Builds deterministic sheet-identity evidence from an already immutable OOXML snapshot. It never opens
    /// Excel and never hashes coverage placeholders as trustworthy identity evidence. Fingerprints may use exact
    /// stored formula/value evidence for review, while stricter recalculation or spill uncertainty remains visible
    /// to the project engine and continues to block merge eligibility.
    /// </summary>
    public static class OpenXmlSheetFingerprintBuilder
    {
        private const string StructureDomain = "ooxml-sheet-structure-fingerprint-v1";
        private const string ContentDomain = "ooxml-sheet-content-fingerprint-v1";
        private const string FormulaDomain = "ooxml-sheet-formula-fingerprint-v1";
        private const string HeaderDomain = "ooxml-sheet-header-fingerprint-v1";

        public static OpenXmlWorkbookSheetFingerprintSet Build(
            OpenXmlWorkbookSnapshot workbook,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            cancellationToken.ThrowIfCancellationRequested();

            bool workbookCellCoverageComplete = !workbook.CoverageMessages.Any(
                message => message.AffectsExactCellComparison);
            string[] workbookCoverage = DistinctCodes(workbook.CoverageMessages.Select(message => message.Code));
            var evidence = new List<OpenXmlSheetFingerprintEvidence>(workbook.Sheets.Count);
            for (int index = 0; index < workbook.Sheets.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                evidence.Add(BuildSheet(
                    workbook,
                    workbook.Sheets[index],
                    index,
                    workbookCellCoverageComplete,
                    cancellationToken));
            }

            return new OpenXmlWorkbookSheetFingerprintSet(workbook, evidence, workbookCoverage);
        }

        private static OpenXmlSheetFingerprintEvidence BuildSheet(
            OpenXmlWorkbookSnapshot workbook,
            OpenXmlWorksheetSnapshot sheet,
            int occurrence,
            bool workbookCellCoverageComplete,
            CancellationToken cancellationToken)
        {
            CanonicalUsefulRange usefulRange = CanonicalUsefulRangeAnalyzer.Analyze(
                sheet.Snapshot,
                cancellationToken);
            bool invalidReference;
            List<ObservedCell> cells = ParseCells(
                sheet,
                usefulRange,
                cancellationToken,
                out invalidReference);
            invalidReference = invalidReference || usefulRange.InvalidCoordinateCellCount > 0;
            int usedRows = usefulRange.LastRow;
            int usedColumns = usefulRange.LastColumn;
            long storedCellCount = usefulRange.RetainedStoredCellCount;
            bool complete = workbookCellCoverageComplete &&
                            sheet.Kind == OpenXmlPhysicalSheetKind.Worksheet &&
                            sheet.CanProjectExactStoredFormulaValueEvidence &&
                            !invalidReference;

            string structure = null;
            string content = null;
            string header = null;
            string formula = null;
            if (complete && cells.Count > 0)
            {
                structure = BuildStructureFingerprint(cells, cancellationToken);
                content = BuildContentFingerprint(cells, cancellationToken);
                formula = BuildFormulaFingerprint(cells, cancellationToken);
                header = BuildHeaderFingerprint(cells, cancellationToken);
            }

            var coverageCodes = new List<string>();
            coverageCodes.AddRange(workbook.CoverageMessages.Select(message => message.Code));
            coverageCodes.AddRange(sheet.CoverageMessages.Select(message => message.Code));
            if (invalidReference)
                coverageCodes.Add(OpenXmlSheetFingerprintCoverageCodes.InvalidCanonicalCellReference);

            var fingerprint = new WorkbookSheetFingerprint(
                BuildSheetKey(workbook, sheet, occurrence),
                sheet.Name,
                MapKind(sheet.Kind),
                sheet.CodeName,
                addInIdentity: null,
                structureFingerprint: structure,
                contentFingerprint: content,
                headerFingerprint: header,
                formulaFingerprint: formula,
                usedRowCount: usedRows,
                usedColumnCount: usedColumns,
                storedCellCount: storedCellCount);
            return new OpenXmlSheetFingerprintEvidence(
                sheet,
                fingerprint,
                complete,
                DistinctCodes(coverageCodes));
        }

        private static string BuildSheetKey(
            OpenXmlWorkbookSnapshot workbook,
            OpenXmlWorksheetSnapshot sheet,
            int occurrence)
        {
            if (!string.IsNullOrWhiteSpace(sheet.Snapshot.RevisionId)) return sheet.Snapshot.RevisionId;
            return workbook.RevisionId + "/sheet-occurrence:" +
                   occurrence.ToString("D6", CultureInfo.InvariantCulture);
        }

        private static WorkbookSheetKind MapKind(OpenXmlPhysicalSheetKind kind) =>
            kind == OpenXmlPhysicalSheetKind.Worksheet
                ? WorkbookSheetKind.Worksheet
                : WorkbookSheetKind.Unknown;

        private static List<ObservedCell> ParseCells(
            OpenXmlWorksheetSnapshot sheet,
            CanonicalUsefulRange usefulRange,
            CancellationToken cancellationToken,
            out bool invalidReference)
        {
            invalidReference = false;
            var result = new List<ObservedCell>(checked((int)usefulRange.RetainedStoredCellCount));
            int processed = 0;
            foreach (CanonicalCellState cell in sheet.Snapshot.Cells.Values)
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!TryParseCellReference(cell.SubjectKey, out int row, out int column))
                {
                    invalidReference = true;
                    continue;
                }
                if (!usefulRange.Contains(row, column)) continue;
                result.Add(new ObservedCell(row, column, cell));
            }
            cancellationToken.ThrowIfCancellationRequested();
            result.Sort(CompareObservedCells);
            cancellationToken.ThrowIfCancellationRequested();
            return result;
        }

        private static string BuildStructureFingerprint(
            IReadOnlyList<ObservedCell> cells,
            CancellationToken cancellationToken)
        {
            using (var writer = new FingerprintWriter(StructureDomain))
            {
                writer.Write(cells.Count.ToString(CultureInfo.InvariantCulture));
                for (int index = 0; index < cells.Count; index++)
                {
                    if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                    ObservedCell observed = cells[index];
                    CanonicalCellState cell = observed.Cell;
                    writer.Write(cell.SubjectKey);
                    writer.Write(cell.HasFormula ? "formula" : "constant");
                    writer.Write(cell.HasFormula
                        ? cell.FormulaMetadata
                        : ((int)cell.ConstantValue.Kind).ToString(CultureInfo.InvariantCulture));
                }
                return writer.Complete();
            }
        }

        private static string BuildContentFingerprint(
            IReadOnlyList<ObservedCell> cells,
            CancellationToken cancellationToken)
        {
            using (var writer = new FingerprintWriter(ContentDomain))
            {
                writer.Write(cells.Count.ToString(CultureInfo.InvariantCulture));
                for (int index = 0; index < cells.Count; index++)
                {
                    if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                    ObservedCell observed = cells[index];
                    CanonicalCellState cell = observed.Cell;
                    writer.Write(cell.SubjectKey);
                    if (cell.HasFormula)
                    {
                        writer.Write("formula");
                        writer.Write(cell.FormulaText);
                        writer.Write(cell.FormulaMetadata);
                    }
                    else
                    {
                        writer.Write("constant");
                        WriteCanonicalValue(writer, cell.ConstantValue);
                    }
                }
                return writer.Complete();
            }
        }

        private static string BuildFormulaFingerprint(
            IReadOnlyList<ObservedCell> cells,
            CancellationToken cancellationToken)
        {
            int formulaCount = 0;
            for (int index = 0; index < cells.Count; index++)
            {
                if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (cells[index].Cell.HasFormula) formulaCount++;
            }
            if (formulaCount == 0) return null;

            using (var writer = new FingerprintWriter(FormulaDomain))
            {
                writer.Write(formulaCount.ToString(CultureInfo.InvariantCulture));
                for (int index = 0; index < cells.Count; index++)
                {
                    if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                    ObservedCell observed = cells[index];
                    CanonicalCellState cell = observed.Cell;
                    if (!cell.HasFormula) continue;
                    writer.Write(cell.SubjectKey);
                    writer.Write(cell.FormulaText);
                    writer.Write(cell.FormulaMetadata);
                }
                return writer.Complete();
            }
        }

        private static string BuildHeaderFingerprint(
            IReadOnlyList<ObservedCell> cells,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int firstRow = cells[0].Row;
            List<ObservedCell> header = cells.Where(cell => cell.Row == firstRow).ToList();
            if (header.Count < 2) return null;

            var normalized = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < header.Count; index++)
            {
                ObservedCell item = header[index];
                CanonicalCellState cell = item.Cell;
                if (cell.HasFormula || cell.ConstantValue.Kind != CanonicalValueKind.Text ||
                    string.IsNullOrWhiteSpace(cell.ConstantValue.ExactText))
                    return null;
                if (index > 0 && item.Column != header[index - 1].Column + 1) return null;
                if (!normalized.Add(cell.ConstantValue.ExactText.Trim())) return null;
            }

            int firstColumn = header[0].Column;
            int lastColumn = header[header.Count - 1].Column;
            IGrouping<int, ObservedCell> firstDataRow = cells
                .Where(cell => cell.Row > firstRow)
                .GroupBy(cell => cell.Row)
                .OrderBy(group => group.Key)
                .FirstOrDefault();
            if (firstDataRow == null) return null;

            List<ObservedCell> dataCells = firstDataRow
                .Where(cell => cell.Column >= firstColumn && cell.Column <= lastColumn)
                .ToList();
            cancellationToken.ThrowIfCancellationRequested();
            if (dataCells.Count < (header.Count + 1) / 2) return null;
            bool hasTypeTransition = dataCells.Any(item =>
                item.Cell.HasFormula ||
                (item.Cell.ConstantValue.Kind != CanonicalValueKind.Text &&
                 item.Cell.ConstantValue.Kind != CanonicalValueKind.Blank));
            if (!hasTypeTransition) return null;

            using (var writer = new FingerprintWriter(HeaderDomain))
            {
                writer.Write(header.Count.ToString(CultureInfo.InvariantCulture));
                foreach (ObservedCell item in header)
                {
                    writer.Write(item.Column.ToString(CultureInfo.InvariantCulture));
                    writer.Write(item.Cell.ConstantValue.ExactText);
                }
                return writer.Complete();
            }
        }

        private static void WriteCanonicalValue(FingerprintWriter writer, CanonicalValue value)
        {
            value = value ?? CanonicalValue.Blank;
            writer.Write(((int)value.Kind).ToString(CultureInfo.InvariantCulture));
            writer.Write(value.ExactText);
            writer.Write(((int)value.DateSystem).ToString(CultureInfo.InvariantCulture));
        }

        private static bool TryParseCellReference(string value, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrEmpty(value)) return false;

            int index = 0;
            while (index < value.Length && value[index] >= 'A' && value[index] <= 'Z')
            {
                column = checked(column * 26 + value[index] - 'A' + 1);
                index++;
            }
            if (index == 0 || index == value.Length || column > 16384) return false;

            int rowStart = index;
            while (index < value.Length && value[index] >= '0' && value[index] <= '9')
            {
                row = checked(row * 10 + value[index] - '0');
                index++;
            }
            return rowStart < index && index == value.Length && row > 0 && row <= 1048576;
        }

        private static int CompareObservedCells(ObservedCell left, ObservedCell right)
        {
            int row = left.Row.CompareTo(right.Row);
            if (row != 0) return row;
            int column = left.Column.CompareTo(right.Column);
            return column != 0
                ? column
                : string.Compare(left.Cell.SubjectKey, right.Cell.SubjectKey, StringComparison.Ordinal);
        }

        private static string[] DistinctCodes(IEnumerable<string> codes) =>
            codes
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(code => code, StringComparer.Ordinal)
                .ToArray();

        private sealed class ObservedCell
        {
            internal ObservedCell(int row, int column, CanonicalCellState cell)
            {
                Row = row;
                Column = column;
                Cell = cell;
            }

            internal int Row { get; }
            internal int Column { get; }
            internal CanonicalCellState Cell { get; }
        }

        private sealed class FingerprintWriter : IDisposable
        {
            private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);
            private readonly SHA256 _hash = SHA256.Create();
            private readonly CryptoStream _stream;
            private bool _completed;

            internal FingerprintWriter(string domain)
            {
                _stream = new CryptoStream(Stream.Null, _hash, CryptoStreamMode.Write);
                Write("ExcelPowerTools.Comparison.OpenXmlSheetFingerprint");
                Write("sha256-lp-utf8-v1");
                Write(domain);
            }

            internal void Write(string value)
            {
                if (_completed) throw new InvalidOperationException("The fingerprint is already complete.");
                if (value == null)
                {
                    WriteLength(-1);
                    return;
                }
                byte[] bytes = StrictUtf8.GetBytes(value);
                WriteLength(bytes.Length);
                _stream.Write(bytes, 0, bytes.Length);
            }

            internal string Complete()
            {
                if (!_completed)
                {
                    _stream.FlushFinalBlock();
                    _completed = true;
                }
                return ToHex(_hash.Hash);
            }

            public void Dispose()
            {
                if (!_completed)
                {
                    _stream.FlushFinalBlock();
                    _completed = true;
                }
                _stream.Dispose();
                _hash.Dispose();
            }

            private void WriteLength(int value)
            {
                var bytes = new[]
                {
                    (byte)(value >> 24),
                    (byte)(value >> 16),
                    (byte)(value >> 8),
                    (byte)value,
                };
                _stream.Write(bytes, 0, bytes.Length);
            }

            private static string ToHex(byte[] bytes)
            {
                var chars = new char[bytes.Length * 2];
                const string hex = "0123456789ABCDEF";
                for (int index = 0; index < bytes.Length; index++)
                {
                    chars[index * 2] = hex[bytes[index] >> 4];
                    chars[index * 2 + 1] = hex[bytes[index] & 0x0F];
                }
                return new string(chars);
            }
        }
    }

    public static class OpenXmlWorkbookSheetMappingInference
    {
        public static OpenXmlWorkbookSheetMappingResult Infer(
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after,
            WorkbookSheetMappingOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (after == null) throw new ArgumentNullException(nameof(after));

            OpenXmlWorkbookSheetFingerprintSet beforeEvidence = OpenXmlSheetFingerprintBuilder.Build(
                before,
                cancellationToken);
            OpenXmlWorkbookSheetFingerprintSet afterEvidence = OpenXmlSheetFingerprintBuilder.Build(
                after,
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            WorkbookSheetMappingResult mapping = WorkbookSheetMappingInference.Infer(
                beforeEvidence.Fingerprints,
                afterEvidence.Fingerprints,
                options);
            return new OpenXmlWorkbookSheetMappingResult(beforeEvidence, afterEvidence, mapping);
        }

        public static WorkbookSheetMappingResult InferMappings(
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after,
            WorkbookSheetMappingOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            Infer(before, after, options, cancellationToken).Mapping;
    }
}
