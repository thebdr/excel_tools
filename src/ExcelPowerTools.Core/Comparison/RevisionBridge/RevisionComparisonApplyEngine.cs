using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.RevisionBridge
{
    public sealed class RevisionComparisonApplyResult
    {
        internal RevisionComparisonApplyResult(
            string outputPath,
            int selectedActionCount,
            int changedCellCount,
            int addedRowCount,
            int alreadyCurrentCount)
        {
            OutputPath = outputPath;
            SelectedActionCount = selectedActionCount;
            ChangedCellCount = changedCellCount;
            AddedRowCount = addedRowCount;
            AlreadyCurrentCount = alreadyCurrentCount;
        }

        public string OutputPath { get; }
        public int SelectedActionCount { get; }
        public int ChangedCellCount { get; }
        public int AddedRowCount { get; }
        public int AlreadyCurrentCount { get; }
    }

    /// <summary>
    /// Rebuilds the comparison from unchanged sources, reanchors every selected action by explicit row keys, applies
    /// into a private copy of the first revision, reopens that copy and proves every final written cell before the
    /// destination becomes visible. Source revisions and the review workbook are never written.
    /// </summary>
    public static class RevisionComparisonApplyEngine
    {
        public static RevisionComparisonApplyResult ApplyReview(string reviewPath, string outputPath)
        {
            RevisionComparisonReviewSelection selection = RevisionComparisonReviewFile.ReadSelection(reviewPath);
            if (selection.SelectedActionIds.Count == 0)
                throw new InvalidOperationException(
                    "No actions are selected. Type Yes in the Apply column on Added or Changed, save and close the review, then retry.");

            string[] sourcePaths = selection.Sources.OrderBy(item => item.Index).Select(item => item.Path).ToArray();
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(sourcePaths);
            ValidateSources(selection, inspection);
            RevisionComparisonResult comparison = RevisionComparisonEngine.Compare(inspection, selection.Plan);
            if (!string.Equals(comparison.PlanFingerprint, selection.PlanFingerprint, StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The review plan no longer matches the re-read sources. Create a fresh comparison review.");

            var actionsById = comparison.Actions.ToDictionary(item => item.ActionId, StringComparer.Ordinal);
            var selected = new List<RevisionComparisonAction>(selection.SelectedActionIds.Count);
            foreach (string actionId in selection.SelectedActionIds)
            {
                if (!actionsById.TryGetValue(actionId, out RevisionComparisonAction action))
                    throw new InvalidDataException(
                        "A selected action is missing from the freshly rebuilt comparison. Create a fresh review.");
                selected.Add(action);
            }
            selected = selected.OrderBy(item => item.Order).ToList();

            if (selection.Plan.ApplyTargetRevisionIndex != selection.Plan.OrderedSourceIndexes[0])
                throw new NotSupportedException(
                    "This first working apply path requires the first revision in the confirmed order as its output base.");

            RevisionComparisonReviewSource targetSource = selection.Sources[
                selection.Plan.ApplyTargetRevisionIndex];
            string destination = NormalizeOutput(outputPath, targetSource.Path, reviewPath, selection.Sources);
            byte[] targetBytes = RecheckSourcesAndReadTarget(selection, targetSource.Index);
            byte[] targetVba = ReadVbaProject(targetBytes);
            string temporary = destination + ".writing-" + Guid.NewGuid().ToString("N");

            int changed = 0;
            int added = 0;
            int alreadyCurrent = 0;
            try
            {
                File.WriteAllBytes(temporary, targetBytes);
                var expected = new Dictionary<ExpectedCoordinate, CanonicalCellState>();
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(temporary, true))
                {
                    var mutable = new MutableWorkbook(document, inspection.Sources[targetSource.Index].Snapshot);
                    foreach (RevisionComparisonAction action in selected)
                    {
                        TargetContext context = ResolveTargetContext(selection.Plan, action);
                        MutableSheet sheet = mutable.GetSheet(context.SheetName);
                        switch (action.Kind)
                        {
                            case RevisionComparisonActionKind.CellChange:
                                ApplyCellChange(
                                    inspection,
                                    action,
                                    context,
                                    sheet,
                                    expected,
                                    ref changed,
                                    ref alreadyCurrent);
                                break;

                            case RevisionComparisonActionKind.AddedRow:
                                ApplyAddedRow(
                                    inspection,
                                    action,
                                    context,
                                    sheet,
                                    expected,
                                    ref added,
                                    ref alreadyCurrent);
                                break;

                            default:
                                throw new InvalidDataException("The selected comparison action kind is unsupported.");
                        }
                    }

                    if (mutable.WroteCell)
                    {
                        WorkbookPart workbookPart = document.WorkbookPart;
                        if (workbookPart.CalculationChainPart != null)
                            workbookPart.DeletePart(workbookPart.CalculationChainPart);
                        CalculationProperties calculation = workbookPart.Workbook.CalculationProperties;
                        if (calculation == null)
                            workbookPart.Workbook.CalculationProperties = calculation = new CalculationProperties();
                        calculation.CalculationId = 0U;
                        calculation.FullCalculationOnLoad = true;
                        calculation.ForceFullCalculation = true;
                    }
                    mutable.Save();
                }

                byte[] outputBytes = File.ReadAllBytes(temporary);
                VerifyOutput(outputBytes, expected, targetVba);
                RecheckSources(selection, "immediately before output publication");
                Publish(temporary, destination);
                return new RevisionComparisonApplyResult(
                    destination,
                    selected.Count,
                    changed,
                    added,
                    alreadyCurrent);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }

        private static void ApplyCellChange(
            RevisionComparisonInspection inspection,
            RevisionComparisonAction action,
            TargetContext context,
            MutableSheet sheet,
            IDictionary<ExpectedCoordinate, CanonicalCellState> expected,
            ref int changed,
            ref int alreadyCurrent)
        {
            IReadOnlyList<int> rows = sheet.FindRowsByKey(
                context.HeaderRow,
                context.KeyHeaders,
                action.BeforeKeyParts);
            if (rows.Count != 1)
                throw Conflict(action, rows.Count == 0
                    ? "the keyed row is missing in the output base"
                    : "the keyed row is duplicated in the output base");

            int targetRow = rows[0];
            int targetColumn = sheet.FindUniqueHeader(context.HeaderRow, action.BeforeColumnHeader);
            EnsureFormulaCanBeApplied(inspection, action, action.AfterState, action.AfterRow, action.AfterColumn);
            if (!sheet.WasWritten(targetRow, targetColumn))
                EnsureTargetFormulaIsNormal(inspection, context, targetRow, targetColumn, action);

            CanonicalCellState expectedBefore = TranslateState(
                action.BeforeState,
                action.BeforeRow,
                action.BeforeColumn,
                targetRow,
                targetColumn,
                action);
            CanonicalCellState desiredAfter = TranslateState(
                action.AfterState,
                action.AfterRow,
                action.AfterColumn,
                targetRow,
                targetColumn,
                action);
            CanonicalCellState current = sheet.Get(targetRow, targetColumn);
            if (RevisionComparisonEngine.CellStatesEqual(current, desiredAfter))
            {
                expected[new ExpectedCoordinate(sheet.Name, targetRow, targetColumn)] = desiredAfter;
                alreadyCurrent++;
                return;
            }
            if (!RevisionComparisonEngine.CellStatesEqual(current, expectedBefore))
                throw Conflict(action, "the current cell is neither the compared Before value nor the selected After value");

            sheet.Set(targetRow, targetColumn, desiredAfter);
            expected[new ExpectedCoordinate(sheet.Name, targetRow, targetColumn)] = desiredAfter;
            changed++;
        }

        private static void ApplyAddedRow(
            RevisionComparisonInspection inspection,
            RevisionComparisonAction action,
            TargetContext context,
            MutableSheet sheet,
            IDictionary<ExpectedCoordinate, CanonicalCellState> expected,
            ref int added,
            ref int alreadyCurrent)
        {
            IReadOnlyList<int> existing = sheet.FindRowsByKey(
                context.HeaderRow,
                context.KeyHeaders,
                action.AfterKeyParts);
            if (existing.Count > 1) throw Conflict(action, "the added key is duplicated in the output base");

            var mapped = new List<MappedPayload>(action.RowCells.Count);
            foreach (RevisionComparisonCellPayload payload in action.RowCells)
            {
                string targetHeader = payload.Header;
                for (int keyIndex = 0; keyIndex < action.AfterKeyHeaders.Count; keyIndex++)
                {
                    if (!string.Equals(
                            action.AfterKeyHeaders[keyIndex],
                            payload.Header,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (context.KeyHeaders.Count != action.AfterKeyHeaders.Count)
                        throw Conflict(action, "the key-column mapping changed across revision transitions");
                    targetHeader = context.KeyHeaders[keyIndex];
                    break;
                }
                int column = sheet.FindUniqueHeader(context.HeaderRow, targetHeader);
                EnsureFormulaCanBeApplied(
                    inspection,
                    action,
                    payload.State,
                    action.AfterRow,
                    payload.SourceColumn);
                mapped.Add(new MappedPayload(payload, column));
            }

            if (existing.Count == 1)
            {
                int row = existing[0];
                foreach (MappedPayload item in mapped)
                {
                    CanonicalCellState desired = TranslateState(
                        item.Payload.State,
                        action.AfterRow,
                        item.Payload.SourceColumn,
                        row,
                        item.TargetColumn,
                        action);
                    if (!RevisionComparisonEngine.CellStatesEqual(sheet.Get(row, item.TargetColumn), desired))
                        throw Conflict(action, "the added key already exists with different row content");
                    expected[new ExpectedCoordinate(sheet.Name, row, item.TargetColumn)] = desired;
                }
                alreadyCurrent++;
                return;
            }

            int newRow = sheet.NextAppendRow(context.HeaderRow);
            var desiredStates = new List<KeyValuePair<int, CanonicalCellState>>(mapped.Count);
            foreach (MappedPayload item in mapped)
            {
                CanonicalCellState desired = TranslateState(
                    item.Payload.State,
                    action.AfterRow,
                    item.Payload.SourceColumn,
                    newRow,
                    item.TargetColumn,
                    action);
                desiredStates.Add(new KeyValuePair<int, CanonicalCellState>(item.TargetColumn, desired));
            }
            foreach (KeyValuePair<int, CanonicalCellState> item in desiredStates)
            {
                if (!RevisionComparisonEngine.IsBlank(item.Value))
                    sheet.Set(newRow, item.Key, item.Value);
                expected[new ExpectedCoordinate(sheet.Name, newRow, item.Key)] = item.Value;
            }

            IReadOnlyList<int> inserted = sheet.FindRowsByKey(
                context.HeaderRow,
                context.KeyHeaders,
                action.AfterKeyParts);
            if (inserted.Count != 1 || inserted[0] != newRow)
                throw Conflict(action, "the inserted row did not reanchor to its selected key");
            added++;
        }

        private static CanonicalCellState TranslateState(
            CanonicalCellState state,
            int sourceRow,
            int sourceColumn,
            int targetRow,
            int targetColumn,
            RevisionComparisonAction action)
        {
            string targetReference = CellReference(targetRow, targetColumn);
            if (RevisionComparisonEngine.IsBlank(state)) return CanonicalCellState.Blank(targetReference);
            if (!state.HasFormula)
                return CanonicalCellState.Constant(targetReference, state.ConstantValue, state.DisplayText);

            string formula = state.FormulaText;
            string sourceReference = CellReference(sourceRow, sourceColumn);
            if (!string.Equals(sourceReference, targetReference, StringComparison.Ordinal) &&
                !OpenXmlSharedFormulaTranslator.TryTranslate(
                    formula,
                    sourceReference,
                    targetReference,
                    out formula,
                    out string failureReason))
                throw Conflict(action, "the formula cannot be translated to its keyed output row: " + failureReason);
            return CanonicalCellState.Formula(
                targetReference,
                formula,
                state.CachedResult,
                state.DisplayText);
        }

        private static void EnsureFormulaCanBeApplied(
            RevisionComparisonInspection inspection,
            RevisionComparisonAction action,
            CanonicalCellState state,
            int sourceRow,
            int sourceColumn)
        {
            if (state?.HasFormula != true) return;
            RevisionComparisonSourceInspection source = inspection.Sources[action.AfterRevisionIndex];
            OpenXmlWorksheetSnapshot sheet = FindSnapshot(source, action.AfterSheetName);
            string reference = CellReference(sourceRow, sourceColumn);
            OpenXmlFormulaObservation observation = sheet.FormulaObservations.FirstOrDefault(item =>
                string.Equals(item.CellReference, reference, StringComparison.Ordinal));
            if (observation == null || observation.Form != OpenXmlFormulaForm.Normal)
                throw Conflict(
                    action,
                    "formula apply currently supports ordinary single-cell formulas only (" + reference + ")");
        }

        private static void EnsureTargetFormulaIsNormal(
            RevisionComparisonInspection inspection,
            TargetContext context,
            int row,
            int column,
            RevisionComparisonAction action)
        {
            RevisionComparisonSourceInspection target = inspection.Sources[context.RevisionIndex];
            OpenXmlWorksheetSnapshot sheet = FindSnapshot(target, context.SheetName);
            string reference = CellReference(row, column);
            OpenXmlFormulaObservation observation = sheet.FormulaObservations.FirstOrDefault(item =>
                string.Equals(item.CellReference, reference, StringComparison.Ordinal));
            if (observation != null && observation.Form != OpenXmlFormulaForm.Normal)
                throw Conflict(action, "the target cell belongs to a grouped or special formula and cannot be edited safely");
        }

        private static TargetContext ResolveTargetContext(
            RevisionComparisonPlan plan,
            RevisionComparisonAction action)
        {
            int revisionIndex = action.BeforeRevisionIndex;
            string sheetName = action.BeforeSheetName;
            int headerRow = action.BeforeHeaderRow;
            IReadOnlyList<string> keyHeaders = action.BeforeKeyHeaders;
            while (revisionIndex != plan.ApplyTargetRevisionIndex)
            {
                int position = plan.OrderedSourceIndexes.ToList().IndexOf(revisionIndex);
                if (position <= 0)
                    throw Conflict(action, "the action cannot be traced back to the selected output base");
                int predecessor = plan.OrderedSourceIndexes[position - 1];
                List<RevisionComparisonPairSpec> matches = plan.Pairs.Where(pair =>
                    pair.BeforeRevisionIndex == predecessor &&
                    pair.AfterRevisionIndex == revisionIndex &&
                    string.Equals(pair.AfterSheetName, sheetName, StringComparison.Ordinal)).ToList();
                if (matches.Count != 1)
                    throw Conflict(action, "the sheet track cannot be traced uniquely to the output base");
                RevisionComparisonPairSpec pair = matches[0];
                revisionIndex = predecessor;
                sheetName = pair.BeforeSheetName;
                headerRow = pair.BeforeHeaderRow;
                keyHeaders = pair.BeforeKeyHeaders;
            }
            return new TargetContext(revisionIndex, sheetName, headerRow, keyHeaders);
        }

        private static void ValidateSources(
            RevisionComparisonReviewSelection selection,
            RevisionComparisonInspection inspection)
        {
            if (selection.Sources.Count != inspection.Sources.Count)
                throw new InvalidDataException("The review source count does not match the re-read comparison.");
            for (int index = 0; index < selection.Sources.Count; index++)
            {
                RevisionComparisonReviewSource expected = selection.Sources[index];
                RevisionComparisonSourceInspection actual = inspection.Sources[index];
                if (expected.Index != index ||
                    !string.Equals(expected.Path, actual.Path, StringComparison.OrdinalIgnoreCase) ||
                    expected.ByteLength != actual.ByteLength ||
                    !string.Equals(expected.ContentSha256, actual.ContentSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(
                        "A source revision changed after the review was created: " + expected.Path);
            }
        }

        private static byte[] RecheckSourcesAndReadTarget(
            RevisionComparisonReviewSelection selection,
            int targetIndex)
        {
            byte[] target = null;
            foreach (RevisionComparisonReviewSource source in selection.Sources)
            {
                byte[] bytes = File.ReadAllBytes(source.Path);
                if (bytes.LongLength != source.ByteLength ||
                    !string.Equals(StableId.ForBytes(bytes), source.ContentSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(
                        "A source revision changed immediately before output creation: " + source.Path);
                if (source.Index == targetIndex) target = bytes;
            }
            return target ?? throw new InvalidDataException("The selected output base is missing from the review.");
        }

        private static void RecheckSources(
            RevisionComparisonReviewSelection selection,
            string stage)
        {
            foreach (RevisionComparisonReviewSource source in selection.Sources)
            {
                byte[] bytes = File.ReadAllBytes(source.Path);
                if (bytes.LongLength != source.ByteLength ||
                    !string.Equals(StableId.ForBytes(bytes), source.ContentSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(
                        "A source revision changed " + stage + ": " + source.Path);
            }
        }

        private static void VerifyOutput(
            byte[] outputBytes,
            IReadOnlyDictionary<ExpectedCoordinate, CanonicalCellState> expected,
            byte[] expectedVba)
        {
            OpenXmlWorkbookSnapshot snapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes("applied-output", outputBytes);
            foreach (KeyValuePair<ExpectedCoordinate, CanonicalCellState> item in expected)
            {
                OpenXmlWorksheetSnapshot sheet = snapshot.Sheets.FirstOrDefault(candidate =>
                    candidate.Kind == OpenXmlPhysicalSheetKind.Worksheet &&
                    string.Equals(candidate.Name, item.Key.SheetName, StringComparison.Ordinal));
                if (sheet == null)
                    throw new InvalidDataException("The applied output is missing sheet '" + item.Key.SheetName + "'.");
                string reference = CellReference(item.Key.Row, item.Key.Column);
                CanonicalCellState actual = sheet.Snapshot.TryGetCell(reference, out CanonicalCellState found)
                    ? found
                    : CanonicalCellState.Blank(reference);
                if (!RevisionComparisonEngine.CellStatesEqual(actual, item.Value))
                    throw new InvalidDataException(
                        "The applied output failed verification at " + item.Key.SheetName + "!" + reference + ".");
            }

            byte[] actualVba = ReadVbaProject(outputBytes);
            if (!ByteArraysEqual(expectedVba, actualVba))
                throw new InvalidDataException("The applied output did not preserve its VBA project byte-for-byte.");
        }

        private static byte[] ReadVbaProject(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes, false))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                VbaProjectPart part = document.WorkbookPart?.VbaProjectPart;
                if (part == null) return null;
                using (Stream source = part.GetStream(FileMode.Open, FileAccess.Read))
                using (var destination = new MemoryStream())
                {
                    source.CopyTo(destination);
                    return destination.ToArray();
                }
            }
        }

        private static bool ByteArraysEqual(byte[] left, byte[] right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (left == null || right == null || left.Length != right.Length) return false;
            for (int index = 0; index < left.Length; index++)
                if (left[index] != right[index]) return false;
            return true;
        }

        private static OpenXmlWorksheetSnapshot FindSnapshot(
            RevisionComparisonSourceInspection source,
            string sheetName)
        {
            return source.Snapshot.Sheets.FirstOrDefault(item =>
                item.Kind == OpenXmlPhysicalSheetKind.Worksheet &&
                string.Equals(item.Name, sheetName, StringComparison.Ordinal)) ??
                throw new InvalidDataException("Comparison sheet not found: " + sheetName);
        }

        private static InvalidDataException Conflict(RevisionComparisonAction action, string reason) =>
            new InvalidDataException(
                "Cannot apply " + action.Transition + ", " + action.AfterSheetName + ", key " +
                action.KeyDisplay + (action.AfterColumnHeader.Length == 0 ? "" : ", column " + action.AfterColumnHeader) +
                ": " + reason + ". No output was published.");

        private static string NormalizeOutput(
            string outputPath,
            string targetPath,
            string reviewPath,
            IEnumerable<RevisionComparisonReviewSource> sources)
        {
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
            string full = System.IO.Path.GetFullPath(outputPath);
            string targetExtension = System.IO.Path.GetExtension(targetPath);
            if (!string.Equals(System.IO.Path.GetExtension(full), targetExtension, StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("The output must keep the base revision's " + targetExtension + " extension.");
            if (string.Equals(full, System.IO.Path.GetFullPath(reviewPath), StringComparison.OrdinalIgnoreCase) ||
                sources.Any(source => string.Equals(full, source.Path, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The output cannot overwrite the review or any source revision.");
            if (File.Exists(full))
                throw new InvalidOperationException(
                    "The output must be a new workbook. Choose a filename that does not already exist.");
            string directory = System.IO.Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException("The output folder does not exist: " + directory);
            return full;
        }

        private static void Publish(string temporary, string destination)
        {
            if (File.Exists(destination))
                throw new IOException("The output path became occupied before publication; no existing file was replaced.");
            File.Move(temporary, destination);
        }

        private static string CellReference(int row, int column) =>
            ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

        private static string ColumnName(int column)
        {
            if (column < 1 || column > 16384) throw new ArgumentOutOfRangeException(nameof(column));
            var result = new StringBuilder();
            while (column > 0)
            {
                column--;
                result.Insert(0, (char)('A' + column % 26));
                column /= 26;
            }
            return result.ToString();
        }

        private sealed class MutableWorkbook
        {
            private readonly SpreadsheetDocument _document;
            private readonly Dictionary<string, MutableSheet> _sheets;

            internal MutableWorkbook(SpreadsheetDocument document, OpenXmlWorkbookSnapshot snapshot)
            {
                _document = document ?? throw new ArgumentNullException(nameof(document));
                WorkbookPart workbookPart = document.WorkbookPart ??
                    throw new InvalidDataException("The output base has no workbook part.");
                _sheets = new Dictionary<string, MutableSheet>(StringComparer.Ordinal);
                foreach (Sheet sheet in workbookPart.Workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>())
                {
                    if (sheet.Id == null || sheet.Name == null) continue;
                    WorksheetPart part = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    OpenXmlWorksheetSnapshot source = snapshot.Sheets.FirstOrDefault(item =>
                        item.Kind == OpenXmlPhysicalSheetKind.Worksheet &&
                        string.Equals(item.Name, sheet.Name.Value, StringComparison.Ordinal));
                    if (part != null && source != null)
                        _sheets.Add(sheet.Name.Value, new MutableSheet(this, sheet.Name.Value, part, source.Snapshot));
                }
            }

            internal bool WroteCell { get; private set; }

            internal MutableSheet GetSheet(string name)
            {
                if (!_sheets.TryGetValue(name, out MutableSheet sheet))
                    throw new InvalidDataException("The output base is missing mapped sheet '" + name + "'.");
                return sheet;
            }

            internal void NoteWrite() => WroteCell = true;

            internal void Save()
            {
                foreach (MutableSheet sheet in _sheets.Values) sheet.Save();
                _document.WorkbookPart.Workbook.Save();
            }
        }

        private sealed class MutableSheet
        {
            private readonly MutableWorkbook _owner;
            private readonly WorksheetPart _part;
            private readonly Dictionary<Coordinate, CanonicalCellState> _cells =
                new Dictionary<Coordinate, CanonicalCellState>();
            private readonly HashSet<Coordinate> _written = new HashSet<Coordinate>();
            private int _maximumRow;

            internal MutableSheet(
                MutableWorkbook owner,
                string name,
                WorksheetPart part,
                RevisionSnapshot snapshot)
            {
                _owner = owner;
                Name = name;
                _part = part;
                foreach (CanonicalCellState cell in snapshot.Cells.Values)
                {
                    if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(cell.SubjectKey, out int row, out int column))
                        continue;
                    _cells[new Coordinate(row, column)] = cell;
                    if (!RevisionComparisonEngine.IsBlank(cell))
                        _maximumRow = Math.Max(_maximumRow, row);
                }
            }

            internal string Name { get; }

            internal CanonicalCellState Get(int row, int column)
            {
                _cells.TryGetValue(new Coordinate(row, column), out CanonicalCellState state);
                return state;
            }

            internal bool WasWritten(int row, int column) => _written.Contains(new Coordinate(row, column));

            internal int FindUniqueHeader(int headerRow, string header)
            {
                string wanted = (header ?? "").Trim();
                List<int> matches = _cells.Where(item =>
                        item.Key.Row == headerRow &&
                        string.Equals(RevisionComparisonEngine.Display(item.Value).Trim(), wanted, StringComparison.OrdinalIgnoreCase))
                    .Select(item => item.Key.Column)
                    .OrderBy(item => item)
                    .ToList();
                if (matches.Count != 1)
                    throw new InvalidDataException(
                        "Header '" + header + "' was not found uniquely on output sheet '" + Name + "'.");
                return matches[0];
            }

            internal IReadOnlyList<int> FindRowsByKey(
                int headerRow,
                IReadOnlyList<string> keyHeaders,
                IReadOnlyList<string> keyParts)
            {
                if (keyHeaders.Count != keyParts.Count)
                    throw new InvalidDataException("The selected action has incomplete key evidence.");
                int[] columns = keyHeaders.Select(header => FindUniqueHeader(headerRow, header)).ToArray();
                var candidateRows = new SortedSet<int>(_cells.Keys.Where(item => item.Row > headerRow).Select(item => item.Row));
                var result = new List<int>();
                foreach (int row in candidateRows)
                {
                    bool matches = true;
                    for (int index = 0; index < columns.Length; index++)
                    {
                        CanonicalCellState cell = Get(row, columns[index]);
                        if (RevisionComparisonEngine.IsBlank(cell) || cell.HasFormula ||
                            !string.Equals(RevisionComparisonEngine.KeySignature(cell), keyParts[index], StringComparison.Ordinal))
                        {
                            matches = false;
                            break;
                        }
                    }
                    if (matches) result.Add(row);
                }
                return new ReadOnlyCollection<int>(result);
            }

            internal int NextAppendRow(int headerRow)
            {
                int next = Math.Max(_maximumRow + 1, headerRow + 1);
                if (next > 1048576) throw new InvalidDataException("The output sheet has no row available for an insertion.");
                return next;
            }

            internal void Set(int row, int column, CanonicalCellState state)
            {
                Coordinate coordinate = new Coordinate(row, column);
                _cells[coordinate] = state;
                _written.Add(coordinate);
                _maximumRow = Math.Max(_maximumRow, row);
                WriteCell(row, column, state);
            }

            internal void Save() => _part.Worksheet.Save();

            private void WriteCell(int rowNumber, int columnNumber, CanonicalCellState state)
            {
                SheetData sheetData = _part.Worksheet.GetFirstChild<SheetData>();
                if (sheetData == null)
                {
                    sheetData = new SheetData();
                    _part.Worksheet.Append(sheetData);
                }
                Row row = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value == (uint)rowNumber);
                if (row == null)
                {
                    row = new Row { RowIndex = (uint)rowNumber };
                    Row nextRow = sheetData.Elements<Row>().FirstOrDefault(item => item.RowIndex?.Value > (uint)rowNumber);
                    if (nextRow == null) sheetData.Append(row);
                    else sheetData.InsertBefore(row, nextRow);
                }

                string reference = CellReference(rowNumber, columnNumber);
                Cell cell = row.Elements<Cell>().FirstOrDefault(item =>
                    string.Equals(item.CellReference?.Value, reference, StringComparison.Ordinal));
                if (cell == null)
                {
                    cell = new Cell { CellReference = reference };
                    Cell next = row.Elements<Cell>().FirstOrDefault(item =>
                        ColumnIndex(item.CellReference?.Value) > columnNumber);
                    if (next == null) row.Append(cell);
                    else row.InsertBefore(cell, next);
                }

                cell.RemoveAllChildren<CellFormula>();
                cell.RemoveAllChildren<CellValue>();
                cell.RemoveAllChildren<InlineString>();
                cell.DataType = null;
                _owner.NoteWrite();
                if (RevisionComparisonEngine.IsBlank(state)) return;
                if (state.HasFormula)
                {
                    cell.CellFormula = new CellFormula(state.FormulaText);
                    WriteScalar(cell, state.CachedResult, true);
                }
                else
                {
                    WriteScalar(cell, state.ConstantValue, false);
                }
            }

            private static void WriteScalar(Cell cell, CanonicalValue value, bool formulaCache)
            {
                if (value == null || value.Kind == CanonicalValueKind.Blank) return;
                switch (value.Kind)
                {
                    case CanonicalValueKind.Text:
                        if (formulaCache)
                        {
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(value.ExactText);
                        }
                        else
                        {
                            var text = new DocumentFormat.OpenXml.Spreadsheet.Text(value.ExactText);
                            if (text.Text.Length != text.Text.Trim().Length)
                                text.Space = SpaceProcessingModeValues.Preserve;
                            cell.DataType = CellValues.InlineString;
                            cell.InlineString = new InlineString(text);
                        }
                        break;

                    case CanonicalValueKind.Number:
                    case CanonicalValueKind.DateSerial:
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(value.ExactText);
                        break;

                    case CanonicalValueKind.Boolean:
                        cell.DataType = CellValues.Boolean;
                        cell.CellValue = new CellValue(
                            string.Equals(value.ExactText, "true", StringComparison.OrdinalIgnoreCase) ? "1" : "0");
                        break;

                    case CanonicalValueKind.Error:
                        cell.DataType = CellValues.Error;
                        cell.CellValue = new CellValue(value.ExactText);
                        break;

                    case CanonicalValueKind.IsoDateTime:
                        cell.DataType = CellValues.Date;
                        cell.CellValue = new CellValue(value.ExactText);
                        break;

                    default:
                        throw new InvalidDataException("The selected comparison value kind is unsupported.");
                }
            }
        }

        private sealed class TargetContext
        {
            internal TargetContext(int revisionIndex, string sheetName, int headerRow, IReadOnlyList<string> keyHeaders)
            {
                RevisionIndex = revisionIndex;
                SheetName = sheetName;
                HeaderRow = headerRow;
                KeyHeaders = keyHeaders;
            }
            internal int RevisionIndex { get; }
            internal string SheetName { get; }
            internal int HeaderRow { get; }
            internal IReadOnlyList<string> KeyHeaders { get; }
        }

        private sealed class MappedPayload
        {
            internal MappedPayload(RevisionComparisonCellPayload payload, int targetColumn)
            {
                Payload = payload;
                TargetColumn = targetColumn;
            }
            internal RevisionComparisonCellPayload Payload { get; }
            internal int TargetColumn { get; }
        }

        private struct Coordinate : IEquatable<Coordinate>
        {
            internal Coordinate(int row, int column) { Row = row; Column = column; }
            internal int Row { get; }
            internal int Column { get; }
            public bool Equals(Coordinate other) => Row == other.Row && Column == other.Column;
            public override bool Equals(object obj) => obj is Coordinate other && Equals(other);
            public override int GetHashCode() { unchecked { return (Row * 397) ^ Column; } }
        }

        private struct ExpectedCoordinate : IEquatable<ExpectedCoordinate>
        {
            internal ExpectedCoordinate(string sheetName, int row, int column)
            {
                SheetName = sheetName;
                Row = row;
                Column = column;
            }
            internal string SheetName { get; }
            internal int Row { get; }
            internal int Column { get; }
            public bool Equals(ExpectedCoordinate other) =>
                Row == other.Row && Column == other.Column &&
                string.Equals(SheetName, other.SheetName, StringComparison.Ordinal);
            public override bool Equals(object obj) => obj is ExpectedCoordinate other && Equals(other);
            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = StringComparer.Ordinal.GetHashCode(SheetName ?? "");
                    hash = (hash * 397) ^ Row;
                    return (hash * 397) ^ Column;
                }
            }
        }

        private static int ColumnIndex(string reference)
        {
            if (string.IsNullOrEmpty(reference)) return 0;
            int result = 0;
            int index = 0;
            while (index < reference.Length && reference[index] >= 'A' && reference[index] <= 'Z')
            {
                result = checked(result * 26 + reference[index] - 'A' + 1);
                index++;
            }
            return result;
        }
    }
}
