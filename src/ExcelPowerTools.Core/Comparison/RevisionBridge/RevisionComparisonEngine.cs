using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.RevisionBridge
{
    public sealed class RevisionComparisonSourceInspection
    {
        private readonly IReadOnlyList<RevisionComparisonSheetInspection> _sheets;

        internal RevisionComparisonSourceInspection(
            string path,
            string contentSha256,
            long byteLength,
            OpenXmlWorkbookSnapshot snapshot,
            IEnumerable<RevisionComparisonSheetInspection> sheets)
        {
            Path = path;
            FileName = System.IO.Path.GetFileName(path);
            ContentSha256 = contentSha256;
            ByteLength = byteLength;
            Snapshot = snapshot;
            _sheets = new ReadOnlyCollection<RevisionComparisonSheetInspection>(sheets.ToList());
        }

        public string Path { get; }
        public string FileName { get; }
        public string ContentSha256 { get; }
        public long ByteLength { get; }
        public IReadOnlyList<RevisionComparisonSheetInspection> Sheets => _sheets;
        internal OpenXmlWorkbookSnapshot Snapshot { get; }

        public RevisionComparisonSheetInspection FindSheet(string name) =>
            _sheets.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.Ordinal));
    }

    public sealed class RevisionComparisonSheetInspection
    {
        private readonly IReadOnlyList<string> _headers;

        internal RevisionComparisonSheetInspection(
            string name,
            int suggestedHeaderRow,
            string suggestedKeyHeader,
            IEnumerable<string> headers,
            int dataRowCount)
        {
            Name = name;
            SuggestedHeaderRow = suggestedHeaderRow;
            SuggestedKeyHeader = suggestedKeyHeader ?? "";
            _headers = new ReadOnlyCollection<string>(headers.ToList());
            DataRowCount = dataRowCount;
        }

        public string Name { get; }
        public int SuggestedHeaderRow { get; }
        public string SuggestedKeyHeader { get; }
        public IReadOnlyList<string> Headers => _headers;
        public int DataRowCount { get; }
    }

    public sealed class RevisionComparisonInspection
    {
        private readonly IReadOnlyList<RevisionComparisonSourceInspection> _sources;

        internal RevisionComparisonInspection(IEnumerable<RevisionComparisonSourceInspection> sources)
        {
            _sources = new ReadOnlyCollection<RevisionComparisonSourceInspection>(sources.ToList());
        }

        public IReadOnlyList<RevisionComparisonSourceInspection> Sources => _sources;
    }

    public sealed class RevisionComparisonPairSpec
    {
        public RevisionComparisonPairSpec(
            int beforeRevisionIndex,
            int afterRevisionIndex,
            string beforeSheetName,
            string afterSheetName,
            int beforeHeaderRow,
            int afterHeaderRow,
            IEnumerable<string> beforeKeyHeaders,
            IEnumerable<string> afterKeyHeaders)
        {
            BeforeRevisionIndex = beforeRevisionIndex;
            AfterRevisionIndex = afterRevisionIndex;
            BeforeSheetName = beforeSheetName ?? "";
            AfterSheetName = afterSheetName ?? "";
            BeforeHeaderRow = beforeHeaderRow;
            AfterHeaderRow = afterHeaderRow;
            BeforeKeyHeaders = CleanKeys(beforeKeyHeaders);
            AfterKeyHeaders = CleanKeys(afterKeyHeaders);
        }

        public int BeforeRevisionIndex { get; }
        public int AfterRevisionIndex { get; }
        public string BeforeSheetName { get; }
        public string AfterSheetName { get; }
        public int BeforeHeaderRow { get; }
        public int AfterHeaderRow { get; }
        public IReadOnlyList<string> BeforeKeyHeaders { get; }
        public IReadOnlyList<string> AfterKeyHeaders { get; }

        private static IReadOnlyList<string> CleanKeys(IEnumerable<string> keys)
        {
            var result = new List<string>();
            foreach (string key in keys ?? Array.Empty<string>())
            {
                string clean = (key ?? "").Trim();
                if (clean.Length == 0) continue;
                if (result.Contains(clean, StringComparer.OrdinalIgnoreCase))
                    throw new ArgumentException("A key column cannot be selected more than once.", nameof(keys));
                result.Add(clean);
            }
            return new ReadOnlyCollection<string>(result);
        }
    }

    public sealed class RevisionComparisonPlan
    {
        private readonly IReadOnlyList<int> _orderedSourceIndexes;
        private readonly IReadOnlyList<RevisionComparisonPairSpec> _pairs;

        public RevisionComparisonPlan(
            IEnumerable<int> orderedSourceIndexes,
            IEnumerable<RevisionComparisonPairSpec> pairs,
            int applyTargetRevisionIndex)
        {
            _orderedSourceIndexes = new ReadOnlyCollection<int>((orderedSourceIndexes ?? throw new ArgumentNullException(nameof(orderedSourceIndexes))).ToList());
            _pairs = new ReadOnlyCollection<RevisionComparisonPairSpec>((pairs ?? throw new ArgumentNullException(nameof(pairs))).ToList());
            ApplyTargetRevisionIndex = applyTargetRevisionIndex;
        }

        public IReadOnlyList<int> OrderedSourceIndexes => _orderedSourceIndexes;
        public IReadOnlyList<RevisionComparisonPairSpec> Pairs => _pairs;
        public int ApplyTargetRevisionIndex { get; }
    }

    public sealed class RevisionComparisonKeyPreview
    {
        internal RevisionComparisonKeyPreview(
            int pairIndex,
            string transition,
            string sheetPair,
            int beforeRows,
            int beforeBlankKeys,
            int beforeDuplicateRows,
            int afterRows,
            int afterBlankKeys,
            int afterDuplicateRows)
        {
            PairIndex = pairIndex;
            Transition = transition;
            SheetPair = sheetPair;
            BeforeRows = beforeRows;
            BeforeBlankKeys = beforeBlankKeys;
            BeforeDuplicateRows = beforeDuplicateRows;
            AfterRows = afterRows;
            AfterBlankKeys = afterBlankKeys;
            AfterDuplicateRows = afterDuplicateRows;
        }

        public int PairIndex { get; }
        public string Transition { get; }
        public string SheetPair { get; }
        public int BeforeRows { get; }
        public int BeforeBlankKeys { get; }
        public int BeforeDuplicateRows { get; }
        public int AfterRows { get; }
        public int AfterBlankKeys { get; }
        public int AfterDuplicateRows { get; }
    }

    public enum RevisionComparisonActionKind
    {
        CellChange,
        AddedRow,
    }

    public sealed class RevisionComparisonCellPayload
    {
        internal RevisionComparisonCellPayload(string header, int sourceColumn, CanonicalCellState state)
        {
            Header = header;
            SourceColumn = sourceColumn;
            State = state;
        }

        public string Header { get; }
        public int SourceColumn { get; }
        public CanonicalCellState State { get; }
    }

    public sealed class RevisionComparisonAction
    {
        private readonly IReadOnlyList<string> _beforeKeyHeaders;
        private readonly IReadOnlyList<string> _afterKeyHeaders;
        private readonly IReadOnlyList<string> _beforeKeyParts;
        private readonly IReadOnlyList<string> _afterKeyParts;
        private readonly IReadOnlyList<RevisionComparisonCellPayload> _rowCells;

        internal RevisionComparisonAction(
            string actionId,
            RevisionComparisonActionKind kind,
            int order,
            int beforeRevisionIndex,
            int afterRevisionIndex,
            string transition,
            string beforeSheetName,
            string afterSheetName,
            int beforeHeaderRow,
            int afterHeaderRow,
            IEnumerable<string> beforeKeyHeaders,
            IEnumerable<string> afterKeyHeaders,
            IEnumerable<string> beforeKeyParts,
            IEnumerable<string> afterKeyParts,
            string keyDisplay,
            int beforeRow,
            int afterRow,
            string beforeColumnHeader,
            string afterColumnHeader,
            int beforeColumn,
            int afterColumn,
            CanonicalCellState beforeState,
            CanonicalCellState afterState,
            IEnumerable<RevisionComparisonCellPayload> rowCells)
        {
            ActionId = actionId;
            Kind = kind;
            Order = order;
            BeforeRevisionIndex = beforeRevisionIndex;
            AfterRevisionIndex = afterRevisionIndex;
            Transition = transition;
            BeforeSheetName = beforeSheetName;
            AfterSheetName = afterSheetName;
            BeforeHeaderRow = beforeHeaderRow;
            AfterHeaderRow = afterHeaderRow;
            _beforeKeyHeaders = ReadOnly(beforeKeyHeaders);
            _afterKeyHeaders = ReadOnly(afterKeyHeaders);
            _beforeKeyParts = ReadOnly(beforeKeyParts);
            _afterKeyParts = ReadOnly(afterKeyParts);
            KeyDisplay = keyDisplay;
            BeforeRow = beforeRow;
            AfterRow = afterRow;
            BeforeColumnHeader = beforeColumnHeader ?? "";
            AfterColumnHeader = afterColumnHeader ?? "";
            BeforeColumn = beforeColumn;
            AfterColumn = afterColumn;
            BeforeState = beforeState;
            AfterState = afterState;
            _rowCells = new ReadOnlyCollection<RevisionComparisonCellPayload>((rowCells ?? Array.Empty<RevisionComparisonCellPayload>()).ToList());
        }

        public string ActionId { get; }
        public RevisionComparisonActionKind Kind { get; }
        public int Order { get; }
        public int BeforeRevisionIndex { get; }
        public int AfterRevisionIndex { get; }
        public string Transition { get; }
        public string BeforeSheetName { get; }
        public string AfterSheetName { get; }
        public int BeforeHeaderRow { get; }
        public int AfterHeaderRow { get; }
        public IReadOnlyList<string> BeforeKeyHeaders => _beforeKeyHeaders;
        public IReadOnlyList<string> AfterKeyHeaders => _afterKeyHeaders;
        public IReadOnlyList<string> BeforeKeyParts => _beforeKeyParts;
        public IReadOnlyList<string> AfterKeyParts => _afterKeyParts;
        public string KeyDisplay { get; }
        public int BeforeRow { get; }
        public int AfterRow { get; }
        public string BeforeColumnHeader { get; }
        public string AfterColumnHeader { get; }
        public int BeforeColumn { get; }
        public int AfterColumn { get; }
        public CanonicalCellState BeforeState { get; }
        public CanonicalCellState AfterState { get; }
        public IReadOnlyList<RevisionComparisonCellPayload> RowCells => _rowCells;

        private static IReadOnlyList<string> ReadOnly(IEnumerable<string> values) =>
            new ReadOnlyCollection<string>((values ?? Array.Empty<string>()).ToList());
    }

    public sealed class RevisionComparisonRemovedRow
    {
        internal RevisionComparisonRemovedRow(string transition, string sheet, string keyDisplay, int sourceRow)
        {
            Transition = transition;
            Sheet = sheet;
            KeyDisplay = keyDisplay;
            SourceRow = sourceRow;
        }

        public string Transition { get; }
        public string Sheet { get; }
        public string KeyDisplay { get; }
        public int SourceRow { get; }
    }

    public sealed class RevisionComparisonDuplicateKey
    {
        internal RevisionComparisonDuplicateKey(string transition, string side, string sheet, string keyDisplay, IEnumerable<int> rows)
        {
            Transition = transition;
            Side = side;
            Sheet = sheet;
            KeyDisplay = keyDisplay;
            Rows = new ReadOnlyCollection<int>(rows.ToList());
        }

        public string Transition { get; }
        public string Side { get; }
        public string Sheet { get; }
        public string KeyDisplay { get; }
        public IReadOnlyList<int> Rows { get; }
    }

    public sealed class RevisionComparisonResult
    {
        internal RevisionComparisonResult(
            RevisionComparisonInspection inspection,
            RevisionComparisonPlan plan,
            string planFingerprint,
            IEnumerable<RevisionComparisonKeyPreview> previews,
            IEnumerable<RevisionComparisonAction> actions,
            IEnumerable<RevisionComparisonRemovedRow> removedRows,
            IEnumerable<RevisionComparisonDuplicateKey> duplicateKeys)
        {
            Inspection = inspection;
            Plan = plan;
            PlanFingerprint = planFingerprint;
            KeyPreviews = new ReadOnlyCollection<RevisionComparisonKeyPreview>(previews.ToList());
            Actions = new ReadOnlyCollection<RevisionComparisonAction>(actions.ToList());
            RemovedRows = new ReadOnlyCollection<RevisionComparisonRemovedRow>(removedRows.ToList());
            DuplicateKeys = new ReadOnlyCollection<RevisionComparisonDuplicateKey>(duplicateKeys.ToList());
        }

        public RevisionComparisonInspection Inspection { get; }
        public RevisionComparisonPlan Plan { get; }
        public string PlanFingerprint { get; }
        public IReadOnlyList<RevisionComparisonKeyPreview> KeyPreviews { get; }
        public IReadOnlyList<RevisionComparisonAction> Actions { get; }
        public IReadOnlyList<RevisionComparisonRemovedRow> RemovedRows { get; }
        public IReadOnlyList<RevisionComparisonDuplicateKey> DuplicateKeys { get; }
        public int AddedRowCount => Actions.Count(item => item.Kind == RevisionComparisonActionKind.AddedRow);
        public int ChangedCellCount => Actions.Count(item => item.Kind == RevisionComparisonActionKind.CellChange);
    }

    public static class RevisionComparisonEngine
    {
        public const int MaximumRevisionCount = 20;
        public const long MaximumAggregateSourceBytes = 768L * 1024L * 1024L;

        public static RevisionComparisonInspection Inspect(
            IEnumerable<string> sourcePaths,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string[] paths = (sourcePaths ?? throw new ArgumentNullException(nameof(sourcePaths)))
                .Select(NormalizeInput)
                .ToArray();
            if (paths.Length < 2 || paths.Length > MaximumRevisionCount)
                throw new ArgumentException("Select between 2 and 20 workbook revisions.", nameof(sourcePaths));
            if (paths.Distinct(StringComparer.OrdinalIgnoreCase).Count() != paths.Length)
                throw new ArgumentException("Each revision path must be distinct.", nameof(sourcePaths));

            long aggregateBytes = 0;
            foreach (string path in paths)
            {
                aggregateBytes = checked(aggregateBytes + new FileInfo(path).Length);
                if (aggregateBytes > MaximumAggregateSourceBytes)
                    throw new InvalidDataException("The selected revisions exceed the 768 MiB aggregate input limit.");
            }

            var sources = new List<RevisionComparisonSourceInspection>(paths.Length);
            for (int index = 0; index < paths.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string extension = System.IO.Path.GetExtension(paths[index]);
                if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Revision comparison accepts .xlsx and .xlsm workbooks only: " + System.IO.Path.GetFileName(paths[index]));

                byte[] bytes = File.ReadAllBytes(paths[index]);
                string hash = StableId.ForBytes(bytes);
                OpenXmlWorkbookSnapshot snapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                    "revision-" + (index + 1).ToString(CultureInfo.InvariantCulture),
                    bytes,
                    cancellationToken);
                var sheets = snapshot.Sheets
                    .Where(item => item.Kind == OpenXmlPhysicalSheetKind.Worksheet)
                    .Select(InspectSheet)
                    .ToList();
                sources.Add(new RevisionComparisonSourceInspection(paths[index], hash, bytes.LongLength, snapshot, sheets));
            }
            return new RevisionComparisonInspection(sources);
        }

        public static IReadOnlyList<RevisionComparisonKeyPreview> PreviewKeys(
            RevisionComparisonInspection inspection,
            RevisionComparisonPlan plan,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            ValidatePlan(inspection, plan);
            var previews = new List<RevisionComparisonKeyPreview>();
            for (int pairIndex = 0; pairIndex < plan.Pairs.Count; pairIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PairContext context = BuildPairContext(inspection, plan, plan.Pairs[pairIndex]);
                previews.Add(CreatePreview(pairIndex, context));
            }
            return new ReadOnlyCollection<RevisionComparisonKeyPreview>(previews);
        }

        public static RevisionComparisonResult Compare(
            RevisionComparisonInspection inspection,
            RevisionComparisonPlan plan,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            ValidatePlan(inspection, plan);
            var previews = new List<RevisionComparisonKeyPreview>();
            var actions = new List<RevisionComparisonAction>();
            var removed = new List<RevisionComparisonRemovedRow>();
            var duplicates = new List<RevisionComparisonDuplicateKey>();
            int order = 0;

            for (int pairIndex = 0; pairIndex < plan.Pairs.Count; pairIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RevisionComparisonPairSpec spec = plan.Pairs[pairIndex];
                PairContext context = BuildPairContext(inspection, plan, spec);
                previews.Add(CreatePreview(pairIndex, context));
                CollectDuplicates(context, duplicates);

                foreach (KeyValuePair<string, RowObservation> item in context.Before.UniqueRows.OrderBy(item => item.Value.RowNumber))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!context.After.UniqueRows.TryGetValue(item.Key, out RowObservation afterRow))
                    {
                        if (context.After.AllKeyIds.Contains(item.Key)) continue;
                        removed.Add(new RevisionComparisonRemovedRow(
                            context.Transition,
                            spec.BeforeSheetName,
                            item.Value.KeyDisplay,
                            item.Value.RowNumber));
                        continue;
                    }
                    CompareMatchedRow(context, item.Value, afterRow, actions, ref order);
                }

                foreach (KeyValuePair<string, RowObservation> item in context.After.UniqueRows.OrderBy(item => item.Value.RowNumber))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (context.Before.AllKeyIds.Contains(item.Key)) continue;
                    RevisionComparisonAction action = CreateAddedRowAction(context, item.Value, order++);
                    actions.Add(action);
                }
            }

            string fingerprint = ComputePlanFingerprint(inspection, plan);
            return new RevisionComparisonResult(inspection, plan, fingerprint, previews, actions, removed, duplicates);
        }

        public static string ComputePlanFingerprint(RevisionComparisonInspection inspection, RevisionComparisonPlan plan)
        {
            ValidatePlan(inspection, plan);
            var parts = new List<string>();
            foreach (int sourceIndex in plan.OrderedSourceIndexes)
                parts.Add(inspection.Sources[sourceIndex].ContentSha256);
            parts.Add("target:" + plan.ApplyTargetRevisionIndex.ToString(CultureInfo.InvariantCulture));
            foreach (RevisionComparisonPairSpec pair in plan.Pairs)
            {
                parts.Add(string.Join("\u001F", new[]
                {
                    pair.BeforeRevisionIndex.ToString(CultureInfo.InvariantCulture),
                    pair.AfterRevisionIndex.ToString(CultureInfo.InvariantCulture),
                    pair.BeforeSheetName,
                    pair.AfterSheetName,
                    pair.BeforeHeaderRow.ToString(CultureInfo.InvariantCulture),
                    pair.AfterHeaderRow.ToString(CultureInfo.InvariantCulture),
                    string.Join("\u001E", pair.BeforeKeyHeaders),
                    string.Join("\u001E", pair.AfterKeyHeaders),
                }));
            }
            return StableId.Compute("revision-comparison-plan-v1", parts.ToArray());
        }

        internal static bool CellStatesEqual(CanonicalCellState left, CanonicalCellState right)
        {
            bool leftBlank = IsBlank(left);
            bool rightBlank = IsBlank(right);
            if (leftBlank || rightBlank) return leftBlank == rightBlank;
            if (left.HasFormula || right.HasFormula)
                return left.HasFormula && right.HasFormula &&
                       string.Equals(left.FormulaText, right.FormulaText, StringComparison.Ordinal);
            return left.ConstantValue.Equals(right.ConstantValue);
        }

        internal static bool IsBlank(CanonicalCellState state) =>
            state == null || (!state.HasFormula && state.ConstantValue.Kind == CanonicalValueKind.Blank);

        internal static string Display(CanonicalCellState state)
        {
            if (IsBlank(state)) return "";
            if (state.HasFormula) return "=" + state.FormulaText;
            return state.ConstantValue.ExactText;
        }

        internal static string ValueDisplay(CanonicalCellState state)
        {
            if (IsBlank(state)) return "";
            CanonicalValue value = state.HasFormula ? state.CachedResult : state.ConstantValue;
            return value == null || value.Kind == CanonicalValueKind.Blank ? "" : value.ExactText;
        }

        internal static string StateSignature(CanonicalCellState state)
        {
            if (IsBlank(state)) return "B";
            if (state.HasFormula)
                return "F:" + state.FormulaText;
            CanonicalValue value = state.ConstantValue;
            return ValueSignature(value);
        }

        private static bool ValuesEqual(CanonicalValue left, CanonicalValue right)
        {
            if (ReferenceEquals(left, right)) return true;
            return left != null && left.Equals(right);
        }

        private static string ValueSignature(CanonicalValue value) =>
            value == null
                ? "V:(missing)"
                : "V:" + value.Kind + ":" + value.DateSystem + ":" + value.ExactText;

        internal static string KeySignature(CanonicalCellState state)
        {
            if (IsBlank(state)) return "B";
            if (!state.HasFormula) return ValueSignature(state.ConstantValue);
            return state.CachedResult == null
                ? "F:" + state.FormulaText
                : ValueSignature(state.CachedResult);
        }

        internal static string KeyDisplay(CanonicalCellState state)
        {
            if (IsBlank(state)) return "";
            if (!state.HasFormula) return state.ConstantValue.ExactText;
            return state.CachedResult == null ? "=" + state.FormulaText : state.CachedResult.ExactText;
        }

        private static RevisionComparisonSheetInspection InspectSheet(OpenXmlWorksheetSnapshot sheet)
        {
            SheetGrid grid = SheetGrid.Create(sheet.Snapshot);
            int bestRow = 1;
            int bestScore = -1;
            int maxProbe = Math.Min(10, Math.Max(1, grid.MaxRow));
            for (int row = 1; row <= maxProbe; row++)
            {
                int nonblank = grid.Cells.Count(item => item.Key.Row == row && !IsBlank(item.Value));
                if (nonblank > bestScore)
                {
                    bestScore = nonblank;
                    bestRow = row;
                }
            }

            HeaderMap headers = HeaderMap.CreateForInspection(grid, bestRow);
            string suggestedKey = "";
            int dataRows = grid.ContentRows.Count(row => row > bestRow);
            foreach (HeaderColumn header in headers.UniqueColumns.Values.OrderBy(item => item.Column))
            {
                var seen = new HashSet<string>(StringComparer.Ordinal);
                bool unique = true;
                bool hasValue = false;
                foreach (int row in grid.ContentRows.Where(item => item > bestRow))
                {
                    CanonicalCellState cell = grid.Get(row, header.Column);
                    if (IsBlank(cell) || cell.HasFormula) { unique = false; break; }
                    hasValue = true;
                    if (!seen.Add(KeySignature(cell))) { unique = false; break; }
                }
                if (unique && hasValue)
                {
                    suggestedKey = header.Name;
                    break;
                }
            }
            if (suggestedKey.Length == 0 && headers.UniqueColumns.Count > 0)
                suggestedKey = headers.UniqueColumns.Values.OrderBy(item => item.Column).First().Name;

            return new RevisionComparisonSheetInspection(
                sheet.Name,
                bestRow,
                suggestedKey,
                headers.UniqueColumns.Values.OrderBy(item => item.Column).Select(item => item.Name),
                dataRows);
        }

        private static PairContext BuildPairContext(
            RevisionComparisonInspection inspection,
            RevisionComparisonPlan plan,
            RevisionComparisonPairSpec spec)
        {
            RevisionComparisonSourceInspection beforeSource = inspection.Sources[spec.BeforeRevisionIndex];
            RevisionComparisonSourceInspection afterSource = inspection.Sources[spec.AfterRevisionIndex];
            OpenXmlWorksheetSnapshot beforeSheet = FindSnapshot(beforeSource, spec.BeforeSheetName);
            OpenXmlWorksheetSnapshot afterSheet = FindSnapshot(afterSource, spec.AfterSheetName);
            EnsureReviewable(beforeSource, beforeSheet);
            EnsureReviewable(afterSource, afterSheet);
            SheetSide before = SheetSide.Create(beforeSheet, spec.BeforeHeaderRow, spec.BeforeKeyHeaders);
            SheetSide after = SheetSide.Create(afterSheet, spec.AfterHeaderRow, spec.AfterKeyHeaders);
            EnsureCompatibleHeaders(spec, before, after);
            int beforeOrdinal = plan.OrderedSourceIndexes.ToList().IndexOf(spec.BeforeRevisionIndex) + 1;
            int afterOrdinal = plan.OrderedSourceIndexes.ToList().IndexOf(spec.AfterRevisionIndex) + 1;
            return new PairContext(
                spec,
                "R" + beforeOrdinal.ToString(CultureInfo.InvariantCulture) + " (" + beforeSource.FileName + ") -> " +
                "R" + afterOrdinal.ToString(CultureInfo.InvariantCulture) + " (" + afterSource.FileName + ")",
                beforeSource.ContentSha256,
                afterSource.ContentSha256,
                before,
                after);
        }

        private static void EnsureCompatibleHeaders(
            RevisionComparisonPairSpec spec,
            SheetSide before,
            SheetSide after)
        {
            IEnumerable<string> beforeNonKeys = before.Headers.UniqueColumns.Keys
                .Where(header => !spec.BeforeKeyHeaders.Contains(header, StringComparer.OrdinalIgnoreCase));
            IEnumerable<string> afterNonKeys = after.Headers.UniqueColumns.Keys
                .Where(header => !spec.AfterKeyHeaders.Contains(header, StringComparer.OrdinalIgnoreCase));
            string[] beforeOnly = beforeNonKeys
                .Except(afterNonKeys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            string[] afterOnly = afterNonKeys
                .Except(beforeNonKeys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (beforeOnly.Length == 0 && afterOnly.Length == 0) return;

            string detail = (beforeOnly.Length == 0 ? "" : " Before only: " + string.Join(", ", beforeOnly) + ".") +
                            (afterOnly.Length == 0 ? "" : " After only: " + string.Join(", ", afterOnly) + ".");
            throw new InvalidOperationException(
                "The paired tables must have the same unique column headers before comparison can continue (" +
                spec.BeforeSheetName + " -> " + spec.AfterSheetName + ")." + detail);
        }

        private static void EnsureReviewable(
            RevisionComparisonSourceInspection source,
            OpenXmlWorksheetSnapshot sheet)
        {
            if (sheet.CanProjectExactStoredFormulaValueEvidence) return;
            OpenXmlCoverageMessage finding = sheet.CoverageMessages.FirstOrDefault(item => item.AffectsExactCellComparison);
            string detail = finding == null
                ? "The sheet does not provide complete stored formula/value evidence."
                : finding.Code + (string.IsNullOrEmpty(finding.CellReference) ? "" : " at " + finding.CellReference) +
                  ": " + finding.Message;
            throw new InvalidDataException(
                "Cannot compare sheet '" + sheet.Name + "' in " + source.FileName + " safely. " + detail);
        }

        private static OpenXmlWorksheetSnapshot FindSnapshot(RevisionComparisonSourceInspection source, string sheetName)
        {
            OpenXmlWorksheetSnapshot result = source.Snapshot.Sheets.FirstOrDefault(item =>
                item.Kind == OpenXmlPhysicalSheetKind.Worksheet &&
                string.Equals(item.Name, sheetName, StringComparison.Ordinal));
            if (result == null) throw new InvalidOperationException("Sheet '" + sheetName + "' was not found in " + source.FileName + ".");
            return result;
        }

        private static RevisionComparisonKeyPreview CreatePreview(int pairIndex, PairContext context) =>
            new RevisionComparisonKeyPreview(
                pairIndex,
                context.Transition,
                context.Spec.BeforeSheetName == context.Spec.AfterSheetName
                    ? context.Spec.BeforeSheetName
                    : context.Spec.BeforeSheetName + " -> " + context.Spec.AfterSheetName,
                context.Before.RowCount,
                context.Before.BlankKeyRows.Count,
                context.Before.DuplicateRows,
                context.After.RowCount,
                context.After.BlankKeyRows.Count,
                context.After.DuplicateRows);

        private static void CollectDuplicates(PairContext context, IList<RevisionComparisonDuplicateKey> output)
        {
            foreach (KeyValuePair<string, List<RowObservation>> item in context.Before.Duplicates
                         .OrderBy(item => item.Value[0].RowNumber).ThenBy(item => item.Key, StringComparer.Ordinal))
                output.Add(new RevisionComparisonDuplicateKey(
                    context.Transition, "Before", context.Spec.BeforeSheetName,
                    item.Value[0].KeyDisplay, item.Value.Select(row => row.RowNumber)));
            foreach (KeyValuePair<string, List<RowObservation>> item in context.After.Duplicates
                         .OrderBy(item => item.Value[0].RowNumber).ThenBy(item => item.Key, StringComparer.Ordinal))
                output.Add(new RevisionComparisonDuplicateKey(
                    context.Transition, "After", context.Spec.AfterSheetName,
                    item.Value[0].KeyDisplay, item.Value.Select(row => row.RowNumber)));
            foreach (RowObservation row in context.Before.BlankKeyRows)
                output.Add(new RevisionComparisonDuplicateKey(
                    context.Transition, "Before", context.Spec.BeforeSheetName,
                    "(blank key)", new[] { row.RowNumber }));
            foreach (RowObservation row in context.After.BlankKeyRows)
                output.Add(new RevisionComparisonDuplicateKey(
                    context.Transition, "After", context.Spec.AfterSheetName,
                    "(blank key)", new[] { row.RowNumber }));
        }

        private static void CompareMatchedRow(
            PairContext context,
            RowObservation beforeRow,
            RowObservation afterRow,
            IList<RevisionComparisonAction> actions,
            ref int order)
        {
            IEnumerable<string> commonHeaders = context.Before.Headers.UniqueColumns.Keys
                .Intersect(context.After.Headers.UniqueColumns.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => context.After.Headers.UniqueColumns[name].Column);

            foreach (string headerName in commonHeaders)
            {
                HeaderColumn beforeHeader = context.Before.Headers.UniqueColumns[headerName];
                HeaderColumn afterHeader = context.After.Headers.UniqueColumns[headerName];
                CanonicalCellState beforeState = context.Before.Grid.Get(beforeRow.RowNumber, beforeHeader.Column);
                CanonicalCellState afterState = context.After.Grid.Get(afterRow.RowNumber, afterHeader.Column);
                if (CellStatesEqual(beforeState, afterState)) continue;

                string actionId = StableId.Compute(
                    "revision-comparison-cell-action-v1",
                    context.Transition,
                    context.Spec.BeforeSheetName,
                    context.Spec.AfterSheetName,
                    context.BeforeContentSha256,
                    context.AfterContentSha256,
                    beforeRow.KeyId,
                    beforeHeader.Name,
                    StateSignature(beforeState),
                    StateSignature(afterState));
                actions.Add(new RevisionComparisonAction(
                    actionId,
                    RevisionComparisonActionKind.CellChange,
                    order++,
                    context.Spec.BeforeRevisionIndex,
                    context.Spec.AfterRevisionIndex,
                    context.Transition,
                    context.Spec.BeforeSheetName,
                    context.Spec.AfterSheetName,
                    context.Spec.BeforeHeaderRow,
                    context.Spec.AfterHeaderRow,
                    context.Spec.BeforeKeyHeaders,
                    context.Spec.AfterKeyHeaders,
                    beforeRow.KeyParts,
                    afterRow.KeyParts,
                    afterRow.KeyDisplay,
                    beforeRow.RowNumber,
                    afterRow.RowNumber,
                    beforeHeader.Name,
                    afterHeader.Name,
                    beforeHeader.Column,
                    afterHeader.Column,
                    beforeState,
                    afterState,
                    null));
            }
        }

        private static RevisionComparisonAction CreateAddedRowAction(PairContext context, RowObservation row, int order)
        {
            var payload = new List<RevisionComparisonCellPayload>();
            foreach (HeaderColumn header in context.After.Headers.UniqueColumns.Values.OrderBy(item => item.Column))
            {
                CanonicalCellState state = context.After.Grid.Get(row.RowNumber, header.Column);
                payload.Add(new RevisionComparisonCellPayload(header.Name, header.Column, state));
            }
            var actionIdentity = new List<string>
            {
                context.Transition,
                context.Spec.BeforeSheetName,
                context.Spec.AfterSheetName,
                context.BeforeContentSha256,
                context.AfterContentSha256,
                row.KeyId,
            };
            foreach (RevisionComparisonCellPayload cell in payload)
            {
                actionIdentity.Add(cell.Header);
                actionIdentity.Add(StateSignature(cell.State));
            }
            string actionId = StableId.Compute(
                "revision-comparison-added-row-action-v1",
                actionIdentity.ToArray());
            return new RevisionComparisonAction(
                actionId,
                RevisionComparisonActionKind.AddedRow,
                order,
                context.Spec.BeforeRevisionIndex,
                context.Spec.AfterRevisionIndex,
                context.Transition,
                context.Spec.BeforeSheetName,
                context.Spec.AfterSheetName,
                context.Spec.BeforeHeaderRow,
                context.Spec.AfterHeaderRow,
                context.Spec.BeforeKeyHeaders,
                context.Spec.AfterKeyHeaders,
                Array.Empty<string>(),
                row.KeyParts,
                row.KeyDisplay,
                0,
                row.RowNumber,
                "",
                "",
                0,
                0,
                null,
                null,
                payload);
        }

        private static void ValidatePlan(RevisionComparisonInspection inspection, RevisionComparisonPlan plan)
        {
            if (inspection == null) throw new ArgumentNullException(nameof(inspection));
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (plan.OrderedSourceIndexes.Count != inspection.Sources.Count || plan.OrderedSourceIndexes.Count < 2)
                throw new ArgumentException("The revision order must contain every inspected source exactly once.", nameof(plan));
            if (plan.OrderedSourceIndexes.Distinct().Count() != inspection.Sources.Count ||
                plan.OrderedSourceIndexes.Any(index => index < 0 || index >= inspection.Sources.Count))
                throw new ArgumentException("The revision order contains an invalid or repeated source.", nameof(plan));
            if (plan.ApplyTargetRevisionIndex < 0 || plan.ApplyTargetRevisionIndex >= inspection.Sources.Count)
                throw new ArgumentException("Choose a valid apply-target revision.", nameof(plan));
            if (plan.Pairs.Count == 0) throw new ArgumentException("At least one sheet pair is required.", nameof(plan));

            var adjacent = new HashSet<string>(StringComparer.Ordinal);
            var edgePositionById = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < plan.OrderedSourceIndexes.Count - 1; i++)
            {
                string edge = plan.OrderedSourceIndexes[i] + ":" + plan.OrderedSourceIndexes[i + 1];
                adjacent.Add(edge);
                edgePositionById.Add(edge, i);
            }
            var seenBefore = new HashSet<string>(StringComparer.Ordinal);
            var seenAfter = new HashSet<string>(StringComparer.Ordinal);
            var coveredEdges = new HashSet<string>(StringComparer.Ordinal);
            int lastEdgePosition = -1;
            foreach (RevisionComparisonPairSpec pair in plan.Pairs)
            {
                string edge = pair.BeforeRevisionIndex + ":" + pair.AfterRevisionIndex;
                if (!adjacent.Contains(edge))
                    throw new ArgumentException("Every sheet pair must compare adjacent revisions in the chosen order.", nameof(plan));
                int edgePosition = edgePositionById[edge];
                if (edgePosition < lastEdgePosition)
                    throw new ArgumentException("Sheet pairs must follow the confirmed revision order.", nameof(plan));
                lastEdgePosition = edgePosition;
                coveredEdges.Add(edge);
                if (!seenBefore.Add(edge + "\u001F" + pair.BeforeSheetName) ||
                    !seenAfter.Add(edge + "\u001F" + pair.AfterSheetName))
                    throw new ArgumentException("A sheet can participate in only one pair per revision transition.", nameof(plan));
                if (pair.BeforeHeaderRow < 1 || pair.AfterHeaderRow < 1)
                    throw new ArgumentException("Header rows must be positive Excel row numbers.", nameof(plan));
                if (pair.BeforeKeyHeaders.Count == 0 || pair.AfterKeyHeaders.Count == 0 ||
                    pair.BeforeKeyHeaders.Count != pair.AfterKeyHeaders.Count)
                    throw new ArgumentException("Each sheet pair requires the same positive number of explicit Before and After key columns.", nameof(plan));
                if (inspection.Sources[pair.BeforeRevisionIndex].FindSheet(pair.BeforeSheetName) == null ||
                    inspection.Sources[pair.AfterRevisionIndex].FindSheet(pair.AfterSheetName) == null)
                    throw new ArgumentException("A configured sheet name was not found in its revision.", nameof(plan));
            }
            if (coveredEdges.Count != adjacent.Count)
                throw new ArgumentException("Every adjacent revision transition requires at least one explicit sheet pair.", nameof(plan));
        }

        private static string NormalizeInput(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A workbook path is required.", nameof(path));
            string full = System.IO.Path.GetFullPath(path);
            if (!File.Exists(full)) throw new FileNotFoundException("Workbook revision not found.", full);
            return full;
        }

        private sealed class PairContext
        {
            internal PairContext(
                RevisionComparisonPairSpec spec,
                string transition,
                string beforeContentSha256,
                string afterContentSha256,
                SheetSide before,
                SheetSide after)
            {
                Spec = spec;
                Transition = transition;
                BeforeContentSha256 = beforeContentSha256;
                AfterContentSha256 = afterContentSha256;
                Before = before;
                After = after;
            }
            internal RevisionComparisonPairSpec Spec { get; }
            internal string Transition { get; }
            internal string BeforeContentSha256 { get; }
            internal string AfterContentSha256 { get; }
            internal SheetSide Before { get; }
            internal SheetSide After { get; }
        }

        private sealed class SheetSide
        {
            private SheetSide(
                SheetGrid grid,
                HeaderMap headers,
                Dictionary<string, RowObservation> uniqueRows,
                Dictionary<string, List<RowObservation>> duplicates,
                HashSet<string> allKeyIds,
                List<RowObservation> blankKeyRows,
                int rowCount)
            {
                Grid = grid;
                Headers = headers;
                UniqueRows = uniqueRows;
                Duplicates = duplicates;
                AllKeyIds = allKeyIds;
                BlankKeyRows = blankKeyRows;
                RowCount = rowCount;
            }

            internal SheetGrid Grid { get; }
            internal HeaderMap Headers { get; }
            internal Dictionary<string, RowObservation> UniqueRows { get; }
            internal Dictionary<string, List<RowObservation>> Duplicates { get; }
            internal HashSet<string> AllKeyIds { get; }
            internal List<RowObservation> BlankKeyRows { get; }
            internal int DuplicateRows => Duplicates.Values.Sum(rows => rows.Count);
            internal int RowCount { get; }

            internal static SheetSide Create(OpenXmlWorksheetSnapshot sheet, int headerRow, IReadOnlyList<string> keyHeaders)
            {
                SheetGrid grid = SheetGrid.Create(sheet.Snapshot);
                for (int column = 1; column <= grid.MaxColumn; column++)
                {
                    if (Display(grid.Get(headerRow, column)).Trim().Length > 0) continue;
                    int populatedRow = grid.ContentRows.FirstOrDefault(row =>
                        row > headerRow && !IsBlank(grid.Get(row, column)));
                    if (populatedRow > 0)
                        throw new InvalidOperationException(
                            "Column " + column.ToString(CultureInfo.InvariantCulture) + " on sheet '" + sheet.Name +
                            "' has data at row " + populatedRow.ToString(CultureInfo.InvariantCulture) +
                            " but its header cell on row " + headerRow.ToString(CultureInfo.InvariantCulture) +
                            " is blank. Every populated comparison column needs a unique header.");
                }
                HeaderMap headers = HeaderMap.Create(grid, headerRow, sheet.Name);
                var keyColumns = new List<int>();
                foreach (string keyHeader in keyHeaders)
                {
                    if (!headers.UniqueColumns.TryGetValue(keyHeader, out HeaderColumn column))
                        throw new InvalidOperationException("Key column '" + keyHeader + "' was not found uniquely on sheet '" + sheet.Name + "' row " + headerRow + ".");
                    keyColumns.Add(column.Column);
                }

                var all = new Dictionary<string, List<RowObservation>>(StringComparer.Ordinal);
                var blank = new List<RowObservation>();
                int rowCount = 0;
                foreach (int row in grid.ContentRows.Where(item => item > headerRow))
                {
                    rowCount++;
                    var keyParts = new List<string>(keyColumns.Count);
                    var displayParts = new List<string>(keyColumns.Count);
                    bool hasBlank = false;
                    foreach (int column in keyColumns)
                    {
                        CanonicalCellState cell = grid.Get(row, column);
                        if (IsBlank(cell)) hasBlank = true;
                        if (cell?.HasFormula == true)
                            throw new InvalidOperationException(
                                "Formula cells cannot be used as row keys in the first working version (sheet '" +
                                sheet.Name + "', row " + row.ToString(CultureInfo.InvariantCulture) + ").");
                        keyParts.Add(KeySignature(cell));
                        displayParts.Add(KeyDisplay(cell));
                    }
                    string keyId = hasBlank ? "" : StableId.Compute("revision-comparison-row-key-v1", keyParts.ToArray());
                    var observation = new RowObservation(row, keyId, string.Join(" | ", displayParts), keyParts);
                    if (hasBlank)
                    {
                        blank.Add(observation);
                        continue;
                    }
                    if (!all.TryGetValue(keyId, out List<RowObservation> rows))
                        all.Add(keyId, rows = new List<RowObservation>());
                    rows.Add(observation);
                }

                var unique = all.Where(item => item.Value.Count == 1)
                    .ToDictionary(item => item.Key, item => item.Value[0], StringComparer.Ordinal);
                var duplicates = all.Where(item => item.Value.Count > 1)
                    .ToDictionary(item => item.Key, item => item.Value, StringComparer.Ordinal);
                return new SheetSide(
                    grid,
                    headers,
                    unique,
                    duplicates,
                    new HashSet<string>(all.Keys, StringComparer.Ordinal),
                    blank,
                    rowCount);
            }
        }

        private sealed class RowObservation
        {
            internal RowObservation(int rowNumber, string keyId, string keyDisplay, IEnumerable<string> keyParts)
            {
                RowNumber = rowNumber;
                KeyId = keyId;
                KeyDisplay = keyDisplay;
                KeyParts = new ReadOnlyCollection<string>(keyParts.ToList());
            }
            internal int RowNumber { get; }
            internal string KeyId { get; }
            internal string KeyDisplay { get; }
            internal IReadOnlyList<string> KeyParts { get; }
        }

        private sealed class HeaderMap
        {
            private HeaderMap(Dictionary<string, HeaderColumn> uniqueColumns) { UniqueColumns = uniqueColumns; }
            internal Dictionary<string, HeaderColumn> UniqueColumns { get; }

            internal static HeaderMap CreateForInspection(SheetGrid grid, int headerRow) =>
                CreateCore(grid, headerRow, "", rejectDuplicates: false);

            internal static HeaderMap Create(SheetGrid grid, int headerRow, string sheetName) =>
                CreateCore(grid, headerRow, sheetName, rejectDuplicates: true);

            private static HeaderMap CreateCore(
                SheetGrid grid,
                int headerRow,
                string sheetName,
                bool rejectDuplicates)
            {
                var occurrences = new Dictionary<string, List<HeaderColumn>>(StringComparer.OrdinalIgnoreCase);
                for (int column = 1; column <= grid.MaxColumn; column++)
                {
                    string name = Display(grid.Get(headerRow, column)).Trim();
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (!occurrences.TryGetValue(name, out List<HeaderColumn> list))
                        occurrences.Add(name, list = new List<HeaderColumn>());
                    list.Add(new HeaderColumn(name, column));
                }
                var duplicate = occurrences.Where(item => item.Value.Count > 1)
                    .OrderBy(item => item.Value[0].Column)
                    .FirstOrDefault();
                if (rejectDuplicates && duplicate.Value != null)
                    throw new InvalidOperationException("Header '" + duplicate.Key + "' appears more than once on sheet '" + sheetName + "' row " + headerRow + ". Headers used for comparison must be unique.");
                return new HeaderMap(occurrences
                    .Where(item => item.Value.Count == 1)
                    .ToDictionary(item => item.Key, item => item.Value[0], StringComparer.OrdinalIgnoreCase));
            }
        }

        private sealed class HeaderColumn
        {
            internal HeaderColumn(string name, int column) { Name = name; Column = column; }
            internal string Name { get; }
            internal int Column { get; }
        }

        private sealed class SheetGrid
        {
            private readonly Dictionary<CellCoordinate, CanonicalCellState> _cells;

            private SheetGrid(
                Dictionary<CellCoordinate, CanonicalCellState> cells,
                IEnumerable<int> contentRows,
                int maxRow,
                int maxColumn)
            {
                _cells = cells;
                ContentRows = new ReadOnlyCollection<int>(contentRows.OrderBy(item => item).ToList());
                MaxRow = maxRow;
                MaxColumn = maxColumn;
            }

            internal IEnumerable<KeyValuePair<CellCoordinate, CanonicalCellState>> Cells => _cells;
            internal int MaxRow { get; }
            internal int MaxColumn { get; }
            internal IReadOnlyList<int> ContentRows { get; }

            internal CanonicalCellState Get(int row, int column)
            {
                _cells.TryGetValue(new CellCoordinate(row, column), out CanonicalCellState state);
                return state;
            }

            internal static SheetGrid Create(RevisionSnapshot snapshot)
            {
                var cells = new Dictionary<CellCoordinate, CanonicalCellState>();
                var contentRows = new HashSet<int>();
                int maxRow = 0;
                int maxColumn = 0;
                foreach (CanonicalCellState cell in snapshot.Cells.Values)
                {
                    if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(cell.SubjectKey, out int row, out int column)) continue;
                    cells[new CellCoordinate(row, column)] = cell;
                    if (!IsBlank(cell))
                    {
                        contentRows.Add(row);
                        maxRow = Math.Max(maxRow, row);
                        maxColumn = Math.Max(maxColumn, column);
                    }
                }
                return new SheetGrid(cells, contentRows, maxRow, maxColumn);
            }
        }

        private struct CellCoordinate : IEquatable<CellCoordinate>
        {
            internal CellCoordinate(int row, int column) { Row = row; Column = column; }
            internal int Row { get; }
            internal int Column { get; }
            public bool Equals(CellCoordinate other) => Row == other.Row && Column == other.Column;
            public override bool Equals(object obj) => obj is CellCoordinate other && Equals(other);
            public override int GetHashCode() { unchecked { return (Row * 397) ^ Column; } }
        }
    }
}
