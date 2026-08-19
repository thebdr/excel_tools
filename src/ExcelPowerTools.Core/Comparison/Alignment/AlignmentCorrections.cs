using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    public enum AlignmentCorrectionKind
    {
        ConfirmSheetPair = 0,
        ConfirmHeaderRow = 1,
        ConfirmKeyColumns = 2,
        ClearSheetPair = 3,
        ClearHeaderRow = 4,
        ClearKeyColumns = 5,
        ConfirmSheetUnmatched = 6,
        ClearSheetUnmatched = 7,
    }

    public enum AlignmentHeaderConfirmationDecision
    {
        HeaderRow = 0,
        NoHeader = 1,
    }

    /// <summary>
    /// One table-relative axis index together with the one-based physical worksheet coordinate observed when
    /// it was confirmed. The project engine requires both values to match before applying the confirmation.
    /// </summary>
    public sealed class AlignmentGridIndexReference : IEquatable<AlignmentGridIndexReference>
    {
        public AlignmentGridIndexReference(int tableIndex, int physicalIndex)
        {
            if (tableIndex < 0) throw new ArgumentOutOfRangeException(nameof(tableIndex));
            if (physicalIndex <= 0) throw new ArgumentOutOfRangeException(nameof(physicalIndex));
            TableIndex = tableIndex;
            PhysicalIndex = physicalIndex;
        }

        public int TableIndex { get; }
        public int PhysicalIndex { get; }

        public bool Equals(AlignmentGridIndexReference other) =>
            other != null && TableIndex == other.TableIndex && PhysicalIndex == other.PhysicalIndex;

        public override bool Equals(object obj) => Equals(obj as AlignmentGridIndexReference);

        public override int GetHashCode()
        {
            unchecked { return (TableIndex * 397) ^ PhysicalIndex; }
        }

        internal string StableComponent =>
            TableIndex.ToString(CultureInfo.InvariantCulture) + ":" +
            PhysicalIndex.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The exact active-correction slot affected by a request. Unused identity components are exposed as empty
    /// strings so the scope can be serialized without guessing from a nullable confirmation payload.
    /// </summary>
    public sealed class AlignmentCorrectionScope
    {
        private AlignmentCorrectionScope(
            string parentRevisionId,
            string childRevisionId,
            string revisionId,
            string beforeSheetKey,
            string afterSheetKey,
            string physicalSheetKey)
        {
            ParentRevisionId = parentRevisionId ?? "";
            ChildRevisionId = childRevisionId ?? "";
            RevisionId = revisionId ?? "";
            BeforeSheetKey = beforeSheetKey ?? "";
            AfterSheetKey = afterSheetKey ?? "";
            PhysicalSheetKey = physicalSheetKey ?? "";
        }

        public string ParentRevisionId { get; }
        public string ChildRevisionId { get; }
        public string RevisionId { get; }
        public string BeforeSheetKey { get; }
        public string AfterSheetKey { get; }
        public string PhysicalSheetKey { get; }

        internal static AlignmentCorrectionScope SheetPair(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey) =>
            new AlignmentCorrectionScope(
                Required(parentRevisionId, nameof(parentRevisionId)),
                Required(childRevisionId, nameof(childRevisionId)),
                "",
                Required(beforeSheetKey, nameof(beforeSheetKey)),
                "",
                "");

        internal static AlignmentCorrectionScope HeaderRow(
            string revisionId,
            string physicalSheetKey) =>
            new AlignmentCorrectionScope(
                "",
                "",
                Required(revisionId, nameof(revisionId)),
                "",
                "",
                Required(physicalSheetKey, nameof(physicalSheetKey)));

        internal static AlignmentCorrectionScope KeyColumns(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey) =>
            new AlignmentCorrectionScope(
                Required(parentRevisionId, nameof(parentRevisionId)),
                Required(childRevisionId, nameof(childRevisionId)),
                "",
                Required(beforeSheetKey, nameof(beforeSheetKey)),
                Required(afterSheetKey, nameof(afterSheetKey)),
                "");

        private static string Required(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact alignment identity is required.", name);
            return value;
        }
    }

    public sealed class AlignmentSheetPairConfirmation
    {
        internal AlignmentSheetPairConfirmation(
            string confirmationId,
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey)
        {
            ConfirmationId = Require(confirmationId, nameof(confirmationId));
            ParentRevisionId = Require(parentRevisionId, nameof(parentRevisionId));
            ChildRevisionId = Require(childRevisionId, nameof(childRevisionId));
            BeforeSheetKey = Require(beforeSheetKey, nameof(beforeSheetKey));
            AfterSheetKey = Require(afterSheetKey, nameof(afterSheetKey));
        }

        public string ConfirmationId { get; }
        public string ParentRevisionId { get; }
        public string ChildRevisionId { get; }
        public string BeforeSheetKey { get; }
        public string AfterSheetKey { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact alignment identity is required.", name);
            return value;
        }
    }

    /// <summary>
    /// Explicit evidence that one exact source-sheet occurrence has no successor on an adjacent revision edge.
    /// It rejects a rename pairing without suppressing the resulting deleted-sheet and added-sheet history.
    /// </summary>
    public sealed class AlignmentSheetUnmatchedConfirmation
    {
        internal AlignmentSheetUnmatchedConfirmation(
            string confirmationId,
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey)
        {
            ConfirmationId = Require(confirmationId, nameof(confirmationId));
            ParentRevisionId = Require(parentRevisionId, nameof(parentRevisionId));
            ChildRevisionId = Require(childRevisionId, nameof(childRevisionId));
            BeforeSheetKey = Require(beforeSheetKey, nameof(beforeSheetKey));
        }

        public string ConfirmationId { get; }
        public string ParentRevisionId { get; }
        public string ChildRevisionId { get; }
        public string BeforeSheetKey { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact alignment identity is required.", name);
            return value;
        }
    }

    public sealed class AlignmentHeaderRowConfirmation
    {
        internal AlignmentHeaderRowConfirmation(
            string confirmationId,
            string revisionId,
            string physicalSheetKey,
            AlignmentHeaderConfirmationDecision decision,
            AlignmentGridIndexReference row,
            AlignmentGridIndexReference firstDataRow)
        {
            ConfirmationId = Require(confirmationId, nameof(confirmationId));
            RevisionId = Require(revisionId, nameof(revisionId));
            PhysicalSheetKey = Require(physicalSheetKey, nameof(physicalSheetKey));
            Decision = decision;
            if (decision == AlignmentHeaderConfirmationDecision.HeaderRow && row == null)
                throw new ArgumentNullException(nameof(row));
            if (decision == AlignmentHeaderConfirmationDecision.NoHeader &&
                (row != null || firstDataRow != null))
                throw new ArgumentException("A no-header confirmation cannot carry header/data row indices.");
            if (firstDataRow != null &&
                (firstDataRow.TableIndex <= row.TableIndex ||
                 firstDataRow.PhysicalIndex <= row.PhysicalIndex))
                throw new ArgumentException("The first data row must follow the confirmed header row.");
            Row = row;
            FirstDataRow = firstDataRow;
        }

        public string ConfirmationId { get; }
        public string RevisionId { get; }
        public string PhysicalSheetKey { get; }
        public AlignmentHeaderConfirmationDecision Decision { get; }
        public AlignmentGridIndexReference Row { get; }
        /// <summary>
        /// Optional first keyed-data row. Null means the row immediately after the header. It may point one row
        /// beyond the materialized grid to represent a table with headers/subheaders and no stored data rows.
        /// </summary>
        public AlignmentGridIndexReference FirstDataRow { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact alignment identity is required.", name);
            return value;
        }
    }

    public sealed class AlignmentKeyColumnsConfirmation
    {
        internal AlignmentKeyColumnsConfirmation(
            string confirmationId,
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey,
            IEnumerable<AlignmentGridIndexReference> beforeColumns,
            IEnumerable<AlignmentGridIndexReference> afterColumns)
        {
            ConfirmationId = Require(confirmationId, nameof(confirmationId));
            ParentRevisionId = Require(parentRevisionId, nameof(parentRevisionId));
            ChildRevisionId = Require(childRevisionId, nameof(childRevisionId));
            BeforeSheetKey = Require(beforeSheetKey, nameof(beforeSheetKey));
            AfterSheetKey = Require(afterSheetKey, nameof(afterSheetKey));
            BeforeColumns = FreezeColumns(beforeColumns, nameof(beforeColumns));
            AfterColumns = FreezeColumns(afterColumns, nameof(afterColumns));
            if (BeforeColumns.Count != AfterColumns.Count)
                throw new ArgumentException("Confirmed Before and After key columns must have equal length.");
        }

        public string ConfirmationId { get; }
        public string ParentRevisionId { get; }
        public string ChildRevisionId { get; }
        public string BeforeSheetKey { get; }
        public string AfterSheetKey { get; }
        public IReadOnlyList<AlignmentGridIndexReference> BeforeColumns { get; }
        public IReadOnlyList<AlignmentGridIndexReference> AfterColumns { get; }

        private static IReadOnlyList<AlignmentGridIndexReference> FreezeColumns(
            IEnumerable<AlignmentGridIndexReference> columns,
            string name)
        {
            if (columns == null) throw new ArgumentNullException(name);
            var copy = new List<AlignmentGridIndexReference>(columns);
            if (copy.Count == 0 || copy.Count > 4)
                throw new ArgumentException("A confirmed key requires between one and four column pairs.", name);
            if (copy.Any(item => item == null))
                throw new ArgumentException("Confirmed key columns cannot contain null entries.", name);
            if (copy.Select(item => item.TableIndex).Distinct().Count() != copy.Count ||
                copy.Select(item => item.PhysicalIndex).Distinct().Count() != copy.Count)
                throw new ArgumentException("Confirmed key columns cannot repeat a table or physical index.", name);
            return new ReadOnlyCollection<AlignmentGridIndexReference>(copy);
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact alignment identity is required.", name);
            return value;
        }
    }

    /// <summary>Immutable request. Applying it creates a new generation and one append-only audit event.</summary>
    public sealed class AlignmentCorrectionRequest
    {
        private AlignmentCorrectionRequest(
            AlignmentCorrectionKind kind,
            string requestId,
            AlignmentCorrectionScope scope,
            AlignmentSheetPairConfirmation sheetPair,
            AlignmentSheetUnmatchedConfirmation sheetUnmatched,
            AlignmentHeaderRowConfirmation headerRow,
            AlignmentKeyColumnsConfirmation keyColumns,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note)
        {
            Kind = kind;
            RequestId = Required(requestId, nameof(requestId));
            Scope = scope ?? throw new ArgumentNullException(nameof(scope));
            SheetPair = sheetPair;
            SheetUnmatched = sheetUnmatched;
            HeaderRow = headerRow;
            KeyColumns = keyColumns;
            Reviewer = reviewer ?? ReviewerSnapshot.Empty;
            CreatedUtc = createdUtc;
            Note = note ?? "";
        }

        public AlignmentCorrectionKind Kind { get; }
        public string RequestId { get; }
        public AlignmentCorrectionScope Scope { get; }
        public AlignmentSheetPairConfirmation SheetPair { get; }
        public AlignmentSheetUnmatchedConfirmation SheetUnmatched { get; }
        public AlignmentHeaderRowConfirmation HeaderRow { get; }
        public AlignmentKeyColumnsConfirmation KeyColumns { get; }
        public ReviewerSnapshot Reviewer { get; }
        public DateTimeOffset CreatedUtc { get; }
        public string Note { get; }

        public static AlignmentCorrectionRequest ConfirmSheetPair(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            string id = StableId.Compute(
                "alignment-sheet-pair-confirmation-v1",
                Required(parentRevisionId, nameof(parentRevisionId)),
                Required(childRevisionId, nameof(childRevisionId)),
                Required(beforeSheetKey, nameof(beforeSheetKey)),
                Required(afterSheetKey, nameof(afterSheetKey)));
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ConfirmSheetPair,
                id,
                AlignmentCorrectionScope.SheetPair(parentRevisionId, childRevisionId, beforeSheetKey),
                new AlignmentSheetPairConfirmation(
                    id, parentRevisionId, childRevisionId, beforeSheetKey, afterSheetKey),
                null,
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static AlignmentCorrectionRequest ConfirmSheetUnmatched(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            string id = StableId.Compute(
                "alignment-sheet-unmatched-confirmation-v1",
                Required(parentRevisionId, nameof(parentRevisionId)),
                Required(childRevisionId, nameof(childRevisionId)),
                Required(beforeSheetKey, nameof(beforeSheetKey)));
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ConfirmSheetUnmatched,
                id,
                AlignmentCorrectionScope.SheetPair(parentRevisionId, childRevisionId, beforeSheetKey),
                null,
                new AlignmentSheetUnmatchedConfirmation(
                    id, parentRevisionId, childRevisionId, beforeSheetKey),
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static AlignmentCorrectionRequest ConfirmHeaderRow(
            string revisionId,
            string physicalSheetKey,
            int tableRowIndex,
            int oneBasedPhysicalRow,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var row = new AlignmentGridIndexReference(tableRowIndex, oneBasedPhysicalRow);
            string id = StableId.Compute(
                "alignment-header-row-confirmation-v1",
                Required(revisionId, nameof(revisionId)),
                Required(physicalSheetKey, nameof(physicalSheetKey)),
                row.StableComponent);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ConfirmHeaderRow,
                id,
                AlignmentCorrectionScope.HeaderRow(revisionId, physicalSheetKey),
                null,
                null,
                new AlignmentHeaderRowConfirmation(
                    id,
                    revisionId,
                    physicalSheetKey,
                    AlignmentHeaderConfirmationDecision.HeaderRow,
                    row,
                    null),
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static AlignmentCorrectionRequest ConfirmNoHeader(
            string revisionId,
            string physicalSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            string id = StableId.Compute(
                "alignment-no-header-confirmation-v1",
                Required(revisionId, nameof(revisionId)),
                Required(physicalSheetKey, nameof(physicalSheetKey)));
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ConfirmHeaderRow,
                id,
                AlignmentCorrectionScope.HeaderRow(revisionId, physicalSheetKey),
                null,
                null,
                new AlignmentHeaderRowConfirmation(
                    id,
                    revisionId,
                    physicalSheetKey,
                    AlignmentHeaderConfirmationDecision.NoHeader,
                    null,
                    null),
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static AlignmentCorrectionRequest ConfirmHeaderRow(
            string revisionId,
            string physicalSheetKey,
            int tableRowIndex,
            int oneBasedPhysicalRow,
            int firstDataTableRowIndex,
            int oneBasedFirstDataPhysicalRow,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var row = new AlignmentGridIndexReference(tableRowIndex, oneBasedPhysicalRow);
            var firstData = new AlignmentGridIndexReference(
                firstDataTableRowIndex,
                oneBasedFirstDataPhysicalRow);
            if (firstData.TableIndex <= row.TableIndex || firstData.PhysicalIndex <= row.PhysicalIndex)
                throw new ArgumentException("The first data row must follow the confirmed header row.");
            string id = StableId.Compute(
                "alignment-header-row-confirmation-v1",
                Required(revisionId, nameof(revisionId)),
                Required(physicalSheetKey, nameof(physicalSheetKey)),
                row.StableComponent,
                firstData.StableComponent);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ConfirmHeaderRow,
                id,
                AlignmentCorrectionScope.HeaderRow(revisionId, physicalSheetKey),
                null,
                null,
                new AlignmentHeaderRowConfirmation(
                    id,
                    revisionId,
                    physicalSheetKey,
                    AlignmentHeaderConfirmationDecision.HeaderRow,
                    row,
                    firstData),
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static AlignmentCorrectionRequest ConfirmKeyColumns(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey,
            IEnumerable<AlignmentGridIndexReference> beforeColumns,
            IEnumerable<AlignmentGridIndexReference> afterColumns,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var before = new List<AlignmentGridIndexReference>(
                beforeColumns ?? throw new ArgumentNullException(nameof(beforeColumns)));
            var after = new List<AlignmentGridIndexReference>(
                afterColumns ?? throw new ArgumentNullException(nameof(afterColumns)));
            // Constructor performs the authoritative width/distinctness checks before the ID is accepted.
            var provisional = new AlignmentKeyColumnsConfirmation(
                "PENDING",
                parentRevisionId,
                childRevisionId,
                beforeSheetKey,
                afterSheetKey,
                before,
                after);
            var components = new List<string>
            {
                provisional.ParentRevisionId,
                provisional.ChildRevisionId,
                provisional.BeforeSheetKey,
                provisional.AfterSheetKey,
            };
            components.AddRange(provisional.BeforeColumns.Select(item => "B:" + item.StableComponent));
            components.AddRange(provisional.AfterColumns.Select(item => "A:" + item.StableComponent));
            string id = StableId.Compute("alignment-key-columns-confirmation-v1", components.ToArray());
            var confirmation = new AlignmentKeyColumnsConfirmation(
                id,
                parentRevisionId,
                childRevisionId,
                beforeSheetKey,
                afterSheetKey,
                before,
                after);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ConfirmKeyColumns,
                id,
                AlignmentCorrectionScope.KeyColumns(
                    parentRevisionId,
                    childRevisionId,
                    beforeSheetKey,
                    afterSheetKey),
                null,
                null,
                null,
                confirmation,
                reviewer,
                createdUtc,
                note);
        }

        /// <summary>Clears the active source-sheet pairing and returns that exact source scope to auto mapping.</summary>
        public static AlignmentCorrectionRequest ClearSheetPair(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var scope = AlignmentCorrectionScope.SheetPair(
                parentRevisionId,
                childRevisionId,
                beforeSheetKey);
            string id = StableId.Compute(
                "alignment-sheet-pair-clear-v1",
                scope.ParentRevisionId,
                scope.ChildRevisionId,
                scope.BeforeSheetKey);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ClearSheetPair,
                id,
                scope,
                null,
                null,
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        /// <summary>Clears the explicit deleted/new choice and returns that exact source scope to auto mapping.</summary>
        public static AlignmentCorrectionRequest ClearSheetUnmatched(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var scope = AlignmentCorrectionScope.SheetPair(
                parentRevisionId,
                childRevisionId,
                beforeSheetKey);
            string id = StableId.Compute(
                "alignment-sheet-unmatched-clear-v1",
                scope.ParentRevisionId,
                scope.ChildRevisionId,
                scope.BeforeSheetKey);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ClearSheetUnmatched,
                id,
                scope,
                null,
                null,
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        /// <summary>Clears the active header/no-header choice and returns that exact sheet to auto detection.</summary>
        public static AlignmentCorrectionRequest ClearHeaderRow(
            string revisionId,
            string physicalSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var scope = AlignmentCorrectionScope.HeaderRow(revisionId, physicalSheetKey);
            string id = StableId.Compute(
                "alignment-header-row-clear-v1",
                scope.RevisionId,
                scope.PhysicalSheetKey);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ClearHeaderRow,
                id,
                scope,
                null,
                null,
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        /// <summary>Clears the active key choice and returns that exact mapped sheet pair to auto inference.</summary>
        public static AlignmentCorrectionRequest ClearKeyColumns(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            var scope = AlignmentCorrectionScope.KeyColumns(
                parentRevisionId,
                childRevisionId,
                beforeSheetKey,
                afterSheetKey);
            string id = StableId.Compute(
                "alignment-key-columns-clear-v1",
                scope.ParentRevisionId,
                scope.ChildRevisionId,
                scope.BeforeSheetKey,
                scope.AfterSheetKey);
            return new AlignmentCorrectionRequest(
                AlignmentCorrectionKind.ClearKeyColumns,
                id,
                scope,
                null,
                null,
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        internal string ConfirmationId
        {
            get
            {
                switch (Kind)
                {
                    case AlignmentCorrectionKind.ConfirmSheetPair: return SheetPair.ConfirmationId;
                    case AlignmentCorrectionKind.ConfirmSheetUnmatched: return SheetUnmatched.ConfirmationId;
                    case AlignmentCorrectionKind.ConfirmHeaderRow: return HeaderRow.ConfirmationId;
                    case AlignmentCorrectionKind.ConfirmKeyColumns: return KeyColumns.ConfirmationId;
                    case AlignmentCorrectionKind.ClearSheetPair:
                    case AlignmentCorrectionKind.ClearSheetUnmatched:
                    case AlignmentCorrectionKind.ClearHeaderRow:
                    case AlignmentCorrectionKind.ClearKeyColumns:
                        return "";
                    default: throw new ArgumentOutOfRangeException();
                }
            }
        }

        private static string Required(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact alignment identity is required.", name);
            return value;
        }
    }

    public sealed class AlignmentCorrectionEvent
    {
        internal AlignmentCorrectionEvent(
            string eventId,
            string priorGenerationId,
            string newGenerationId,
            string supersededConfirmationId,
            AlignmentCorrectionRequest request)
        {
            EventId = eventId;
            PriorGenerationId = priorGenerationId;
            NewGenerationId = newGenerationId;
            Kind = request.Kind;
            RequestId = request.RequestId;
            Scope = request.Scope;
            ConfirmationId = request.ConfirmationId;
            SupersededConfirmationId = supersededConfirmationId ?? "";
            SheetPair = request.SheetPair;
            SheetUnmatched = request.SheetUnmatched;
            HeaderRow = request.HeaderRow;
            KeyColumns = request.KeyColumns;
            Reviewer = request.Reviewer;
            CreatedUtc = request.CreatedUtc;
            Note = request.Note;
        }

        public string EventId { get; }
        public string PriorGenerationId { get; }
        public string NewGenerationId { get; }
        public AlignmentCorrectionKind Kind { get; }
        public string RequestId { get; }
        public AlignmentCorrectionScope Scope { get; }
        /// <summary>Empty for a clear-to-auto event because it creates no new confirmation.</summary>
        public string ConfirmationId { get; }
        /// <summary>Non-empty when this event replaces or clears the active choice for the exact scope.</summary>
        public string SupersededConfirmationId { get; }
        public AlignmentSheetPairConfirmation SheetPair { get; }
        public AlignmentSheetUnmatchedConfirmation SheetUnmatched { get; }
        public AlignmentHeaderRowConfirmation HeaderRow { get; }
        public AlignmentKeyColumnsConfirmation KeyColumns { get; }
        public ReviewerSnapshot Reviewer { get; }
        public DateTimeOffset CreatedUtc { get; }
        public string Note { get; }
    }

    /// <summary>
    /// Immutable alignment override generation tied to one exact lineage mapping generation. Confirmations use
    /// physical sheet occurrence keys and never fall back to names, positions, or another revision occurrence.
    /// </summary>
    public sealed class AlignmentCorrectionGeneration
    {
        internal AlignmentCorrectionGeneration(
            string generationId,
            string priorGenerationId,
            string baseMappingGenerationId,
            IEnumerable<AlignmentSheetPairConfirmation> sheetPairs,
            IEnumerable<AlignmentSheetUnmatchedConfirmation> unmatchedSheets,
            IEnumerable<AlignmentHeaderRowConfirmation> headerRows,
            IEnumerable<AlignmentKeyColumnsConfirmation> keyColumns,
            IEnumerable<AlignmentCorrectionEvent> events)
        {
            GenerationId = generationId;
            PriorGenerationId = priorGenerationId ?? "";
            BaseMappingGenerationId = baseMappingGenerationId;
            SheetPairs = Freeze(sheetPairs);
            UnmatchedSheets = Freeze(unmatchedSheets);
            HeaderRows = Freeze(headerRows);
            KeyColumns = Freeze(keyColumns);
            Events = Freeze(events);
        }

        public string GenerationId { get; }
        public string PriorGenerationId { get; }
        public string BaseMappingGenerationId { get; }
        public IReadOnlyList<AlignmentSheetPairConfirmation> SheetPairs { get; }
        public IReadOnlyList<AlignmentSheetUnmatchedConfirmation> UnmatchedSheets { get; }
        public IReadOnlyList<AlignmentHeaderRowConfirmation> HeaderRows { get; }
        public IReadOnlyList<AlignmentKeyColumnsConfirmation> KeyColumns { get; }
        public IReadOnlyList<AlignmentCorrectionEvent> Events { get; }

        public IReadOnlyList<AlignmentSheetPairConfirmation> GetSheetPairs(
            string parentRevisionId,
            string childRevisionId) =>
            new ReadOnlyCollection<AlignmentSheetPairConfirmation>(SheetPairs.Where(item =>
                    string.Equals(item.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
                    string.Equals(item.ChildRevisionId, childRevisionId, StringComparison.Ordinal))
                .OrderBy(item => item.BeforeSheetKey, StringComparer.Ordinal)
                .ThenBy(item => item.AfterSheetKey, StringComparer.Ordinal)
                .ToList());

        public IReadOnlyList<AlignmentSheetUnmatchedConfirmation> GetUnmatchedSheets(
            string parentRevisionId,
            string childRevisionId) =>
            new ReadOnlyCollection<AlignmentSheetUnmatchedConfirmation>(UnmatchedSheets.Where(item =>
                    string.Equals(item.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
                    string.Equals(item.ChildRevisionId, childRevisionId, StringComparison.Ordinal))
                .OrderBy(item => item.BeforeSheetKey, StringComparer.Ordinal)
                .ToList());

        public bool TryGetHeaderRow(
            string revisionId,
            string physicalSheetKey,
            out AlignmentHeaderRowConfirmation confirmation)
        {
            confirmation = HeaderRows.SingleOrDefault(item =>
                string.Equals(item.RevisionId, revisionId, StringComparison.Ordinal) &&
                string.Equals(item.PhysicalSheetKey, physicalSheetKey, StringComparison.Ordinal));
            return confirmation != null;
        }

        public bool TryGetKeyColumns(
            string parentRevisionId,
            string childRevisionId,
            string beforeSheetKey,
            string afterSheetKey,
            out AlignmentKeyColumnsConfirmation confirmation)
        {
            confirmation = KeyColumns.SingleOrDefault(item =>
                string.Equals(item.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.ChildRevisionId, childRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.BeforeSheetKey, beforeSheetKey, StringComparison.Ordinal) &&
                string.Equals(item.AfterSheetKey, afterSheetKey, StringComparison.Ordinal));
            return confirmation != null;
        }

        private static IReadOnlyList<T> Freeze<T>(IEnumerable<T> values) =>
            new ReadOnlyCollection<T>(new List<T>(values ?? Array.Empty<T>()));
    }

    public static class AlignmentCorrectionEngine
    {
        public static AlignmentCorrectionGeneration Create(string baseMappingGenerationId)
        {
            if (string.IsNullOrWhiteSpace(baseMappingGenerationId))
                throw new ArgumentException("A lineage mapping generation ID is required.", nameof(baseMappingGenerationId));
            string id = StableId.Compute(
                "alignment-correction-generation-root-v1",
                baseMappingGenerationId);
            return new AlignmentCorrectionGeneration(
                id,
                "",
                baseMappingGenerationId,
                Array.Empty<AlignmentSheetPairConfirmation>(),
                Array.Empty<AlignmentSheetUnmatchedConfirmation>(),
                Array.Empty<AlignmentHeaderRowConfirmation>(),
                Array.Empty<AlignmentKeyColumnsConfirmation>(),
                Array.Empty<AlignmentCorrectionEvent>());
        }

        public static AlignmentCorrectionGeneration Apply(
            AlignmentCorrectionGeneration current,
            AlignmentCorrectionRequest request)
        {
            if (current == null) throw new ArgumentNullException(nameof(current));
            if (request == null) throw new ArgumentNullException(nameof(request));

            var sheetPairs = new List<AlignmentSheetPairConfirmation>(current.SheetPairs);
            var unmatchedSheets = new List<AlignmentSheetUnmatchedConfirmation>(current.UnmatchedSheets);
            var headers = new List<AlignmentHeaderRowConfirmation>(current.HeaderRows);
            var keys = new List<AlignmentKeyColumnsConfirmation>(current.KeyColumns);
            string supersededConfirmationId;
            switch (request.Kind)
            {
                case AlignmentCorrectionKind.ConfirmSheetPair:
                    supersededConfirmationId = AddSheetPair(
                        sheetPairs,
                        unmatchedSheets,
                        request.SheetPair);
                    break;
                case AlignmentCorrectionKind.ConfirmSheetUnmatched:
                    supersededConfirmationId = AddSheetUnmatched(
                        unmatchedSheets,
                        sheetPairs,
                        request.SheetUnmatched);
                    break;
                case AlignmentCorrectionKind.ConfirmHeaderRow:
                    supersededConfirmationId = AddHeader(headers, request.HeaderRow);
                    break;
                case AlignmentCorrectionKind.ConfirmKeyColumns:
                    supersededConfirmationId = AddKey(keys, request.KeyColumns);
                    break;
                case AlignmentCorrectionKind.ClearSheetPair:
                    supersededConfirmationId = RemoveSheetPair(sheetPairs, request.Scope);
                    break;
                case AlignmentCorrectionKind.ClearSheetUnmatched:
                    supersededConfirmationId = RemoveSheetUnmatched(unmatchedSheets, request.Scope);
                    break;
                case AlignmentCorrectionKind.ClearHeaderRow:
                    supersededConfirmationId = RemoveHeader(headers, request.Scope);
                    break;
                case AlignmentCorrectionKind.ClearKeyColumns:
                    supersededConfirmationId = RemoveKey(keys, request.Scope);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(request));
            }

            var eventComponents = new List<string>
            {
                current.GenerationId,
                ((int)request.Kind).ToString(CultureInfo.InvariantCulture),
                request.RequestId,
                request.Reviewer.ConfirmedDisplayName,
                request.Reviewer.OfficeDisplayName,
                request.CreatedUtc.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture),
                request.Note,
                supersededConfirmationId,
            };
            eventComponents.AddRange(request.Reviewer.IdentitySourceLabels.Select(item => "reviewer-source:" + item));
            string eventId = StableId.Compute("alignment-correction-event-v1", eventComponents.ToArray());
            string generationId = StableId.Compute(
                "alignment-correction-generation-v1",
                current.GenerationId,
                eventId);
            var events = new List<AlignmentCorrectionEvent>(current.Events)
            {
                new AlignmentCorrectionEvent(
                    eventId,
                    current.GenerationId,
                    generationId,
                    supersededConfirmationId,
                    request),
            };
            return new AlignmentCorrectionGeneration(
                generationId,
                current.GenerationId,
                current.BaseMappingGenerationId,
                sheetPairs,
                unmatchedSheets,
                headers,
                keys,
                events);
        }

        private static string AddSheetPair(
            IList<AlignmentSheetPairConfirmation> existing,
            IList<AlignmentSheetUnmatchedConfirmation> unmatched,
            AlignmentSheetPairConfirmation added)
        {
            AlignmentSheetPairConfirmation byBefore = existing.SingleOrDefault(item => SameEdge(item, added) &&
                string.Equals(item.BeforeSheetKey, added.BeforeSheetKey, StringComparison.Ordinal));
            AlignmentSheetPairConfirmation byAfter = existing.SingleOrDefault(item => SameEdge(item, added) &&
                string.Equals(item.AfterSheetKey, added.AfterSheetKey, StringComparison.Ordinal));
            if (byBefore != null && string.Equals(
                    byBefore.ConfirmationId,
                    added.ConfirmationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("This exact sheet-pair confirmation already exists.");
            if (byAfter != null && !ReferenceEquals(byAfter, byBefore))
                throw new InvalidOperationException(
                    "The confirmed target sheet is already paired with another source on this edge.");
            AlignmentSheetUnmatchedConfirmation priorUnmatched = unmatched.SingleOrDefault(item =>
                SameEdge(item, added.ParentRevisionId, added.ChildRevisionId) &&
                string.Equals(item.BeforeSheetKey, added.BeforeSheetKey, StringComparison.Ordinal));
            if (byBefore != null && priorUnmatched != null)
                throw new InvalidOperationException(
                    "The correction generation contains conflicting sheet resolutions for one exact source scope.");
            string superseded = "";
            if (byBefore != null)
            {
                superseded = byBefore.ConfirmationId;
                existing.Remove(byBefore);
            }
            else if (priorUnmatched != null)
            {
                superseded = priorUnmatched.ConfirmationId;
                unmatched.Remove(priorUnmatched);
            }
            existing.Add(added);
            return superseded;
        }

        private static string AddSheetUnmatched(
            IList<AlignmentSheetUnmatchedConfirmation> existing,
            IList<AlignmentSheetPairConfirmation> pairs,
            AlignmentSheetUnmatchedConfirmation added)
        {
            AlignmentSheetUnmatchedConfirmation prior = existing.SingleOrDefault(item =>
                SameEdge(item, added.ParentRevisionId, added.ChildRevisionId) &&
                string.Equals(item.BeforeSheetKey, added.BeforeSheetKey, StringComparison.Ordinal));
            if (prior != null && string.Equals(
                    prior.ConfirmationId,
                    added.ConfirmationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("This exact unmatched-sheet confirmation already exists.");
            AlignmentSheetPairConfirmation priorPair = pairs.SingleOrDefault(item =>
                SameEdge(item, added.ParentRevisionId, added.ChildRevisionId) &&
                string.Equals(item.BeforeSheetKey, added.BeforeSheetKey, StringComparison.Ordinal));
            if (prior != null && priorPair != null)
                throw new InvalidOperationException(
                    "The correction generation contains conflicting sheet resolutions for one exact source scope.");

            string superseded = "";
            if (prior != null)
            {
                superseded = prior.ConfirmationId;
                existing.Remove(prior);
            }
            else if (priorPair != null)
            {
                superseded = priorPair.ConfirmationId;
                pairs.Remove(priorPair);
            }
            existing.Add(added);
            return superseded;
        }

        private static string AddHeader(
            IList<AlignmentHeaderRowConfirmation> existing,
            AlignmentHeaderRowConfirmation added)
        {
            AlignmentHeaderRowConfirmation prior = existing.SingleOrDefault(item =>
                string.Equals(item.RevisionId, added.RevisionId, StringComparison.Ordinal) &&
                string.Equals(item.PhysicalSheetKey, added.PhysicalSheetKey, StringComparison.Ordinal));
            if (prior != null)
            {
                if (string.Equals(prior.ConfirmationId, added.ConfirmationId, StringComparison.Ordinal))
                    throw new InvalidOperationException("This exact header confirmation already exists.");
                existing.Remove(prior);
            }
            existing.Add(added);
            return prior?.ConfirmationId ?? "";
        }

        private static string AddKey(
            IList<AlignmentKeyColumnsConfirmation> existing,
            AlignmentKeyColumnsConfirmation added)
        {
            AlignmentKeyColumnsConfirmation prior = existing.SingleOrDefault(item =>
                string.Equals(item.ParentRevisionId, added.ParentRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.ChildRevisionId, added.ChildRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.BeforeSheetKey, added.BeforeSheetKey, StringComparison.Ordinal) &&
                string.Equals(item.AfterSheetKey, added.AfterSheetKey, StringComparison.Ordinal));
            if (prior != null)
            {
                if (string.Equals(prior.ConfirmationId, added.ConfirmationId, StringComparison.Ordinal))
                    throw new InvalidOperationException("This exact key-column confirmation already exists.");
                existing.Remove(prior);
            }
            existing.Add(added);
            return prior?.ConfirmationId ?? "";
        }

        private static string RemoveSheetPair(
            IList<AlignmentSheetPairConfirmation> existing,
            AlignmentCorrectionScope scope)
        {
            AlignmentSheetPairConfirmation prior = existing.SingleOrDefault(item =>
                string.Equals(item.ParentRevisionId, scope.ParentRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.ChildRevisionId, scope.ChildRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.BeforeSheetKey, scope.BeforeSheetKey, StringComparison.Ordinal));
            if (prior == null)
                throw new InvalidOperationException(
                    "No active sheet-pair confirmation exists for this exact source scope.");
            existing.Remove(prior);
            return prior.ConfirmationId;
        }

        private static string RemoveSheetUnmatched(
            IList<AlignmentSheetUnmatchedConfirmation> existing,
            AlignmentCorrectionScope scope)
        {
            AlignmentSheetUnmatchedConfirmation prior = existing.SingleOrDefault(item =>
                string.Equals(item.ParentRevisionId, scope.ParentRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.ChildRevisionId, scope.ChildRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.BeforeSheetKey, scope.BeforeSheetKey, StringComparison.Ordinal));
            if (prior == null)
                throw new InvalidOperationException(
                    "No active unmatched-sheet confirmation exists for this exact source scope.");
            existing.Remove(prior);
            return prior.ConfirmationId;
        }

        private static string RemoveHeader(
            IList<AlignmentHeaderRowConfirmation> existing,
            AlignmentCorrectionScope scope)
        {
            AlignmentHeaderRowConfirmation prior = existing.SingleOrDefault(item =>
                string.Equals(item.RevisionId, scope.RevisionId, StringComparison.Ordinal) &&
                string.Equals(item.PhysicalSheetKey, scope.PhysicalSheetKey, StringComparison.Ordinal));
            if (prior == null)
                throw new InvalidOperationException(
                    "No active header confirmation exists for this exact sheet scope.");
            existing.Remove(prior);
            return prior.ConfirmationId;
        }

        private static string RemoveKey(
            IList<AlignmentKeyColumnsConfirmation> existing,
            AlignmentCorrectionScope scope)
        {
            AlignmentKeyColumnsConfirmation prior = existing.SingleOrDefault(item =>
                string.Equals(item.ParentRevisionId, scope.ParentRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.ChildRevisionId, scope.ChildRevisionId, StringComparison.Ordinal) &&
                string.Equals(item.BeforeSheetKey, scope.BeforeSheetKey, StringComparison.Ordinal) &&
                string.Equals(item.AfterSheetKey, scope.AfterSheetKey, StringComparison.Ordinal));
            if (prior == null)
                throw new InvalidOperationException(
                    "No active key-column confirmation exists for this exact mapped-sheet scope.");
            existing.Remove(prior);
            return prior.ConfirmationId;
        }

        private static bool SameEdge(
            AlignmentSheetPairConfirmation left,
            AlignmentSheetPairConfirmation right) =>
            string.Equals(left.ParentRevisionId, right.ParentRevisionId, StringComparison.Ordinal) &&
            string.Equals(left.ChildRevisionId, right.ChildRevisionId, StringComparison.Ordinal);

        private static bool SameEdge(
            AlignmentSheetPairConfirmation left,
            string parentRevisionId,
            string childRevisionId) =>
            string.Equals(left.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
            string.Equals(left.ChildRevisionId, childRevisionId, StringComparison.Ordinal);

        private static bool SameEdge(
            AlignmentSheetUnmatchedConfirmation left,
            string parentRevisionId,
            string childRevisionId) =>
            string.Equals(left.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
            string.Equals(left.ChildRevisionId, childRevisionId, StringComparison.Ordinal);
    }
}
