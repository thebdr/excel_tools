using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;

namespace ExcelPowerTools.Core.Comparison.Diff
{
    public enum ExactComparisonStatus
    {
        Equal,
        Different,
        Indeterminate,
    }

    public enum ExactSheetMatchKind
    {
        ExactName,
        OnlyBefore,
        OnlyAfter,
        AmbiguousDuplicateName,
    }

    /// <summary>
    /// Exact-name, physical-sheet comparison for the pre-alignment phase. Renames and duplicate names are
    /// reported, never guessed. Atomic cell facts remain available independently of this sheet projection.
    /// </summary>
    public sealed class ExactSheetComparison
    {
        internal ExactSheetComparison(
            string sheetName,
            ExactSheetMatchKind matchKind,
            ExactComparisonStatus status,
            OpenXmlWorksheetSnapshot before,
            OpenXmlWorksheetSnapshot after,
            IReadOnlyList<AtomicCellChange> changes)
        {
            SheetName = sheetName ?? "";
            MatchKind = matchKind;
            Status = status;
            Before = before;
            After = after;
            Changes = new ReadOnlyCollection<AtomicCellChange>(
                new List<AtomicCellChange>(changes ?? Array.Empty<AtomicCellChange>()));
        }

        public string SheetName { get; }
        public ExactSheetMatchKind MatchKind { get; }
        public ExactComparisonStatus Status { get; }
        public OpenXmlWorksheetSnapshot Before { get; }
        public OpenXmlWorksheetSnapshot After { get; }
        public IReadOnlyList<AtomicCellChange> Changes { get; }
        public bool VisibilityChanged => Before != null && After != null &&
                                         Before.Visibility != After.Visibility;
    }

    public sealed class ExactWorkbookComparisonResult
    {
        internal ExactWorkbookComparisonResult(
            string beforeRevisionId,
            string afterRevisionId,
            bool byteIdentical,
            ExactComparisonStatus status,
            IReadOnlyList<ExactSheetComparison> sheets)
        {
            BeforeRevisionId = beforeRevisionId ?? "";
            AfterRevisionId = afterRevisionId ?? "";
            ByteIdentical = byteIdentical;
            Status = status;
            Sheets = sheets;
        }

        public string BeforeRevisionId { get; }
        public string AfterRevisionId { get; }
        public bool ByteIdentical { get; }
        public ExactComparisonStatus Status { get; }
        public IReadOnlyList<ExactSheetComparison> Sheets { get; }
        public bool CanAssertExactCellEquality => Status == ExactComparisonStatus.Equal;
        public int AtomicChangeCount => Sheets.Sum(sheet => sheet.Changes.Count);
    }

    /// <summary>
    /// Deterministic two-revision harness over read-only OOXML snapshots. This is deliberately pre-alignment:
    /// only one-to-one ordinal sheet-name matches are compared, while every other mapping remains explicit.
    /// </summary>
    public static class ExactWorkbookComparisonEngine
    {
        public static ExactWorkbookComparisonResult Compare(
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after,
            ExactComparisonOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (after == null) throw new ArgumentNullException(nameof(after));
            cancellationToken.ThrowIfCancellationRequested();

            bool byteIdentical = string.Equals(
                before.ByteContentId,
                after.ByteContentId,
                StringComparison.Ordinal);
            Dictionary<string, List<OpenXmlWorksheetSnapshot>> beforeByName = GroupByName(before.Sheets);
            Dictionary<string, List<OpenXmlWorksheetSnapshot>> afterByName = GroupByName(after.Sheets);
            var names = new SortedSet<string>(beforeByName.Keys, StringComparer.Ordinal);
            names.UnionWith(afterByName.Keys);

            var sheetResults = new List<ExactSheetComparison>(names.Count);
            bool observedDifference = false;
            bool indeterminate = !byteIdentical &&
                                 (!before.CanAssertExactCellEquality || !after.CanAssertExactCellEquality);

            foreach (string name in names)
            {
                cancellationToken.ThrowIfCancellationRequested();
                beforeByName.TryGetValue(name, out List<OpenXmlWorksheetSnapshot> beforeMatches);
                afterByName.TryGetValue(name, out List<OpenXmlWorksheetSnapshot> afterMatches);
                beforeMatches = beforeMatches ?? new List<OpenXmlWorksheetSnapshot>();
                afterMatches = afterMatches ?? new List<OpenXmlWorksheetSnapshot>();

                if (beforeMatches.Count > 1 || afterMatches.Count > 1)
                {
                    sheetResults.Add(new ExactSheetComparison(
                        name,
                        ExactSheetMatchKind.AmbiguousDuplicateName,
                        byteIdentical ? ExactComparisonStatus.Equal : ExactComparisonStatus.Indeterminate,
                        beforeMatches.FirstOrDefault(),
                        afterMatches.FirstOrDefault(),
                        EmptyChanges()));
                    if (!byteIdentical) indeterminate = true;
                    continue;
                }

                if (beforeMatches.Count == 0)
                {
                    observedDifference = true;
                    sheetResults.Add(new ExactSheetComparison(
                        name,
                        ExactSheetMatchKind.OnlyAfter,
                        ExactComparisonStatus.Different,
                        null,
                        afterMatches[0],
                        EmptyChanges()));
                    continue;
                }

                if (afterMatches.Count == 0)
                {
                    observedDifference = true;
                    sheetResults.Add(new ExactSheetComparison(
                        name,
                        ExactSheetMatchKind.OnlyBefore,
                        ExactComparisonStatus.Different,
                        beforeMatches[0],
                        null,
                        EmptyChanges()));
                    continue;
                }

                OpenXmlWorksheetSnapshot beforeSheet = beforeMatches[0];
                OpenXmlWorksheetSnapshot afterSheet = afterMatches[0];
                if (beforeSheet.Kind != OpenXmlPhysicalSheetKind.Worksheet ||
                    afterSheet.Kind != OpenXmlPhysicalSheetKind.Worksheet)
                {
                    bool kindChanged = beforeSheet.Kind != afterSheet.Kind;
                    observedDifference |= kindChanged;
                    if (!byteIdentical && !kindChanged) indeterminate = true;
                    sheetResults.Add(new ExactSheetComparison(
                        name,
                        ExactSheetMatchKind.ExactName,
                        byteIdentical
                            ? ExactComparisonStatus.Equal
                            : kindChanged
                                ? ExactComparisonStatus.Different
                                : ExactComparisonStatus.Indeterminate,
                        beforeSheet,
                        afterSheet,
                        EmptyChanges()));
                    continue;
                }

                IReadOnlyList<AtomicCellChange> changes = ExactComparisonEngine.Compare(
                    beforeSheet.Snapshot,
                    afterSheet.Snapshot,
                    options,
                    cancellationToken);
                bool sheetDifferent = changes.Count > 0;
                bool sheetIndeterminate = !byteIdentical && !sheetDifferent &&
                                          (!beforeSheet.CanAssertExactCellEquality ||
                                           !afterSheet.CanAssertExactCellEquality);
                observedDifference |= sheetDifferent;
                indeterminate |= sheetIndeterminate;
                sheetResults.Add(new ExactSheetComparison(
                    name,
                    ExactSheetMatchKind.ExactName,
                    byteIdentical
                        ? ExactComparisonStatus.Equal
                        : sheetDifferent
                            ? ExactComparisonStatus.Different
                            : sheetIndeterminate
                                ? ExactComparisonStatus.Indeterminate
                                : ExactComparisonStatus.Equal,
                    beforeSheet,
                    afterSheet,
                    changes));
            }

            ExactComparisonStatus status = byteIdentical
                ? ExactComparisonStatus.Equal
                : observedDifference
                    ? ExactComparisonStatus.Different
                    : indeterminate
                        ? ExactComparisonStatus.Indeterminate
                        : ExactComparisonStatus.Equal;
            return new ExactWorkbookComparisonResult(
                before.RevisionId,
                after.RevisionId,
                byteIdentical,
                status,
                new ReadOnlyCollection<ExactSheetComparison>(sheetResults));
        }

        private static Dictionary<string, List<OpenXmlWorksheetSnapshot>> GroupByName(
            IReadOnlyList<OpenXmlWorksheetSnapshot> sheets)
        {
            var result = new Dictionary<string, List<OpenXmlWorksheetSnapshot>>(StringComparer.Ordinal);
            foreach (OpenXmlWorksheetSnapshot sheet in sheets)
            {
                if (!result.TryGetValue(sheet.Name, out List<OpenXmlWorksheetSnapshot> matches))
                {
                    matches = new List<OpenXmlWorksheetSnapshot>();
                    result.Add(sheet.Name, matches);
                }
                matches.Add(sheet);
            }
            return result;
        }

        private static IReadOnlyList<AtomicCellChange> EmptyChanges() =>
            new ReadOnlyCollection<AtomicCellChange>(Array.Empty<AtomicCellChange>());
    }
}
