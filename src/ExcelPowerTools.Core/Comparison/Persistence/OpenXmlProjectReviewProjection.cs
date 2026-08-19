using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>
    /// Pure, portable projection from an Open XML comparison project into the visible Review.xlsx model.
    /// The supplied JSON remains authoritative; this class only derives deterministic display rows from exact
    /// aligned evidence. The original source locators are included because they are required review evidence.
    /// </summary>
    public static class OpenXmlProjectReviewProjector
    {
        public const string EngineSourceHint =
            "src/ExcelPowerTools.Core/Comparison/Project/OpenXmlComparisonProjectEngine.cs";

        public static ReviewWorkbookState Project(
            OpenXmlComparisonProjectResult result,
            string stateJson,
            IEnumerable<ReviewRevisionDisplayMetadata> revisions)
        {
            if (result == null) throw new ArgumentNullException(nameof(result));
            if (stateJson == null) throw new ArgumentNullException(nameof(stateJson));

            List<ReviewRevisionDisplayMetadata> revisionList = ValidateRevisions(result, revisions);
            ReviewWorkbookState state;
            if (result.History != null)
            {
                IReadOnlyList<ReviewSubjectDisplayMetadata> subjects = ResolveSubjectDisplays(result);
                state = ReviewHistoryProjector.Project(
                    result.History,
                    stateJson,
                    revisionList,
                    subjects);
            }
            else
            {
                state = ProjectBlockedState(result, stateJson, revisionList);
            }

            if (result.IsComplete)
            {
                state.ComparisonStatus = "Complete comparison — exact formula/value history is available.";
                state.DetailUnavailableReason =
                    "No formula or text change is available for character-level detail in this comparison.";
            }
            else if (result.History != null)
            {
                state.ComparisonStatus =
                    "Partial diagnostic comparison — Changes includes only safely aligned exact stored evidence. " +
                    "Merge is disabled; see Diagnostics.";
                state.DetailUnavailableReason =
                    "No formula or text change is available for character-level detail in the reviewable subset. " +
                    "See Diagnostics for excluded evidence.";
            }
            else
            {
                state.ComparisonStatus =
                    "Incomplete diagnostic comparison — no logical cell history is available. See Diagnostics.";
                state.DetailUnavailableReason =
                    "Character-level detail is unavailable because logical cell history could not be built. " +
                    "See Diagnostics for the blocking evidence.";
            }

            AppendStructuralRows(result, state);

            foreach (OpenXmlProjectFinding finding in result.Findings
                         .OrderBy(item => item.FindingId, StringComparer.Ordinal)
                         .ThenBy(item => item.Code, StringComparer.Ordinal))
                state.DiagnosticRows.Add(ProjectFinding(result, finding));

            return state;
        }

        private static void AppendStructuralRows(
            OpenXmlComparisonProjectResult result,
            ReviewWorkbookState state)
        {
            var revisions = result.Revisions.ToDictionary(
                item => item.Source.RevisionId,
                StringComparer.Ordinal);
            foreach (OpenXmlProjectEdgeResult edge in result.Edges
                         .OrderBy(item => revisions[item.ChildRevisionId].Source.Order)
                         .ThenBy(item => item.ChildRevisionId, StringComparer.Ordinal))
            {
                OpenXmlAlignedProjectRevision parent = revisions[edge.ParentRevisionId];
                OpenXmlAlignedProjectRevision child = revisions[edge.ChildRevisionId];

                foreach (ResolvedSheetMapping mapping in edge.SheetMapping.Mapping.Mappings
                             .Where(item => item.Kind == SheetMappingKind.InferredRename)
                             .OrderBy(item => item.Before.Name, StringComparer.Ordinal)
                             .ThenBy(item => item.After.Name, StringComparer.Ordinal))
                {
                    OpenXmlProjectSheetAlignment alignment = edge.SheetAlignments.FirstOrDefault(item =>
                        string.Equals(item.BeforeSheetKey, mapping.Before.SheetKey, StringComparison.Ordinal) &&
                        string.Equals(item.AfterSheetKey, mapping.After.SheetKey, StringComparison.Ordinal));
                    string logicalSheetId = alignment?.LogicalSheetId ?? StableId.Compute(
                        "review-sheet-rename-group-v1",
                        result.ProjectId,
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        mapping.Before.SheetKey,
                        mapping.After.SheetKey);
                    string changeId = StableId.Compute(
                        "review-sheet-rename-v1",
                        result.ProjectId,
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        mapping.Before.SheetKey,
                        mapping.After.SheetKey);
                    state.ChangeRows.Add(new ReviewChangeProjectionRow(
                        changeId,
                        mapping.After.Name,
                        "",
                        "Sheet name",
                        mapping.Before.Name,
                        mapping.After.Name,
                        logicalSheetId,
                        "Structural; Rename",
                        2,
                        2,
                        mapping.Confidence.ToString("0.00", CultureInfo.InvariantCulture)));
                }

                foreach (WorkbookSheetFingerprint deleted in edge.SheetMapping.Mapping.UnmatchedBeforeSheets
                             .OrderBy(item => item.Name, StringComparer.Ordinal)
                             .ThenBy(item => item.SheetKey, StringComparer.Ordinal))
                {
                    OpenXmlAlignedSheetRevision deletedSheet = parent.Sheets.FirstOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, deleted.SheetKey, StringComparison.Ordinal));
                    string logicalSheetId = deletedSheet?.LogicalSheetId ?? StableId.Compute(
                        "review-sheet-deletion-group-v1",
                        result.ProjectId,
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        deleted.SheetKey);
                    string changeId = StableId.Compute(
                        "review-sheet-deletion-v1",
                        result.ProjectId,
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        deleted.SheetKey);
                    state.ChangeRows.Add(new ReviewChangeProjectionRow(
                        changeId,
                        deleted.Name,
                        "",
                        "Sheet structure",
                        "Deleted sheet " + deleted.Name,
                        "",
                        logicalSheetId,
                        "Structural; Deletion",
                        2,
                        2,
                        "1.00"));
                }

                foreach (WorkbookSheetFingerprint inserted in edge.SheetMapping.Mapping.UnmatchedAfterSheets
                             .OrderBy(item => item.Name, StringComparer.Ordinal)
                             .ThenBy(item => item.SheetKey, StringComparer.Ordinal))
                {
                    OpenXmlAlignedSheetRevision insertedSheet = child.Sheets.FirstOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, inserted.SheetKey, StringComparison.Ordinal));
                    string logicalSheetId = insertedSheet?.LogicalSheetId ?? StableId.Compute(
                        "review-sheet-insertion-group-v1",
                        result.ProjectId,
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        inserted.SheetKey);
                    string changeId = StableId.Compute(
                        "review-sheet-insertion-v1",
                        result.ProjectId,
                        edge.ChildRevisionId,
                        inserted.SheetKey);
                    state.ChangeRows.Add(new ReviewChangeProjectionRow(
                        changeId,
                        inserted.Name,
                        "",
                        "Sheet structure",
                        "",
                        "Added sheet " + inserted.Name,
                        logicalSheetId,
                        "Structural; Insertion",
                        2,
                        2,
                        insertedSheet?.Rows?.Confidence.ToString() ?? "1.00"));
                }

                foreach (OpenXmlProjectSheetAlignment alignment in edge.SheetAlignments
                             .OrderBy(item => item.LogicalSheetId, StringComparer.Ordinal)
                             .ThenBy(item => item.AfterSheetName, StringComparer.Ordinal))
                {
                    OpenXmlAlignedSheetRevision beforeSheet = parent.Sheets.FirstOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, alignment.BeforeSheetKey, StringComparison.Ordinal));
                    OpenXmlAlignedSheetRevision afterSheet = child.Sheets.FirstOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, alignment.AfterSheetKey, StringComparison.Ordinal));
                    AppendAxisEvents(
                        state,
                        alignment.Rows,
                        LogicalAxisKind.Row,
                        beforeSheet,
                        afterSheet,
                        alignment.AfterSheetName.Length == 0
                            ? alignment.BeforeSheetName
                            : alignment.AfterSheetName,
                        alignment.LogicalSheetId);
                    AppendAxisEvents(
                        state,
                        alignment.Columns,
                        LogicalAxisKind.Column,
                        beforeSheet,
                        afterSheet,
                        alignment.AfterSheetName.Length == 0
                            ? alignment.BeforeSheetName
                            : alignment.AfterSheetName,
                        alignment.LogicalSheetId);
                }
            }
        }

        private static void AppendAxisEvents(
            ReviewWorkbookState state,
            LogicalAxisIdentityProjection projection,
            LogicalAxisKind axis,
            OpenXmlAlignedSheetRevision beforeSheet,
            OpenXmlAlignedSheetRevision afterSheet,
            string sheetName,
            string logicalSheetId)
        {
            if (projection == null) return;
            foreach (LogicalAxisStructuralEvent item in projection.StructuralEvents
                         .OrderBy(entry => entry.EventId, StringComparer.Ordinal))
            {
                string before = PhysicalAxisLabel(axis, beforeSheet, item.BeforeIndex);
                string after = PhysicalAxisLabel(axis, afterSheet, item.AfterIndex);
                string address = after.Length == 0 ? before : after;
                string beforeText;
                string afterText;
                switch (item.Kind)
                {
                    case LogicalAxisStructuralEventKind.Insertion:
                        beforeText = "";
                        afterText = "Inserted " + after;
                        break;
                    case LogicalAxisStructuralEventKind.Deletion:
                        beforeText = "Deleted " + before;
                        afterText = "";
                        break;
                    default:
                        beforeText = before;
                        afterText = after;
                        break;
                }
                state.ChangeRows.Add(new ReviewChangeProjectionRow(
                    item.EventId,
                    sheetName,
                    address,
                    axis == LogicalAxisKind.Row ? "Row structure" : "Column structure",
                    beforeText,
                    afterText,
                    logicalSheetId,
                    "Structural; " + item.Kind,
                    2,
                    2,
                    item.Confidence.ToString()));
            }
        }

        private static string PhysicalAxisLabel(
            LogicalAxisKind axis,
            OpenXmlAlignedSheetRevision sheet,
            int? tableIndex)
        {
            if (sheet == null || !tableIndex.HasValue || tableIndex.Value < 0) return "";
            if (axis == LogicalAxisKind.Row)
            {
                if (tableIndex.Value >= sheet.Grid.Table.RowCount) return "";
                return "Row " + sheet.Grid.PhysicalRowForTableIndex(tableIndex.Value)
                    .ToString(CultureInfo.InvariantCulture);
            }
            if (tableIndex.Value >= sheet.Grid.Table.ColumnCount) return "";
            int physical = sheet.Grid.PhysicalColumnForTableIndex(tableIndex.Value);
            return "Column " + ColumnName(physical);
        }

        private static string ColumnName(int column)
        {
            if (column <= 0) return "";
            var letters = new StringBuilder();
            int remaining = column;
            while (remaining > 0)
            {
                remaining--;
                letters.Insert(0, (char)('A' + remaining % 26));
                remaining /= 26;
            }
            return letters.ToString();
        }

        private static List<ReviewRevisionDisplayMetadata> ValidateRevisions(
            OpenXmlComparisonProjectResult result,
            IEnumerable<ReviewRevisionDisplayMetadata> revisions)
        {
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));
            List<ReviewRevisionDisplayMetadata> revisionList = revisions.ToList();
            if (revisionList.Any(item => item == null))
                throw new ArgumentException(
                    "Revision display metadata cannot contain null entries.",
                    nameof(revisions));

            var metadataById = new Dictionary<string, ReviewRevisionDisplayMetadata>(StringComparer.Ordinal);
            var orders = new HashSet<int>();
            foreach (ReviewRevisionDisplayMetadata metadata in revisionList)
            {
                if (metadataById.ContainsKey(metadata.RevisionId))
                    throw new ArgumentException(
                        "Duplicate revision display metadata: " + metadata.RevisionId,
                        nameof(revisions));
                if (!orders.Add(metadata.Order))
                    throw new ArgumentException(
                        "Revision display orders must be unique.",
                        nameof(revisions));
                metadataById.Add(metadata.RevisionId, metadata);
            }

            if (metadataById.Count != result.Revisions.Count)
                throw new ArgumentException(
                    "Revision display metadata must contain exactly the project revision occurrences.",
                    nameof(revisions));

            foreach (OpenXmlAlignedProjectRevision aligned in result.Revisions)
            {
                OpenXmlProjectRevision source = aligned.Source;
                if (!metadataById.TryGetValue(source.RevisionId, out ReviewRevisionDisplayMetadata metadata))
                    throw new ArgumentException(
                        "Revision display metadata is missing project revision: " + source.RevisionId,
                        nameof(revisions));
                if (metadata.Order != source.Order ||
                    !string.Equals(metadata.ParentRevisionId, source.ParentRevisionId, StringComparison.Ordinal))
                    throw new ArgumentException(
                        "Revision display metadata does not match project order and parentage: " +
                        source.RevisionId,
                        nameof(revisions));
            }

            return revisionList;
        }

        private static IReadOnlyList<ReviewSubjectDisplayMetadata> ResolveSubjectDisplays(
            OpenXmlComparisonProjectResult result)
        {
            var expected = new HashSet<string>(
                result.History.Subjects.Select(item => item.LogicalSubjectKey),
                StringComparer.Ordinal);
            var resolved = new Dictionary<string, ReviewSubjectDisplayMetadata>(StringComparer.Ordinal);

            foreach (OpenXmlAlignedProjectRevision revision in result.Revisions
                         .OrderBy(item => string.Equals(
                             item.Source.RevisionId,
                             result.History.MasterRevisionId,
                             StringComparison.Ordinal) ? 0 : 1)
                         .ThenByDescending(item => item.Source.Order)
                         .ThenBy(item => item.Source.RevisionId, StringComparer.Ordinal))
            {
                Dictionary<string, ReviewSubjectDisplayMetadata> occurrences =
                    ResolveRevisionOccurrences(revision, expected);
                foreach (KeyValuePair<string, ReviewSubjectDisplayMetadata> occurrence in occurrences
                             .OrderBy(item => item.Key, StringComparer.Ordinal))
                {
                    if (!resolved.ContainsKey(occurrence.Key))
                        resolved.Add(occurrence.Key, occurrence.Value);
                }
            }

            string missing = expected
                .Where(item => !resolved.ContainsKey(item))
                .OrderBy(item => item, StringComparer.Ordinal)
                .FirstOrDefault();
            if (missing != null)
                throw new InvalidOperationException(
                    "No exact physical aligned occurrence exists for logical history subject " + missing +
                    ". The review projection will not invent a sheet or cell address.");

            return new ReadOnlyCollection<ReviewSubjectDisplayMetadata>(resolved.Values
                .OrderBy(item => item.LogicalSubjectKey, StringComparer.Ordinal)
                .ToList());
        }

        private static Dictionary<string, ReviewSubjectDisplayMetadata> ResolveRevisionOccurrences(
            OpenXmlAlignedProjectRevision revision,
            ISet<string> expected)
        {
            var occurrences = new Dictionary<string, ReviewSubjectDisplayMetadata>(StringComparer.Ordinal);
            foreach (OpenXmlAlignedSheetRevision sheet in revision.Sheets
                         .OrderBy(item => item.LogicalSheetId, StringComparer.Ordinal)
                         .ThenBy(item => item.PhysicalSheetKey, StringComparer.Ordinal))
            {
                IReadOnlyDictionary<int, string> rows = ToPhysicalIdentityMap(sheet.Grid, sheet.Rows, true);
                IReadOnlyDictionary<int, string> columns = ToPhysicalIdentityMap(sheet.Grid, sheet.Columns, false);
                foreach (AlignedCanonicalCellProjection cell in sheet.Grid
                             .CreateAlignedCells(sheet.LogicalSheetId, rows, columns)
                             .OrderBy(item => item.Cell.SubjectKey, StringComparer.Ordinal))
                {
                    string logicalSubjectKey = cell.Cell.SubjectKey;
                    if (!expected.Contains(logicalSubjectKey)) continue;
                    if (occurrences.ContainsKey(logicalSubjectKey))
                        throw new InvalidOperationException(
                            "Revision " + revision.Source.RevisionId +
                            " contains more than one physical occurrence for logical history subject " +
                            logicalSubjectKey + ".");

                    int rowIndex = checked(cell.PhysicalRow - sheet.Grid.FirstPhysicalRow);
                    int columnIndex = checked(cell.PhysicalColumn - sheet.Grid.FirstPhysicalColumn);
                    if (!sheet.Rows.AssignmentsByPhysicalIndex.TryGetValue(
                            rowIndex,
                            out LogicalAxisIdentityAssignment row) ||
                        !sheet.Columns.AssignmentsByPhysicalIndex.TryGetValue(
                            columnIndex,
                            out LogicalAxisIdentityAssignment column))
                        throw new InvalidOperationException(
                            "Aligned cell evidence lacks an exact row or column assignment in revision " +
                            revision.Source.RevisionId + ".");
                    if (!string.Equals(row.LogicalIdentityId, cell.RowIdentity, StringComparison.Ordinal) ||
                        !string.Equals(column.LogicalIdentityId, cell.ColumnIdentity, StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            "Aligned cell evidence disagrees with its logical axis assignments in revision " +
                            revision.Source.RevisionId + ".");

                    occurrences.Add(logicalSubjectKey, new ReviewSubjectDisplayMetadata(
                        logicalSubjectKey,
                        sheet.PhysicalSheetName,
                        ToA1(cell.PhysicalRow, cell.PhysicalColumn),
                        Minimum(row.Confidence, column.Confidence)));
                }
            }
            return occurrences;
        }

        private static IReadOnlyDictionary<int, string> ToPhysicalIdentityMap(
            CanonicalGridProjection grid,
            LogicalAxisIdentityProjection identities,
            bool rows)
        {
            var result = new Dictionary<int, string>();
            foreach (KeyValuePair<int, string> pair in identities.IdentityByPhysicalIndex
                         .OrderBy(item => item.Key))
            {
                int physicalIndex = rows
                    ? grid.PhysicalRowForTableIndex(pair.Key)
                    : grid.PhysicalColumnForTableIndex(pair.Key);
                result.Add(physicalIndex, pair.Value);
            }
            return new ReadOnlyDictionary<int, string>(result);
        }

        private static ReviewWorkbookState ProjectBlockedState(
            OpenXmlComparisonProjectResult result,
            string stateJson,
            IEnumerable<ReviewRevisionDisplayMetadata> revisions)
        {
            var state = new ReviewWorkbookState(result.ProjectId, stateJson);
            Dictionary<string, OpenXmlProjectRevision> sources = result.Revisions
                .Select(item => item.Source)
                .ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            OpenXmlProjectRevision master = sources.Values.Single(item => item.IsMaster);
            HashSet<string> masterLineage = BuildMasterLineage(master, sources);

            foreach (ReviewRevisionDisplayMetadata revision in revisions
                         .OrderBy(item => item.Order)
                         .ThenBy(item => item.RevisionId, StringComparer.Ordinal))
            {
                OpenXmlProjectRevision source = sources[revision.RevisionId];
                state.RevisionRows.Add(new ReviewRevisionProjectionRow(
                    revision.RevisionId,
                    revision.Label,
                    revision.Order,
                    revision.ParentRevisionId,
                    RevisionRole(source, masterLineage),
                    revision.SourceBaseName,
                    revision.SourcePath,
                    revision.ByteContentId,
                    revision.ModifiedUtc.HasValue
                        ? revision.ModifiedUtc.Value.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture)
                        : "",
                    revision.Inference,
                    revision.Confidence));
            }
            return state;
        }

        private static HashSet<string> BuildMasterLineage(
            OpenXmlProjectRevision master,
            IReadOnlyDictionary<string, OpenXmlProjectRevision> sources)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            OpenXmlProjectRevision cursor = master;
            while (cursor != null)
            {
                if (!result.Add(cursor.RevisionId))
                    throw new InvalidOperationException("The comparison project contains a revision parent cycle.");
                cursor = cursor.ParentRevisionId.Length == 0 ? null : sources[cursor.ParentRevisionId];
            }
            return result;
        }

        private static string RevisionRole(
            OpenXmlProjectRevision revision,
            ISet<string> masterLineage)
        {
            var roles = new List<string>();
            if (revision.IsMaster) roles.Add("Master");
            if (revision.IsTrustedBaseline) roles.Add("Trusted baseline");
            if (!masterLineage.Contains(revision.RevisionId)) roles.Add("Branch");
            return roles.Count == 0 ? "Revision" : string.Join("; ", roles);
        }

        private static ReviewDiagnosticProjectionRow ProjectFinding(
            OpenXmlComparisonProjectResult result,
            OpenXmlProjectFinding finding)
        {
            string reasons = JoinOrNone(finding.ReasonCodes);
            string candidates = JoinOrNone(finding.CandidateSheetNames);
            string revision = TransitionLabel(finding.BeforeRevisionId, finding.AfterRevisionId);
            string detail =
                "Finding ID: " + finding.FindingId + "\n" +
                "Project ID: " + result.ProjectId + "\n" +
                "Mapping generation ID: " + result.MappingGenerationId + "\n" +
                "Before revision: " + ValueOrNone(finding.BeforeRevisionId) + "\n" +
                "After revision: " + ValueOrNone(finding.AfterRevisionId) + "\n" +
                "Logical sheet ID: " + ValueOrNone(finding.LogicalSheetId) + "\n" +
                "Physical sheet: " + ValueOrNone(finding.SheetName) + "\n" +
                "Reason codes: " + reasons + "\n" +
                "Candidate sheets: " + candidates + "\n" +
                "Message: " + ValueOrNone(finding.Message) + "\n" +
                "Suggested action: " + ValueOrNone(finding.SuggestedAction) + "\n" +
                "Engine source: " + EngineSourceHint;

            return new ReviewDiagnosticProjectionRow(
                finding.FindingId,
                finding.Code,
                finding.Severity.ToString(),
                revision,
                finding.Message,
                detail,
                finding.SuggestedAction);
        }

        private static string TransitionLabel(string beforeRevisionId, string afterRevisionId)
        {
            if (beforeRevisionId.Length == 0) return afterRevisionId;
            if (afterRevisionId.Length == 0) return beforeRevisionId;
            return beforeRevisionId + " -> " + afterRevisionId;
        }

        private static string JoinOrNone(IEnumerable<string> values)
        {
            string[] ordered = (values ?? Array.Empty<string>())
                .Where(item => !string.IsNullOrEmpty(item))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToArray();
            return ordered.Length == 0 ? "(none)" : string.Join(", ", ordered);
        }

        private static string ValueOrNone(string value) =>
            string.IsNullOrEmpty(value) ? "(none)" : value;

        private static AlignmentConfidence Minimum(
            AlignmentConfidence left,
            AlignmentConfidence right) =>
            left <= right ? left : right;

        private static string ToA1(int row, int column)
        {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));

            var letters = new StringBuilder();
            int remaining = column;
            while (remaining > 0)
            {
                remaining--;
                letters.Insert(0, (char)('A' + remaining % 26));
                remaining /= 26;
            }
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}
