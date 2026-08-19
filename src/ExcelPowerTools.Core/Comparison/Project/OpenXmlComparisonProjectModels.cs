using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.Project
{
    /// <summary>Stable, UI-agnostic stages emitted while logical workbook identity is being built.</summary>
    public enum OpenXmlComparisonProjectProgressStage
    {
        Preparing,
        FingerprintingSheets,
        ProjectingGrid,
        MappingSheets,
        DetectingHeaders,
        InferringKeys,
        AligningRows,
        AligningColumns,
        PropagatingIdentities,
        CheckingBranches,
        BuildingHistory,
        Complete,
    }

    /// <summary>
    /// Bounded project-engine progress. Work items are physical sheets across all revisions; stage/detail can
    /// advance while the completed count remains unchanged during one expensive sheet.
    /// </summary>
    public sealed class OpenXmlComparisonProjectProgress
    {
        public OpenXmlComparisonProjectProgress(
            OpenXmlComparisonProjectProgressStage stage,
            int completedWorkItems,
            int totalWorkItems,
            string revisionId,
            string sheetName,
            string detail)
        {
            if (totalWorkItems <= 0) throw new ArgumentOutOfRangeException(nameof(totalWorkItems));
            if (completedWorkItems < 0 || completedWorkItems > totalWorkItems)
                throw new ArgumentOutOfRangeException(nameof(completedWorkItems));
            Stage = stage;
            CompletedWorkItems = completedWorkItems;
            TotalWorkItems = totalWorkItems;
            RevisionId = revisionId ?? "";
            SheetName = sheetName ?? "";
            Detail = detail ?? "";
        }

        public OpenXmlComparisonProjectProgressStage Stage { get; }
        public int CompletedWorkItems { get; }
        public int TotalWorkItems { get; }
        public string RevisionId { get; }
        public string SheetName { get; }
        public string Detail { get; }
    }

    /// <summary>One explicitly ordered occurrence in a rooted comparison revision tree.</summary>
    public sealed class OpenXmlProjectRevision
    {
        public OpenXmlProjectRevision(
            string revisionId,
            string parentRevisionId,
            int order,
            OpenXmlWorkbookSnapshot snapshot,
            bool isMaster = false,
            bool isTrustedBaseline = false)
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (order < 0) throw new ArgumentOutOfRangeException(nameof(order));
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            if (!string.Equals(revisionId, snapshot.RevisionId, StringComparison.Ordinal))
                throw new ArgumentException("The workbook snapshot revision ID must match its project revision.", nameof(snapshot));

            RevisionId = revisionId;
            ParentRevisionId = parentRevisionId ?? "";
            Order = order;
            IsMaster = isMaster;
            IsTrustedBaseline = isTrustedBaseline;
        }

        public string RevisionId { get; }
        public string ParentRevisionId { get; }
        public int Order { get; }
        public OpenXmlWorkbookSnapshot Snapshot { get; }
        public bool IsMaster { get; }
        public bool IsTrustedBaseline { get; }
    }

    public sealed class OpenXmlComparisonProjectOptions
    {
        public WorkbookSheetMappingOptions SheetMapping { get; set; } = WorkbookSheetMappingOptions.Default;
        public HeaderDetectionOptions HeaderDetection { get; set; } = new HeaderDetectionOptions();
        public KeyInferenceOptions KeyInference { get; set; } = new KeyInferenceOptions();
        public TabularAlignmentOptions Alignment { get; set; } = new TabularAlignmentOptions();
        public CanonicalGridProjectionOptions GridProjection { get; set; } = new CanonicalGridProjectionOptions();
        public ExactComparisonOptions ExactComparison { get; set; } = new ExactComparisonOptions();
        /// <summary>
        /// Optional immutable user-confirmation generation. It must be rooted in the exact lineage mapping ID
        /// supplied to the project build; stale generations are rejected before any override is applied.
        /// </summary>
        public AlignmentCorrectionGeneration AlignmentCorrections { get; set; }

        /// <summary>
        /// Weak positional links are useful suggestions but are unsafe as persistent logical identities unless
        /// explicitly confirmed. The project engine blocks them by default.
        /// </summary>
        public bool RequireAtLeastMediumAxisConfidence { get; set; } = true;
    }

    public static class OpenXmlProjectFindingCodes
    {
        public const string UnsupportedSheet = "CMP-PROJECT-SHEET-UNSUPPORTED";
        public const string IncompleteCellEvidence = "CMP-PROJECT-CELL-EVIDENCE-INCOMPLETE";
        public const string GridProjectionFailed = "CMP-PROJECT-GRID-PROJECTION-FAILED";
        public const string SheetMappingAmbiguous = "CMP-PROJECT-SHEET-MAPPING-AMBIGUOUS";
        public const string ParentSheetStateUnavailable = "CMP-PROJECT-PARENT-SHEET-STATE-UNAVAILABLE";
        public const string HeaderMismatch = "CMP-PROJECT-HEADER-MISMATCH";
        public const string NonLeadingHeader = "CMP-PROJECT-NONLEADING-HEADER";
        public const string RowAlignmentAmbiguous = "CMP-PROJECT-ROW-ALIGNMENT-AMBIGUOUS";
        public const string ColumnAlignmentAmbiguous = "CMP-PROJECT-COLUMN-ALIGNMENT-AMBIGUOUS";
        public const string WeakRowAlignment = "CMP-PROJECT-ROW-ALIGNMENT-WEAK";
        public const string WeakColumnAlignment = "CMP-PROJECT-COLUMN-ALIGNMENT-WEAK";
        public const string AxisIdentityIncomplete = "CMP-PROJECT-AXIS-IDENTITY-INCOMPLETE";
        public const string NonFormulaValueContentUnmodeled =
            "CMP-PROJECT-NON-FORMULA-VALUE-CONTENT-UNMODELED";
        public const string SiblingRowInsertionAmbiguous =
            "CMP-PROJECT-SIBLING-ROW-INSERTION-AMBIGUOUS";
        public const string SiblingColumnInsertionAmbiguous =
            "CMP-PROJECT-SIBLING-COLUMN-INSERTION-AMBIGUOUS";
        public const string IndependentSheetInsertionAmbiguous =
            "CMP-PROJECT-BRANCH-SHEET-INSERTION-AMBIGUOUS";
        public const string HistoryBuildFailed = "CMP-PROJECT-HISTORY-BUILD-FAILED";
    }

    /// <summary>A stable, actionable project-build diagnostic. No ambiguous match is hidden behind this type.</summary>
    public sealed class OpenXmlProjectFinding
    {
        internal OpenXmlProjectFinding(
            string findingId,
            string code,
            DiagnosticSeverity severity,
            string beforeRevisionId,
            string afterRevisionId,
            string logicalSheetId,
            string sheetName,
            string message,
            string suggestedAction,
            IEnumerable<string> reasonCodes,
            IEnumerable<string> candidateSheetNames)
        {
            FindingId = Require(findingId, nameof(findingId));
            Code = Require(code, nameof(code));
            Severity = severity;
            BeforeRevisionId = beforeRevisionId ?? "";
            AfterRevisionId = afterRevisionId ?? "";
            LogicalSheetId = logicalSheetId ?? "";
            SheetName = sheetName ?? "";
            Message = message ?? "";
            SuggestedAction = suggestedAction ?? "";
            ReasonCodes = Freeze(reasonCodes);
            CandidateSheetNames = Freeze(candidateSheetNames);
        }

        public string FindingId { get; }
        public string Code { get; }
        public DiagnosticSeverity Severity { get; }
        public string BeforeRevisionId { get; }
        public string AfterRevisionId { get; }
        public string LogicalSheetId { get; }
        public string SheetName { get; }
        public string Message { get; }
        public string SuggestedAction { get; }
        public IReadOnlyList<string> ReasonCodes { get; }
        public IReadOnlyList<string> CandidateSheetNames { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A value is required.", name);
            return value;
        }

        private static IReadOnlyList<string> Freeze(IEnumerable<string> values) =>
            new ReadOnlyCollection<string>(new List<string>(values ?? Array.Empty<string>()));
    }

    public sealed class OpenXmlProjectSheetAlignment
    {
        internal OpenXmlProjectSheetAlignment(
            string logicalSheetId,
            string beforeSheetKey,
            string afterSheetKey,
            string beforeSheetName,
            string afterSheetName,
            HeaderDetectionResult beforeHeader,
            HeaderDetectionResult afterHeader,
            KeyInferenceResult keyInference,
            AxisAlignmentResult rowAlignment,
            AxisAlignmentResult columnAlignment,
            LogicalAxisIdentityProjection rows,
            LogicalAxisIdentityProjection columns,
            bool isComplete)
        {
            LogicalSheetId = logicalSheetId ?? "";
            BeforeSheetKey = beforeSheetKey ?? "";
            AfterSheetKey = afterSheetKey ?? "";
            BeforeSheetName = beforeSheetName ?? "";
            AfterSheetName = afterSheetName ?? "";
            BeforeHeader = beforeHeader;
            AfterHeader = afterHeader;
            KeyInference = keyInference;
            RowAlignment = rowAlignment;
            ColumnAlignment = columnAlignment;
            Rows = rows;
            Columns = columns;
            IsComplete = isComplete;
        }

        public string LogicalSheetId { get; }
        public string BeforeSheetKey { get; }
        public string AfterSheetKey { get; }
        public string BeforeSheetName { get; }
        public string AfterSheetName { get; }
        public HeaderDetectionResult BeforeHeader { get; }
        public HeaderDetectionResult AfterHeader { get; }
        public KeyInferenceResult KeyInference { get; }
        public AxisAlignmentResult RowAlignment { get; }
        public AxisAlignmentResult ColumnAlignment { get; }
        public LogicalAxisIdentityProjection Rows { get; }
        public LogicalAxisIdentityProjection Columns { get; }
        public bool IsComplete { get; }
    }

    public sealed class OpenXmlAlignedSheetRevision
    {
        internal OpenXmlAlignedSheetRevision(
            string physicalSheetKey,
            string physicalSheetName,
            string logicalSheetId,
            CanonicalGridProjection grid,
            LogicalAxisIdentityProjection rows,
            LogicalAxisIdentityProjection columns)
        {
            PhysicalSheetKey = Require(physicalSheetKey, nameof(physicalSheetKey));
            PhysicalSheetName = physicalSheetName ?? "";
            LogicalSheetId = Require(logicalSheetId, nameof(logicalSheetId));
            Grid = grid ?? throw new ArgumentNullException(nameof(grid));
            Rows = rows ?? throw new ArgumentNullException(nameof(rows));
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
        }

        public string PhysicalSheetKey { get; }
        public string PhysicalSheetName { get; }
        public string LogicalSheetId { get; }
        public CanonicalGridProjection Grid { get; }
        public LogicalAxisIdentityProjection Rows { get; }
        public LogicalAxisIdentityProjection Columns { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A value is required.", name);
            return value;
        }
    }

    public sealed class OpenXmlAlignedProjectRevision
    {
        internal OpenXmlAlignedProjectRevision(
            OpenXmlProjectRevision source,
            IEnumerable<OpenXmlAlignedSheetRevision> sheets,
            RevisionSnapshot alignedSnapshot,
            bool isComplete)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Sheets = new ReadOnlyCollection<OpenXmlAlignedSheetRevision>(
                new List<OpenXmlAlignedSheetRevision>(sheets ?? Array.Empty<OpenXmlAlignedSheetRevision>()));
            AlignedSnapshot = alignedSnapshot;
            IsComplete = isComplete;
        }

        public OpenXmlProjectRevision Source { get; }
        public IReadOnlyList<OpenXmlAlignedSheetRevision> Sheets { get; }
        /// <summary>Null when any sheet or axis identity for this revision remains unresolved.</summary>
        public RevisionSnapshot AlignedSnapshot { get; }
        public bool IsComplete { get; }
    }

    public sealed class OpenXmlProjectEdgeResult
    {
        internal OpenXmlProjectEdgeResult(
            string parentRevisionId,
            string childRevisionId,
            OpenXmlWorkbookSheetMappingResult sheetMapping,
            IEnumerable<OpenXmlProjectSheetAlignment> sheetAlignments,
            bool isComplete)
        {
            ParentRevisionId = parentRevisionId ?? "";
            ChildRevisionId = childRevisionId ?? "";
            SheetMapping = sheetMapping ?? throw new ArgumentNullException(nameof(sheetMapping));
            SheetAlignments = new ReadOnlyCollection<OpenXmlProjectSheetAlignment>(
                new List<OpenXmlProjectSheetAlignment>(sheetAlignments ?? Array.Empty<OpenXmlProjectSheetAlignment>()));
            IsComplete = isComplete;
        }

        public string ParentRevisionId { get; }
        public string ChildRevisionId { get; }
        public OpenXmlWorkbookSheetMappingResult SheetMapping { get; }
        public IReadOnlyList<OpenXmlProjectSheetAlignment> SheetAlignments { get; }
        public bool IsComplete { get; }
    }

    public sealed class OpenXmlComparisonProjectResult
    {
        internal OpenXmlComparisonProjectResult(
            string projectId,
            string mappingGenerationId,
            string alignmentCorrectionGenerationId,
            string effectiveMappingGenerationId,
            IEnumerable<OpenXmlAlignedProjectRevision> revisions,
            IEnumerable<OpenXmlProjectEdgeResult> edges,
            IEnumerable<OpenXmlProjectFinding> findings,
            LogicalHistoryResult history,
            bool isComplete)
        {
            ProjectId = projectId ?? "";
            MappingGenerationId = mappingGenerationId ?? "";
            AlignmentCorrectionGenerationId = alignmentCorrectionGenerationId ?? "";
            EffectiveMappingGenerationId = effectiveMappingGenerationId ?? MappingGenerationId;
            Revisions = new ReadOnlyCollection<OpenXmlAlignedProjectRevision>(
                new List<OpenXmlAlignedProjectRevision>(revisions ?? Array.Empty<OpenXmlAlignedProjectRevision>()));
            Edges = new ReadOnlyCollection<OpenXmlProjectEdgeResult>(
                new List<OpenXmlProjectEdgeResult>(edges ?? Array.Empty<OpenXmlProjectEdgeResult>()));
            Findings = new ReadOnlyCollection<OpenXmlProjectFinding>(
                new List<OpenXmlProjectFinding>(findings ?? Array.Empty<OpenXmlProjectFinding>()));
            History = history;
            IsComplete = isComplete;
        }

        public string ProjectId { get; }
        /// <summary>The confirmed revision-lineage mapping generation supplied by the caller.</summary>
        public string MappingGenerationId { get; }
        public string AlignmentCorrectionGenerationId { get; }
        /// <summary>Stable identity over both lineage mapping and alignment confirmations.</summary>
        public string EffectiveMappingGenerationId { get; }
        public IReadOnlyList<OpenXmlAlignedProjectRevision> Revisions { get; }
        public IReadOnlyList<OpenXmlProjectEdgeResult> Edges { get; }
        public IReadOnlyList<OpenXmlProjectFinding> Findings { get; }
        /// <summary>
        /// Exact formula/value history for the full project or for the safely aligned diagnostic subset. A non-null
        /// value does not grant merge eligibility; callers must require <see cref="IsComplete"/> for that boundary.
        /// </summary>
        public LogicalHistoryResult History { get; }
        public bool IsComplete { get; }
    }
}
