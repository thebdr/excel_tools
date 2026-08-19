using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Comparison.Decisions;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>The durable state and representative character diff rendered into a review workbook.</summary>
    public sealed class ReviewWorkbookState
    {
        public ReviewWorkbookState(string projectId, string stateJson)
        {
            if (string.IsNullOrWhiteSpace(projectId))
                throw new ArgumentException("A project identifier is required.", nameof(projectId));
            ProjectId = projectId;
            StateJson = stateJson ?? throw new ArgumentNullException(nameof(stateJson));
            ComparisonStatus = "Comparison review";
            DetailBefore = "";
            DetailAfter = "";
            DetailUnavailableReason = "No formula or text change is available for character-level detail.";
            ChangeRows = new List<ReviewChangeProjectionRow>();
            RevisionRows = new List<ReviewRevisionProjectionRow>();
            DiagnosticRows = new List<ReviewDiagnosticProjectionRow>();
        }

        public string ProjectId { get; }
        public string StateJson { get; }
        public string ComparisonStatus { get; set; }
        public string DetailBefore { get; set; }
        public string DetailAfter { get; set; }
        public string DetailUnavailableReason { get; set; }
        /// <summary>
        /// Visible, non-authoritative projection rows. Exact evidence remains in <see cref="StateJson"/>; these
        /// rows are integrity-bound to it by the workbook projection digest.
        /// </summary>
        public IList<ReviewChangeProjectionRow> ChangeRows { get; }
        /// <summary>Revision-graph projection, including the user-requested original filesystem locator.</summary>
        public IList<ReviewRevisionProjectionRow> RevisionRows { get; }
        /// <summary>Visible portable diagnostics. The embedded state remains authoritative.</summary>
        public IList<ReviewDiagnosticProjectionRow> DiagnosticRows { get; }
        /// <summary>
        /// Optional immutable review-decision generation persisted beside the authoritative project state.
        /// Legacy and diagnostic Review.xlsx files omit it. The generation retains its exact evidence catalog;
        /// non-root generations persist their complete ordered event chain and are rebuilt by policy replay.
        /// </summary>
        public ReviewDecisionGeneration DecisionGeneration { get; set; }
        /// <summary>
        /// Independently supplied authoritative evidence generation for <see cref="DecisionGeneration"/>.
        /// Required when a decision sidecar is written; it prevents a same-project catalog for different input
        /// bytes from being attached accidentally.
        /// </summary>
        public string DecisionEvidenceGenerationId { get; set; } = "";
        /// <summary>Authoritative effective mapping generation independently bound to the project state.</summary>
        public string DecisionMappingGenerationId { get; set; } = "";
    }

    public sealed class ReviewChangeProjectionRow
    {
        public ReviewChangeProjectionRow(
            string changeId,
            string sheetName,
            string address,
            string aspect,
            string beforeText,
            string afterText)
            : this(
                changeId,
                sheetName,
                address,
                aspect,
                beforeText,
                afterText,
                "",
                "",
                2,
                2,
                "")
        {
        }

        public ReviewChangeProjectionRow(
            string changeId,
            string sheetName,
            string address,
            string aspect,
            string beforeText,
            string afterText,
            string groupId,
            string flags,
            int revisionCount,
            int distinctStateCount,
            string alignmentConfidence)
        {
            if (string.IsNullOrWhiteSpace(changeId))
                throw new ArgumentException("A projected change identifier is required.", nameof(changeId));
            if (revisionCount < 0) throw new ArgumentOutOfRangeException(nameof(revisionCount));
            if (distinctStateCount < 0) throw new ArgumentOutOfRangeException(nameof(distinctStateCount));
            ChangeId = changeId;
            SheetName = sheetName ?? "";
            Address = address ?? "";
            Aspect = aspect ?? "";
            BeforeText = beforeText ?? "";
            AfterText = afterText ?? "";
            GroupId = groupId ?? "";
            Flags = flags ?? "";
            RevisionCount = revisionCount;
            DistinctStateCount = distinctStateCount;
            AlignmentConfidence = alignmentConfidence ?? "";
        }

        public string ChangeId { get; }
        public string SheetName { get; }
        public string Address { get; }
        public string Aspect { get; }
        public string BeforeText { get; }
        public string AfterText { get; }
        public string GroupId { get; }
        public string Flags { get; }
        public int RevisionCount { get; }
        public int DistinctStateCount { get; }
        public string AlignmentConfidence { get; }
    }

    public sealed class ReviewRevisionProjectionRow
    {
        public ReviewRevisionProjectionRow(
            string revisionId,
            string label,
            int order,
            string parentRevisionId,
            string role,
            string sourceBaseName,
            string byteContentId,
            string modifiedUtc,
            string inference,
            string confidence)
            : this(
                revisionId,
                label,
                order,
                parentRevisionId,
                role,
                sourceBaseName,
                "",
                byteContentId,
                modifiedUtc,
                inference,
                confidence)
        {
        }

        public ReviewRevisionProjectionRow(
            string revisionId,
            string label,
            int order,
            string parentRevisionId,
            string role,
            string sourceBaseName,
            string sourcePath,
            string byteContentId,
            string modifiedUtc,
            string inference,
            string confidence)
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A projected revision identifier is required.", nameof(revisionId));
            if (order < 0) throw new ArgumentOutOfRangeException(nameof(order));
            RevisionId = revisionId;
            Label = label ?? "";
            Order = order;
            ParentRevisionId = parentRevisionId ?? "";
            Role = role ?? "";
            SourceBaseName = sourceBaseName ?? "";
            SourcePath = sourcePath ?? "";
            ByteContentId = byteContentId ?? "";
            ModifiedUtc = modifiedUtc ?? "";
            Inference = inference ?? "";
            Confidence = confidence ?? "";
        }

        public string RevisionId { get; }
        public string Label { get; }
        public int Order { get; }
        public string ParentRevisionId { get; }
        public string Role { get; }
        public string SourceBaseName { get; }
        public string SourcePath { get; }
        public string ByteContentId { get; }
        public string ModifiedUtc { get; }
        public string Inference { get; }
        public string Confidence { get; }
    }

    public sealed class ReviewDiagnosticProjectionRow
    {
        public ReviewDiagnosticProjectionRow(
            string findingUid,
            string code,
            string severity,
            string revisionId,
            string summary,
            string detail,
            string suggestedAction)
        {
            if (string.IsNullOrWhiteSpace(findingUid))
                throw new ArgumentException("A projected diagnostic UID is required.", nameof(findingUid));
            if (string.IsNullOrWhiteSpace(code))
                throw new ArgumentException("A projected diagnostic code is required.", nameof(code));
            FindingUid = findingUid;
            Code = code;
            Severity = severity ?? "";
            RevisionId = revisionId ?? "";
            Summary = summary ?? "";
            Detail = detail ?? "";
            SuggestedAction = suggestedAction ?? "";
        }

        public string FindingUid { get; }
        public string Code { get; }
        public string Severity { get; }
        public string RevisionId { get; }
        public string Summary { get; }
        public string Detail { get; }
        public string SuggestedAction { get; }
    }

    public sealed class ReviewWorkbookWriteOptions
    {
        public const int DefaultStateChunkSizeBytes = 1024 * 1024;

        public int StateChunkSizeBytes { get; set; } = DefaultStateChunkSizeBytes;
    }

    /// <summary>Validated state read from a review workbook, independent of Open XML SDK types.</summary>
    public sealed class ReviewWorkbookReadResult
    {
        internal ReviewWorkbookReadResult(
            string projectId,
            string stateJson,
            string manifestItemId,
            string stateSha256,
            string projectionSha256,
            IReadOnlyList<string> chunkItemIds,
            ReviewDecisionGeneration decisionGeneration = null,
            string decisionItemId = "",
            string decisionSha256 = "",
            ReviewWorkbookProjectionValidationMetrics projectionValidationMetrics = null)
        {
            ProjectId = projectId;
            StateJson = stateJson;
            ManifestItemId = manifestItemId;
            StateSha256 = stateSha256;
            ProjectionSha256 = projectionSha256;
            ChunkItemIds = chunkItemIds;
            DecisionGeneration = decisionGeneration;
            DecisionItemId = decisionItemId ?? "";
            DecisionSha256 = decisionSha256 ?? "";
            ProjectionValidationMetrics = projectionValidationMetrics;
        }

        public string ProjectId { get; }
        public string StateJson { get; }
        public string ManifestItemId { get; }
        public string StateSha256 { get; }
        public string ProjectionSha256 { get; }
        public IReadOnlyList<string> ChunkItemIds { get; }
        public int ChunkCount => ChunkItemIds.Count;
        /// <summary>
        /// Validated immutable decision generation reconstructed from the complete sidecar event chain, or null
        /// for a legacy or diagnostic workbook with no decision sidecar.
        /// </summary>
        public ReviewDecisionGeneration DecisionGeneration { get; }
        public bool HasDecisionState => DecisionGeneration != null;
        public string DecisionItemId { get; }
        public string DecisionSha256 { get; }
        /// <summary>Bounded deterministic work counters from the projection-integrity pass.</summary>
        public ReviewWorkbookProjectionValidationMetrics ProjectionValidationMetrics { get; }
    }
}
