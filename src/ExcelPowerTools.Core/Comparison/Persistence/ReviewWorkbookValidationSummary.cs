using System;
using System.IO;
using System.Text;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>
    /// Small immutable description of one successfully integrity-validated Review.xlsx. It intentionally exposes
    /// no embedded state text and performs no JSON parsing; source-artifact paths therefore cannot leak through
    /// this contract. FullPath identifies only the selected Review workbook itself.
    /// </summary>
    public sealed class ReviewWorkbookValidationSummary
    {
        internal ReviewWorkbookValidationSummary(
            string fullPath,
            string fileName,
            long fileByteLength,
            DateTimeOffset modifiedUtc,
            string projectId,
            string stateSha256Prefix,
            string projectionSha256Prefix,
            int stateChunkCount,
            int stateByteCount,
            bool hasDecisionSidecar,
            int decisionEvidenceRevisionCount,
            int decisionSubjectCount,
            int decisionObservationCount,
            int activeDecisionCount,
            string decisionCatalogIdPrefix,
            string decisionGenerationIdPrefix)
        {
            FullPath = fullPath ?? throw new ArgumentNullException(nameof(fullPath));
            FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
            FileByteLength = fileByteLength;
            ModifiedUtc = modifiedUtc;
            ProjectId = projectId ?? throw new ArgumentNullException(nameof(projectId));
            StateSha256Prefix = stateSha256Prefix ?? throw new ArgumentNullException(nameof(stateSha256Prefix));
            ProjectionSha256Prefix = projectionSha256Prefix ??
                                     throw new ArgumentNullException(nameof(projectionSha256Prefix));
            StateChunkCount = stateChunkCount;
            StateByteCount = stateByteCount;
            HasDecisionSidecar = hasDecisionSidecar;
            DecisionEvidenceRevisionCount = decisionEvidenceRevisionCount;
            DecisionSubjectCount = decisionSubjectCount;
            DecisionObservationCount = decisionObservationCount;
            ActiveDecisionCount = activeDecisionCount;
            DecisionCatalogIdPrefix = decisionCatalogIdPrefix ??
                                      throw new ArgumentNullException(nameof(decisionCatalogIdPrefix));
            DecisionGenerationIdPrefix = decisionGenerationIdPrefix ??
                                         throw new ArgumentNullException(nameof(decisionGenerationIdPrefix));
        }

        /// <summary>The selected Review workbook path; never an embedded input-artifact path.</summary>
        public string FullPath { get; }
        public string FileName { get; }
        public long FileByteLength { get; }
        public DateTimeOffset ModifiedUtc { get; }
        /// <summary>Raw only for a SHA-256 identity; other manifest text is represented by a bounded opaque ID.</summary>
        public string ProjectId { get; }
        public string StateSha256Prefix { get; }
        public string ProjectionSha256Prefix { get; }
        public int StateChunkCount { get; }
        public int StateByteCount { get; }
        /// <summary>True only when the package contained and validated an integrity-bound decision sidecar.</summary>
        public bool HasDecisionSidecar { get; }
        public int DecisionEvidenceRevisionCount { get; }
        public int DecisionSubjectCount { get; }
        public int DecisionObservationCount { get; }
        public int ActiveDecisionCount { get; }
        public string DecisionCatalogIdPrefix { get; }
        public string DecisionGenerationIdPrefix { get; }
    }

    public static class ReviewWorkbookValidationSummaryService
    {
        public const int HashPrefixLength = 12;
        private const string OpaqueProjectIdPrefix = "opaque:";

        /// <summary>
        /// Returns the privacy-safe portion of a complete protected-handle hydration. The public summary contract
        /// continues to expose no embedded StateJson, decision payloads, or source-artifact paths.
        /// </summary>
        public static ReviewWorkbookValidationSummary Validate(string workbookPath) =>
            ReviewWorkbookHydrationService.Hydrate(workbookPath).Summary;

        internal static ReviewWorkbookValidationSummary CreateValidatedSummary(
            string path,
            long fileBytes,
            DateTimeOffset modifiedUtc,
            ReviewWorkbookReadResult validation)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A full path is required.", nameof(path));
            if (validation == null) throw new ArgumentNullException(nameof(validation));
            int stateBytes = new UTF8Encoding(false, true).GetByteCount(validation.StateJson);
            bool hasDecisionSidecar = validation.HasDecisionState;
            int evidenceRevisionCount = 0;
            int subjectCount = 0;
            int observationCount = 0;
            int activeDecisionCount = 0;
            string catalogIdPrefix = "";
            string generationIdPrefix = "";
            if (hasDecisionSidecar)
            {
                var generation = validation.DecisionGeneration;
                if (generation.Catalog.EvidenceGeneration == null)
                    throw new InvalidDataException(
                        "The validated decision sidecar has no exact evidence generation.");
                evidenceRevisionCount = generation.Catalog.EvidenceGeneration.Revisions.Count;
                subjectCount = generation.Catalog.Subjects.Count;
                observationCount = generation.Catalog.Observations.Count;
                activeDecisionCount = generation.ActiveDecisions.Count;
                catalogIdPrefix = Prefix(generation.CatalogId);
                generationIdPrefix = Prefix(generation.GenerationId);
            }
            return new ReviewWorkbookValidationSummary(
                path,
                Path.GetFileName(path),
                fileBytes,
                modifiedUtc,
                PrivacySafeProjectId(validation.ProjectId),
                Prefix(validation.StateSha256),
                Prefix(validation.ProjectionSha256),
                validation.ChunkCount,
                stateBytes,
                hasDecisionSidecar,
                evidenceRevisionCount,
                subjectCount,
                observationCount,
                activeDecisionCount,
                catalogIdPrefix,
                generationIdPrefix);
        }

        private static string PrivacySafeProjectId(string projectId)
        {
            if (StableId.IsSha256(projectId)) return projectId;
            return OpaqueProjectIdPrefix + StableId.Compute(
                "review-summary-project-id-v1",
                projectId ?? "");
        }

        private static string Prefix(string sha256)
        {
            if (string.IsNullOrEmpty(sha256) || sha256.Length < HashPrefixLength)
                throw new InvalidDataException("The validated Review workbook returned an invalid SHA-256 identity.");
            return sha256.Substring(0, HashPrefixLength);
        }
    }
}
