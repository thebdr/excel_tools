using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Inventory;
using ExcelPowerTools.Core.Comparison.Lineage;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Transient result of inventorying a comparison-project selection. Absolute paths deliberately live only
    /// in the artifact and inventory objects returned to the current command; project and revision identifiers
    /// are one-way SHA-256 values and are safe to carry into persisted comparison state.
    /// </summary>
    internal sealed class ComparisonProjectPreparationResult
    {
        private readonly IReadOnlyList<ComparisonProjectArtifactInfo> _orderedArtifacts;
        private readonly IReadOnlyList<ArtifactInventoryRecord> _orderedInventoryRecords;

        internal ComparisonProjectPreparationResult(
            string projectId,
            IEnumerable<ComparisonProjectArtifactInfo> orderedArtifacts,
            IEnumerable<ArtifactInventoryRecord> orderedInventoryRecords,
            LineageInferenceResult inference,
            string suggestedOutputPath)
        {
            if (!StableId.IsSha256(projectId))
                throw new ArgumentException("A project SHA-256 identifier is required.", nameof(projectId));

            ProjectId = projectId;
            _orderedArtifacts = new ReadOnlyCollection<ComparisonProjectArtifactInfo>(
                new List<ComparisonProjectArtifactInfo>(
                    orderedArtifacts ?? throw new ArgumentNullException(nameof(orderedArtifacts))));
            _orderedInventoryRecords = new ReadOnlyCollection<ArtifactInventoryRecord>(
                new List<ArtifactInventoryRecord>(
                    orderedInventoryRecords ?? throw new ArgumentNullException(nameof(orderedInventoryRecords))));
            Inference = inference ?? throw new ArgumentNullException(nameof(inference));
            SuggestedOutputPath = suggestedOutputPath ?? throw new ArgumentNullException(nameof(suggestedOutputPath));

            if (_orderedArtifacts.Count != _orderedInventoryRecords.Count ||
                _orderedArtifacts.Count != Inference.Proposals.Count)
            {
                throw new ArgumentException(
                    "Prepared artifacts, inventory records, and lineage proposals must have matching counts.");
            }
        }

        internal string ProjectId { get; }

        /// <summary>Artifacts in exactly the same order as <see cref="Inference"/> proposals.</summary>
        internal IReadOnlyList<ComparisonProjectArtifactInfo> OrderedArtifacts => _orderedArtifacts;

        /// <summary>
        /// Exact immutable inventory observations corresponding by index to <see cref="OrderedArtifacts"/>.
        /// These records are command-lifetime data and must not be copied wholesale into persisted project state.
        /// </summary>
        internal IReadOnlyList<ArtifactInventoryRecord> OrderedInventoryRecords => _orderedInventoryRecords;

        internal LineageInferenceResult Inference { get; }
        internal string SuggestedOutputPath { get; }
    }

    /// <summary>
    /// Builds the immutable input boundary for a multi-revision comparison project. The current project engine
    /// accepts only bounded OOXML workbooks; unsupported families are rejected before any parser is selected.
    /// </summary>
    internal static class ComparisonProjectPreparationService
    {
        internal const int MinimumRevisionCount = 2;
        internal const int MaximumRevisionCount = 20;
        internal const long MaximumAggregateInputBytes = 768L * 1024 * 1024;

        internal static async Task<ComparisonProjectPreparationResult> PrepareAsync(
            IReadOnlyList<string> selectedPaths,
            IProgress<ComparisonPreviewProgress> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string[] normalizedPaths = NormalizeSelection(selectedPaths);

            progress?.Report(new ComparisonPreviewProgress(
                "Preparing comparison project",
                "Validating " + normalizedPaths.Length.ToString(CultureInfo.InvariantCulture) +
                " distinct local workbook selections",
                0,
                true));

            var inventoryService = new ArtifactInventoryService();
            var records = new List<ArtifactInventoryRecord>(normalizedPaths.Length);
            for (int index = 0; index < normalizedPaths.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string path = normalizedPaths[index];
                int startPercent = AggregatePercent(index, 0d, normalizedPaths.Length);
                progress?.Report(new ComparisonPreviewProgress(
                    "Inventorying revisions",
                    "Workbook " + (index + 1).ToString(CultureInfo.InvariantCulture) + " of " +
                    normalizedPaths.Length.ToString(CultureInfo.InvariantCulture) + ": " + Path.GetFileName(path),
                    startPercent,
                    true));

                ArtifactInventoryRecord record = await inventoryService.InventoryAsync(
                    path,
                    new ProjectInventoryProgressForwarder(progress, index, normalizedPaths.Length),
                    cancellationToken).ConfigureAwait(false);
                ValidateCurrentEngineInput(record);
                records.Add(record);
                long aggregateBytes = records.Sum(item => item.ByteLength);
                if (aggregateBytes > MaximumAggregateInputBytes)
                {
                    throw new InvalidDataException(
                        "The selected revision set totals " + FormatBytes(aggregateBytes) +
                        ", exceeding the current in-process comparison-project limit of " +
                        FormatBytes(MaximumAggregateInputBytes) + ". Split the set into smaller projects for now; " +
                        "the planned streaming/isolated engine will raise this limit without risking Excel stability.");
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonPreviewProgress(
                "Inferring revision lineage",
                "Using filename evidence conservatively; order and parent proposals remain user-correctable",
                91,
                true));

            string[] occurrenceTokens = records
                .Select(CreateOccurrenceToken)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
            string projectId = StableId.Compute("comparison-project-v1", occurrenceTokens);

            var artifactByRevisionId = new Dictionary<string, ComparisonProjectArtifactInfo>(StringComparer.Ordinal);
            var recordByRevisionId = new Dictionary<string, ArtifactInventoryRecord>(StringComparer.Ordinal);
            var evidence = new List<LineageRevisionEvidence>(records.Count);
            foreach (ArtifactInventoryRecord record in records)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string occurrenceToken = CreateOccurrenceToken(record);
                string revisionId = StableId.Compute(
                    "comparison-project-revision-occurrence-v1",
                    projectId,
                    occurrenceToken);

                var artifact = new ComparisonProjectArtifactInfo(
                    record.FullNormalizedPath,
                    revisionId,
                    record.ContentSha256,
                    record.ByteLength,
                    record.WeakCreationTimeUtc,
                    record.WeakLastWriteTimeUtc,
                    record.Format.EffectiveFormat.ToString().ToUpperInvariant());
                artifactByRevisionId.Add(revisionId, artifact);
                recordByRevisionId.Add(revisionId, record);
                evidence.Add(new LineageRevisionEvidence(
                    revisionId,
                    Path.GetFileName(record.FullNormalizedPath),
                    record.ContentSha256,
                    record.WeakCreationTimeUtc,
                    record.WeakLastWriteTimeUtc));
            }

            LineageInferenceResult inference = MultiRevisionLineageInference.Infer(evidence);
            var orderedArtifacts = new List<ComparisonProjectArtifactInfo>(records.Count);
            var orderedRecords = new List<ArtifactInventoryRecord>(records.Count);
            foreach (LineageRevisionProposal proposal in inference.Proposals)
            {
                orderedArtifacts.Add(artifactByRevisionId[proposal.RevisionId]);
                orderedRecords.Add(recordByRevisionId[proposal.RevisionId]);
            }

            cancellationToken.ThrowIfCancellationRequested();
            string suggestedOutputPath = SuggestNewOutputPath(orderedArtifacts);
            progress?.Report(new ComparisonPreviewProgress(
                "Comparison project ready",
                "Review the proposed revision order, parent links, master, and baseline before analysis",
                100,
                true));

            return new ComparisonProjectPreparationResult(
                projectId,
                orderedArtifacts,
                orderedRecords,
                inference,
                suggestedOutputPath);
        }

        private static string[] NormalizeSelection(IReadOnlyList<string> selectedPaths)
        {
            if (selectedPaths == null) throw new ArgumentNullException(nameof(selectedPaths));
            if (selectedPaths.Count < MinimumRevisionCount || selectedPaths.Count > MaximumRevisionCount)
            {
                throw new ArgumentException(
                    "Select between " + MinimumRevisionCount.ToString(CultureInfo.InvariantCulture) + " and " +
                    MaximumRevisionCount.ToString(CultureInfo.InvariantCulture) +
                    " distinct local workbook files for one comparison project.",
                    nameof(selectedPaths));
            }

            var normalized = new string[selectedPaths.Count];
            var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < selectedPaths.Count; index++)
            {
                string selected = selectedPaths[index];
                if (string.IsNullOrWhiteSpace(selected))
                {
                    throw new ArgumentException(
                        "Every comparison-project selection must be a non-blank local file path.",
                        nameof(selectedPaths));
                }

                try
                {
                    normalized[index] = Path.GetFullPath(selected);
                }
                catch (Exception exception) when (
                    exception is ArgumentException ||
                    exception is NotSupportedException ||
                    exception is PathTooLongException)
                {
                    throw new ArgumentException(
                        "A selected comparison-project path is invalid. Select existing local workbook files and retry.",
                        nameof(selectedPaths),
                        exception);
                }

                if (!distinct.Add(normalized[index]))
                {
                    throw new ArgumentException(
                        "The same local path was selected more than once: '" + Path.GetFileName(normalized[index]) +
                        "'. Select distinct paths. Byte-identical copies at different paths are supported.",
                        nameof(selectedPaths));
                }
            }

            return normalized;
        }

        private static void ValidateCurrentEngineInput(ArtifactInventoryRecord record)
        {
            ArtifactFormatAssessment format = record.Format;
            string name = Path.GetFileName(record.FullNormalizedPath);

            if (format.HasMismatch ||
                format.MatchKind == ArtifactFormatMatchKind.UnsupportedExtension ||
                format.MatchKind == ArtifactFormatMatchKind.UnrecognizedContent ||
                !format.IsSupportedForComparison)
            {
                throw new InvalidDataException(
                    "'" + name + "' was inventoried but cannot be accepted as a comparison-project source. " +
                    format.Message + " Correct the extension/content mismatch or save a clean .xlsx/.xlsm copy, " +
                    "then select that copy.");
            }

            ArtifactFormatKind effective = format.EffectiveFormat;
            if (effective == ArtifactFormatKind.Xlsb)
            {
                throw new NotSupportedException(
                    "'" + name + "' is XLSB. The current multi-revision alignment engine accepts .xlsx and .xlsm " +
                    "only. Save a comparison copy as .xlsx/.xlsm and retry; the original file is not modified.");
            }

            if (effective == ArtifactFormatKind.Csv ||
                effective == ArtifactFormatKind.Tsv ||
                effective == ArtifactFormatKind.DelimitedText)
            {
                throw new NotSupportedException(
                    "'" + name + "' is a separated-value file. The current multi-revision alignment engine " +
                    "accepts .xlsx and .xlsm only. Save/import the revisions as .xlsx, or use Exact two-file " +
                    "preview for a two-file CSV/TSV/TXT comparison.");
            }

            if (effective != ArtifactFormatKind.Xlsx && effective != ArtifactFormatKind.Xlsm)
            {
                throw new NotSupportedException(
                    "'" + name + "' is not a supported OOXML workbook. The current multi-revision alignment " +
                    "engine accepts .xlsx and .xlsm only.");
            }

            if (record.ByteLength > OpenXmlWorkbookSnapshotAdapter.MaxSourcePackageBytes)
            {
                throw new InvalidDataException(
                    "'" + name + "' is " + FormatBytes(record.ByteLength) +
                    ", exceeding the current OOXML adapter limit of " +
                    FormatBytes(OpenXmlWorkbookSnapshotAdapter.MaxSourcePackageBytes) +
                    ". Reduce the workbook or split the revision set, then retry. The file was inventoried but " +
                    "was not parsed.");
            }
        }

        private static string CreateOccurrenceToken(ArtifactInventoryRecord record)
        {
            // InventoryFingerprint is itself a SHA-256 over normalized path, metadata, and exact bytes. Including
            // the content address explicitly makes the occurrence boundary self-documenting without retaining a
            // reversible local path in either the project ID or revision ID.
            return StableId.Compute(
                "comparison-project-source-occurrence-v1",
                record.InventoryFingerprint,
                record.ContentSha256);
        }

        private static int AggregatePercent(int completedIndex, double currentFraction, int totalCount)
        {
            double bounded = Math.Max(0d, Math.Min(1d, currentFraction));
            double aggregate = (completedIndex + bounded) / totalCount;
            return (int)Math.Round(aggregate * 88d);
        }

        private static string SuggestNewOutputPath(IReadOnlyList<ComparisonProjectArtifactInfo> orderedArtifacts)
        {
            ComparisonProjectArtifactInfo latestProposal = orderedArtifacts[orderedArtifacts.Count - 1];
            string directory = Path.GetDirectoryName(latestProposal.Path);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                directory = Path.GetTempPath();

            string stem = Path.GetFileNameWithoutExtension(latestProposal.Path);
            if (string.IsNullOrWhiteSpace(stem)) stem = "spreadsheet";
            if (stem.Length > 96) stem = stem.Substring(0, 96).TrimEnd(' ', '.');
            stem += "-comparison-review";

            string candidate = Path.Combine(directory, stem + ".xlsx");
            for (int suffix = 2; File.Exists(candidate) || Directory.Exists(candidate); suffix++)
            {
                candidate = Path.Combine(
                    directory,
                    stem + "-" + suffix.ToString(CultureInfo.InvariantCulture) + ".xlsx");
            }

            return candidate;
        }

        private static string FormatBytes(long value)
        {
            string[] units = { "B", "KiB", "MiB", "GiB" };
            double display = value;
            int unit = 0;
            while (display >= 1024d && unit < units.Length - 1)
            {
                display /= 1024d;
                unit++;
            }

            return display.ToString(unit == 0 ? "N0" : "N1", CultureInfo.InvariantCulture) + " " + units[unit];
        }

        private sealed class ProjectInventoryProgressForwarder : IProgress<ArtifactInventoryProgress>
        {
            private readonly IProgress<ComparisonPreviewProgress> _outer;
            private readonly int _index;
            private readonly int _count;

            internal ProjectInventoryProgressForwarder(
                IProgress<ComparisonPreviewProgress> outer,
                int index,
                int count)
            {
                _outer = outer;
                _index = index;
                _count = count;
            }

            public void Report(ArtifactInventoryProgress value)
            {
                if (_outer == null || value == null) return;
                _outer.Report(new ComparisonPreviewProgress(
                    "Inventorying revisions",
                    "Workbook " + (_index + 1).ToString(CultureInfo.InvariantCulture) + " of " +
                    _count.ToString(CultureInfo.InvariantCulture) + ": " +
                    value.Stage + " - " + Path.GetFileName(value.FullNormalizedPath) +
                    (value.TotalBytes <= 0
                        ? ""
                        : "  " + FormatBytes(value.BytesRead) + " / " + FormatBytes(value.TotalBytes)),
                    AggregatePercent(_index, value.Fraction, _count),
                    true));
            }
        }
    }
}
