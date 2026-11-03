using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Lineage;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// User-facing entry point for the first multi-revision comparison-project workflow. The command keeps all
    /// source access outside Excel, requires explicit lineage confirmation, and publishes one new validated
    /// Review.xlsx only after alignment is complete or the user explicitly requests an incomplete diagnostic.
    /// </summary>
    public static class ComparisonProjectCommands
    {
        private const string InputFilter =
            "Comparison project workbooks|*.xlsx;*.xlsm|" +
            "Excel workbooks (.xlsx, .xlsm)|*.xlsx;*.xlsm|All files|*.*";

        [ExcelCommand(
            Name = "PT_CompareRevisionProject",
            MenuName = "Power Tools",
            MenuText = "Compare revisions...")]
        public static void StartComparisonProject()
        {
            string[] selected = Dialogs.OpenFiles(InputFilter);
            if (selected == null || selected.Length == 0) return;
            if (selected.Length < ComparisonProjectPreparationService.MinimumRevisionCount)
            {
                Dialogs.Info(
                    "A comparison project needs at least two revision files.\n\n" +
                    "Select every available revision together; you can correct their order and parent links " +
                    "before the files are compared.");
                return;
            }

            ComparisonProjectPreparationResult preparation;
            try
            {
                preparation = ComparisonBuildProgressDialog.Run(
                    ExcelWindow.Current,
                    new ComparisonBuildProgressDialogOptions
                    {
                        Title = "Prepare comparison project",
                        Heading = "Inventorying revision files",
                        InitialPhase = "Preparing",
                        InitialDetail = "Hashing and identifying the selected files without opening them in Excel.",
                    },
                    async (progress, cancellationToken) => await ComparisonProjectPreparationService.PrepareAsync(
                        selected,
                        new PreviewProgressAdapter(progress),
                        cancellationToken).ConfigureAwait(false));
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception exception)
            {
                ShowFailure("The revision files could not be prepared. Nothing was changed.", exception);
                return;
            }

            ComparisonProjectSetupResult setup;
            using (var form = new ComparisonProjectSetupForm(
                       preparation.OrderedArtifacts,
                       preparation.Inference,
                       preparation.SuggestedOutputPath))
            {
                if (form.ShowDialog(ExcelWindow.Current) != DialogResult.OK || form.Result == null) return;
                setup = form.Result;
            }

            ComparisonProjectAnalysisResult analysis;
            ComparisonProjectBuildResult result;
            try
            {
                long totalBytes = setup.OrderedArtifacts.Sum(item => item.ByteLength);
                string workload = ComparisonWorkloadBucket.Create(
                    setup.OrderedArtifacts.Count,
                    totalBytes,
                    totalSheetCount: 0,
                    approximateUsedCellCount: 0);
                analysis = ComparisonBuildProgressDialog.Run(
                    ExcelWindow.Current,
                    new ComparisonBuildProgressDialogOptions
                    {
                        Title = "Analyze comparison project",
                        Heading = "Aligning and comparing revisions",
                        InitialPhase = "Reading immutable inputs",
                        InitialDetail = "Confirming that every source still matches its inventory hash.",
                        WorkloadBucket = workload,
                    },
                    async (progress, cancellationToken) =>
                    {
                        ComparisonProjectBuildSession session = await LoadInputsAsync(
                            preparation,
                            setup,
                            progress,
                            cancellationToken).ConfigureAwait(false);
                        AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create(
                            setup.Mapping.GenerationId);
                        return Analyze(session, corrections, progress, cancellationToken);
                    });

                while (!analysis.ProjectResult.IsComplete ||
                       analysis.AlignmentCorrections.SheetPairs.Count > 0 ||
                       analysis.AlignmentCorrections.UnmatchedSheets.Count > 0 ||
                       analysis.AlignmentCorrections.HeaderRows.Count > 0 ||
                       analysis.AlignmentCorrections.KeyColumns.Count > 0)
                {
                    ComparisonAlignmentReviewResult review;
                    using (var form = new ComparisonAlignmentReviewForm(analysis))
                    {
                        form.ShowDialog(ExcelWindow.Current);
                        review = form.Result;
                    }

                    if (review == null || review.Action == ComparisonAlignmentReviewAction.Cancel) return;
                    bool publishAfterRerun = review.Action == ComparisonAlignmentReviewAction.PublishDiagnostic;
                    bool publishCompleteAfterRerun = review.Action == ComparisonAlignmentReviewAction.Publish;
                    if (!publishAfterRerun && !publishCompleteAfterRerun &&
                        review.Action != ComparisonAlignmentReviewAction.Reanalyze)
                        throw new InvalidOperationException("The alignment review returned an unknown action.");
                    if (publishCompleteAfterRerun && !analysis.ProjectResult.IsComplete)
                        throw new InvalidOperationException(
                            "Only a complete analysis can continue to the final Review workbook.");
                    if (review.Corrections == null)
                        throw new InvalidOperationException("The alignment review returned no correction generation.");

                    bool correctionsChanged = !string.Equals(
                        review.Corrections.GenerationId,
                        analysis.AlignmentCorrections.GenerationId,
                        StringComparison.Ordinal);
                    if (!correctionsChanged)
                    {
                        if (publishAfterRerun || publishCompleteAfterRerun) break;
                        throw new InvalidOperationException(
                            "Reanalysis requires at least one new or replaced alignment confirmation.");
                    }

                    AlignmentCorrectionGeneration rerunCorrections = review.Corrections;
                    ComparisonProjectBuildSession retainedSession = analysis.Session;
                    analysis = ComparisonBuildProgressDialog.Run(
                        ExcelWindow.Current,
                        new ComparisonBuildProgressDialogOptions
                        {
                            Title = "Reanalyze comparison project",
                            Heading = "Applying confirmed alignment",
                            InitialPhase = "Validating confirmations",
                            InitialDetail = "Rebuilding logical identities from the retained immutable snapshots.",
                        },
                        (progress, cancellationToken) => Task.FromResult(
                            Analyze(retainedSession, rerunCorrections, progress, cancellationToken)));
                    if (publishAfterRerun) break;
                    if (publishCompleteAfterRerun && analysis.ProjectResult.IsComplete) break;
                }

                result = ComparisonBuildProgressDialog.Run(
                    ExcelWindow.Current,
                    new ComparisonBuildProgressDialogOptions
                    {
                        Title = "Publish comparison project",
                        Heading = analysis.ProjectResult.IsComplete
                            ? "Creating the comparison review"
                            : "Creating a diagnostic review",
                        InitialPhase = "Preparing Review.xlsx",
                        InitialDetail = "Rechecking the new destination before atomic publication.",
                    },
                    (progress, cancellationToken) => Task.FromResult(
                        Publish(analysis, progress, cancellationToken)));
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception exception)
            {
                ShowFailure(
                    "The comparison project could not be completed safely. Source files were not changed. " +
                    "If a Review.xlsx exists at the chosen destination, leave it closed and inspect this error before retrying.",
                    exception);
                return;
            }

            try
            {
                ComparisonCommands.OpenValidatedReviewInExcel(result.OutputPath);
            }
            catch (Exception exception)
            {
                Dialogs.Info(
                    "The Review.xlsx was created and verified, but Excel could not open it.\n\n" +
                    result.OutputPath + "\n\n" + ComparisonCommands.DescribeException(exception));
                return;
            }

            string usefulRangeNotice = BuildUsefulRangeNotice(result.InputBounds);
            if (!result.IsComplete)
            {
                Dialogs.Info(
                    "The sources were inventoried and a diagnostic Review.xlsx was created, but logical history " +
                    "was not guessed because " + result.FindingCount.ToString("N0", CultureInfo.CurrentCulture) +
                    " mapping or coverage finding(s) need attention.\n\n" +
                    "Open the Diagnostics sheet for stable finding IDs, full evidence, and corrective actions." +
                    usefulRangeNotice);
            }
            else if (!string.IsNullOrEmpty(usefulRangeNotice))
                Dialogs.Info("The comparison Review.xlsx was created successfully." + usefulRangeNotice);
        }

        /// <summary>
        /// Establishes the immutable comparison input boundary. Every source is read, content-hash verified,
        /// parsed, and bounded exactly once; later analysis passes must use the retained snapshots rather than
        /// returning to potentially changing source paths.
        /// </summary>
        internal static async Task<ComparisonProjectBuildSession> LoadInputsAsync(
            ComparisonProjectPreparationResult preparation,
            ComparisonProjectSetupResult setup,
            IProgress<ComparisonBuildProgressUpdate> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (preparation == null) throw new ArgumentNullException(nameof(preparation));
            if (setup == null) throw new ArgumentNullException(nameof(setup));
            ValidateSetupBoundary(preparation, setup);

            var snapshots = new Dictionary<string, OpenXmlWorkbookSnapshot>(StringComparer.Ordinal);
            var sheetBounds = new List<ComparisonInputSheetBounds>();
            int count = setup.OrderedArtifacts.Count;
            long aggregateStoredCells = 0;
            long aggregateRetainedStoredCells = 0;
            long aggregateGridCells = 0;
            for (int index = 0; index < count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ComparisonProjectArtifactInfo artifact = setup.OrderedArtifacts[index];
                int readStart = Scale(index, count, 2, 35);
                int readEnd = Scale(index + 1, count, 2, 35);
                progress?.Report(new ComparisonBuildProgressUpdate(
                    "Reading immutable inputs",
                    "Revision " + (index + 1).ToString(CultureInfo.InvariantCulture) + " of " +
                    count.ToString(CultureInfo.InvariantCulture) + ": " + artifact.BaseName,
                    readStart,
                    true));

                byte[] bytes = await ReadAndVerifyAsync(
                    artifact,
                    readStart,
                    readEnd,
                    progress,
                    cancellationToken).ConfigureAwait(false);

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new ComparisonBuildProgressUpdate(
                    "Parsing comparison content",
                    "Reading stored values, saved calculated values, and formulas from " + artifact.BaseName,
                    Scale(index, count, 36, 57),
                    true));
                OpenXmlWorkbookSnapshot parsedSnapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                    artifact.RevisionId,
                    bytes,
                    cancellationToken);
                OpenXmlWorkbookCompactionResult compaction =
                    OpenXmlWorkbookSnapshotCompactor.CompactToUsefulRanges(
                        parsedSnapshot,
                        cancellationToken);
                OpenXmlWorkbookSnapshot snapshot = compaction.Snapshot;
                IReadOnlyList<ComparisonInputSheetBounds> workbookBounds = BuildInputBounds(
                    artifact,
                    compaction.Worksheets);
                sheetBounds.AddRange(workbookBounds);
                long workbookStoredCells = workbookBounds.Sum(
                    item => item.Range.TotalStoredCellCount);
                long workbookRetainedStoredCells = workbookBounds.Sum(
                    item => item.Range.RetainedStoredCellCount);
                long workbookGridCells = workbookBounds.Sum(item => item.Range.GridCellCount);
                aggregateStoredCells = checked(aggregateStoredCells + workbookStoredCells);
                aggregateRetainedStoredCells = checked(
                    aggregateRetainedStoredCells + workbookRetainedStoredCells);
                aggregateGridCells = checked(aggregateGridCells + workbookGridCells);
                ComparisonAlignmentEnvelopeAssessment envelope =
                    ComparisonAlignmentEnvelopePolicy.Evaluate(
                        workbookStoredCells,
                        aggregateStoredCells,
                        aggregateRetainedStoredCells,
                        aggregateGridCells);
                if (!envelope.CanAnalyze)
                {
                    throw new InvalidDataException(BuildAlignmentEnvelopeError(
                        preparation.ProjectId,
                        artifact,
                        sheetBounds,
                        aggregateStoredCells,
                        aggregateGridCells,
                        envelope.ExceededLimit));
                }
                snapshots.Add(artifact.RevisionId, snapshot);
                long ignoredOutsideUsefulRange = workbookBounds.Sum(
                    item => item.Range.IgnoredOutsideRangeCellCount);
                progress?.Report(new ComparisonBuildProgressUpdate(
                    "Parsing comparison content",
                    "Parsed " + artifact.BaseName + " (" + snapshot.Sheets.Count.ToString(
                        CultureInfo.InvariantCulture) + " sheet(s); " +
                    workbookBounds.Sum(item => item.Range.GridCellCount).ToString(
                        "N0", CultureInfo.InvariantCulture) + " effective grid cells" +
                    (ignoredOutsideUsefulRange == 0
                        ? ""
                        : "; ignored " + ignoredOutsideUsefulRange.ToString(
                            "N0", CultureInfo.InvariantCulture) +
                          " physically stored blank cell(s) outside comparison-content bounds") + ")",
                    Scale(index + 1, count, 36, 57),
                    true));
            }

            cancellationToken.ThrowIfCancellationRequested();
            var projectRevisions = new List<OpenXmlProjectRevision>(count);
            for (int order = 0; order < setup.Mapping.OrderedRevisionIds.Count; order++)
            {
                string revisionId = setup.Mapping.OrderedRevisionIds[order];
                projectRevisions.Add(new OpenXmlProjectRevision(
                    revisionId,
                    setup.Mapping.GetParent(revisionId) ?? "",
                    order,
                    snapshots[revisionId],
                    string.Equals(revisionId, setup.Mapping.MasterRevisionId, StringComparison.Ordinal),
                    string.Equals(
                        revisionId,
                        setup.Mapping.TrustedBaselineRevisionId,
                        StringComparison.Ordinal)));
            }

            IReadOnlyList<ReviewRevisionDisplayMetadata> revisionMetadata = BuildRevisionMetadata(
                preparation,
                setup);
            return new ComparisonProjectBuildSession(
                preparation,
                setup,
                snapshots,
                projectRevisions,
                revisionMetadata,
                new ComparisonInputBoundsSummary(sheetBounds));
        }

        /// <summary>
        /// Runs the deterministic project engine over a retained immutable input session. This phase is safe to
        /// rerun when a future correction generation changes logical mapping; it performs no source I/O and does
        /// not create or modify the destination workbook.
        /// </summary>
        internal static ComparisonProjectAnalysisResult Analyze(
            ComparisonProjectBuildSession session,
            IProgress<ComparisonBuildProgressUpdate> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (session == null) throw new ArgumentNullException(nameof(session));
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create(
                session.Setup.Mapping.GenerationId);
            return Analyze(session, corrections, progress, cancellationToken);
        }

        /// <summary>
        /// Runs analysis with one immutable alignment-correction generation. The source evidence remains the same
        /// session snapshot set; only explicit logical confirmations are allowed to affect the engine pass.
        /// </summary>
        internal static ComparisonProjectAnalysisResult Analyze(
            ComparisonProjectBuildSession session,
            AlignmentCorrectionGeneration corrections,
            IProgress<ComparisonBuildProgressUpdate> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (corrections == null) throw new ArgumentNullException(nameof(corrections));

            cancellationToken.ThrowIfCancellationRequested();
            var projectOptions = new OpenXmlComparisonProjectOptions
            {
                AlignmentCorrections = corrections,
            };
            if (corrections.SheetPairs.Count > 0 || corrections.UnmatchedSheets.Count > 0)
            {
                progress?.Report(new ComparisonBuildProgressUpdate(
                    "Reconciling confirmed alignment",
                    "Checking exact pair-scoped key choices against confirmed sheet pairs and deleted/new choices.",
                    58,
                    true));
                AlignmentCorrectionReconciliationResult reconciliation =
                    AlignmentCorrectionReconciler.Reconcile(
                        session.ProjectRevisions,
                        corrections,
                        projectOptions.SheetMapping,
                        new ReviewerSnapshot(
                            identitySourceLabels: new[] { "power-tools-safety-reconciliation" }),
                        DateTimeOffset.UtcNow,
                        cancellationToken);
                corrections = reconciliation.Generation;
                projectOptions.AlignmentCorrections = corrections;
                if (reconciliation.Actions.Count > 0)
                    progress?.Report(new ComparisonBuildProgressUpdate(
                        "Reconciled dependent alignment",
                        reconciliation.Actions.Count.ToString(CultureInfo.InvariantCulture) +
                        " displaced key choice(s) were returned to automatic inference with audit events.",
                        59,
                        true));
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonBuildProgressUpdate(
                "Aligning revision history",
                "Mapping sheets, detecting headers and keys, and propagating logical row and column identities.",
                60,
                true));

            OpenXmlComparisonProjectResult projectResult = OpenXmlComparisonProjectEngine.Build(
                session.Preparation.ProjectId,
                session.Setup.Mapping.GenerationId,
                session.ProjectRevisions,
                projectOptions,
                cancellationToken: cancellationToken,
                progress: progress == null ? null : new ProjectEngineProgressAdapter(progress));

            progress?.Report(new ComparisonBuildProgressUpdate(
                "Building the review ledger",
                projectResult.IsComplete
                    ? "Collapsing adjacent changes into multi-revision histories and repeated-change groups."
                    : "Preserving blocking alignment evidence and corrective actions without guessing.",
                85,
                true));

            return new ComparisonProjectAnalysisResult(session, corrections, projectResult);
        }

        /// <summary>
        /// Materializes and validates the self-contained review artifact. The destination boundary is checked
        /// again on entry so time spent in analysis cannot turn a formerly free name into an overwrite.
        /// </summary>
        internal static ComparisonProjectBuildResult Publish(
            ComparisonProjectAnalysisResult analysis,
            IProgress<ComparisonBuildProgressUpdate> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (analysis == null) throw new ArgumentNullException(nameof(analysis));

            ComparisonProjectBuildSession session = analysis.Session;
            ComparisonProjectPreparationResult preparation = session.Preparation;
            ComparisonProjectSetupResult setup = session.Setup;
            OpenXmlComparisonProjectResult projectResult = analysis.ProjectResult;

            // This is intentionally the first publication-phase operation. ReviewWorkbookFile.CreateNew also
            // enforces new-only atomic creation, but the explicit recheck produces the actionable setup error.
            ValidateDestinationBoundary(setup.OutputPath);

            string stateJson = ComparisonProjectStateJson.Serialize(
                preparation,
                setup,
                analysis.AlignmentCorrections,
                projectResult);
            ReviewWorkbookState reviewState = OpenXmlProjectReviewProjector.Project(
                projectResult,
                stateJson,
                session.RevisionMetadata);
            ReviewDecisionGeneration expectedDecisionGeneration = null;
            if (projectResult.IsComplete && projectResult.History != null)
            {
                ReviewDecisionCatalog decisionCatalog = ReviewDecisionCatalogProjector.Project(
                    projectResult,
                    preparation.OrderedArtifacts.Select(artifact =>
                        new ReviewDecisionRevisionEvidence(
                            artifact.RevisionId,
                            artifact.ByteContentId,
                            artifact.BaseName)),
                    cancellationToken);
                expectedDecisionGeneration = ReviewDecisionEngine.Create(decisionCatalog);
                reviewState.DecisionGeneration = expectedDecisionGeneration;
                reviewState.DecisionEvidenceGenerationId = decisionCatalog.EvidenceGenerationId;
                reviewState.DecisionMappingGenerationId = projectResult.EffectiveMappingGenerationId;
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonBuildProgressUpdate(
                "Publishing Review.xlsx",
                "Writing an integrity-bound workbook. Cancellation is paused during atomic publication.",
                93,
                false));
            ReviewWorkbookFile.CreateNew(setup.OutputPath, reviewState);

            progress?.Report(new ComparisonBuildProgressUpdate(
                "Verifying Review.xlsx",
                "Reopening the closed package and checking state and projection integrity.",
                98,
                false));
            ReviewWorkbookReadResult reopened = ReviewWorkbookFile.ReadAndValidate(setup.OutputPath);
            if (!string.Equals(reopened.ProjectId, preparation.ProjectId, StringComparison.Ordinal) ||
                !string.Equals(reopened.StateJson, stateJson, StringComparison.Ordinal) ||
                (expectedDecisionGeneration == null && reopened.HasDecisionState) ||
                (expectedDecisionGeneration != null &&
                 (!reopened.HasDecisionState ||
                  !string.Equals(
                      reopened.DecisionGeneration.CatalogId,
                      expectedDecisionGeneration.CatalogId,
                      StringComparison.Ordinal) ||
                  !string.Equals(
                      reopened.DecisionGeneration.GenerationId,
                      expectedDecisionGeneration.GenerationId,
                      StringComparison.Ordinal))))
            {
                // Do not delete here. Another process could replace the newly published path between atomic
                // publication and this read-back; deleting on a failed ownership check could delete that file.
                throw new InvalidDataException(
                    "The closed Review.xlsx does not contain the exact comparison-project state that was " +
                    "published. The destination was left untouched for forensic inspection.");
            }

            progress?.Report(new ComparisonBuildProgressUpdate(
                "Complete",
                projectResult.IsComplete
                    ? projectResult.History.AtomicTransitionCount.ToString("N0", CultureInfo.InvariantCulture) +
                      " exact transition(s) across " + projectResult.History.Subjects.Count.ToString(
                          "N0", CultureInfo.InvariantCulture) + " logical finding(s)." +
                      BuildProgressUsefulRangeSuffix(session.InputBounds)
                    : projectResult.Findings.Count.ToString("N0", CultureInfo.InvariantCulture) +
                      " blocking/correctable finding(s) were preserved in Diagnostics." +
                      BuildProgressUsefulRangeSuffix(session.InputBounds),
                100,
                false));

            return new ComparisonProjectBuildResult(
                setup.OutputPath,
                preparation.ProjectId,
                projectResult.IsComplete,
                projectResult.History?.Subjects.Count ?? 0,
                projectResult.History?.AtomicTransitionCount ?? 0,
                projectResult.Findings.Count,
                session.InputBounds);
        }

        private static IReadOnlyList<ReviewRevisionDisplayMetadata> BuildRevisionMetadata(
            ComparisonProjectPreparationResult preparation,
            ComparisonProjectSetupResult setup)
        {
            var artifactById = preparation.OrderedArtifacts.ToDictionary(
                item => item.RevisionId,
                StringComparer.Ordinal);
            var proposalsById = preparation.Inference.Proposals.ToDictionary(
                item => item.RevisionId,
                StringComparer.Ordinal);
            var result = new List<ReviewRevisionDisplayMetadata>(setup.Mapping.OrderedRevisionIds.Count);
            for (int order = 0; order < setup.Mapping.OrderedRevisionIds.Count; order++)
            {
                string revisionId = setup.Mapping.OrderedRevisionIds[order];
                ComparisonProjectArtifactInfo artifact = artifactById[revisionId];
                LineageRevisionProposal proposal = proposalsById[revisionId];
                string inference = string.Join(" ", proposal.OrderReasons
                    .Concat(proposal.ParentReasons)
                    .Select(item => item.Message)
                    .Where(item => !string.IsNullOrWhiteSpace(item)));
                string confidence = "Order " + proposal.OrderConfidence.Level + " " +
                                    proposal.OrderConfidence.Score.ToString("0.00", CultureInfo.InvariantCulture) +
                                    "; parent " + proposal.ParentConfidence.Level + " " +
                                    proposal.ParentConfidence.Score.ToString("0.00", CultureInfo.InvariantCulture);
                result.Add(new ReviewRevisionDisplayMetadata(
                    revisionId,
                    setup.Mapping.GetLabel(revisionId),
                    order,
                    setup.Mapping.GetParent(revisionId) ?? "",
                    artifact.Path,
                    artifact.ByteContentId,
                    artifact.ModifiedUtc,
                    inference,
                    confidence));
            }
            return result.AsReadOnly();
        }

        private static async Task<byte[]> ReadAndVerifyAsync(
            ComparisonProjectArtifactInfo artifact,
            int startPercent,
            int endPercent,
            IProgress<ComparisonBuildProgressUpdate> progress,
            CancellationToken cancellationToken)
        {
            if (artifact.ByteLength < 0 || artifact.ByteLength > OpenXmlWorkbookSnapshotAdapter.MaxSourcePackageBytes)
                throw new InvalidDataException("The source size is outside the current OOXML adapter envelope.");
            if (artifact.ByteLength > int.MaxValue)
                throw new InvalidDataException("The source is too large for the current in-memory comparison path.");

            var bytes = new byte[(int)artifact.ByteLength];
            int offset = 0;
            using (var stream = new FileStream(
                       artifact.Path,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete,
                       128 * 1024,
                       useAsync: true))
            {
                if (stream.Length != artifact.ByteLength)
                    throw new IOException(
                        "The source size changed after inventory: " + artifact.BaseName +
                        ". Re-select the revision files and retry.");
                while (offset < bytes.Length)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = await stream.ReadAsync(
                        bytes,
                        offset,
                        Math.Min(128 * 1024, bytes.Length - offset),
                        cancellationToken).ConfigureAwait(false);
                    if (read == 0)
                        throw new EndOfStreamException("The source ended while it was being verified: " + artifact.BaseName);
                    offset += read;
                    if (progress != null && (offset == bytes.Length || offset % (4 * 1024 * 1024) < read))
                    {
                        int percent = bytes.Length == 0
                            ? endPercent
                            : startPercent + (int)Math.Round(
                                (endPercent - startPercent) * (offset / (double)bytes.Length));
                        progress.Report(new ComparisonBuildProgressUpdate(
                            "Reading immutable inputs",
                            artifact.BaseName + "  " + FormatBytes(offset) + " / " + FormatBytes(bytes.Length),
                            percent,
                            true));
                    }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            string actualHash = StableId.ComputeByteContentId(bytes);
            if (!string.Equals(actualHash, artifact.ByteContentId, StringComparison.OrdinalIgnoreCase))
                throw new IOException(
                    "The exact bytes changed after inventory: " + artifact.BaseName +
                    ". Re-select the revisions so the comparison cannot mix file states.");
            return bytes;
        }

        private static void ValidateSetupBoundary(
            ComparisonProjectPreparationResult preparation,
            ComparisonProjectSetupResult setup)
        {
            string[] prepared = preparation.OrderedArtifacts.Select(item => item.RevisionId)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToArray();
            string[] confirmed = setup.OrderedArtifacts.Select(item => item.RevisionId)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToArray();
            if (!prepared.SequenceEqual(confirmed, StringComparer.Ordinal) ||
                !prepared.SequenceEqual(
                    setup.Mapping.OrderedRevisionIds.OrderBy(item => item, StringComparer.Ordinal),
                    StringComparer.Ordinal))
                throw new InvalidOperationException(
                    "The confirmed lineage no longer describes the prepared revision occurrence set.");
            if (string.IsNullOrEmpty(setup.Mapping.MasterRevisionId))
                throw new InvalidOperationException("The comparison project has no explicitly confirmed master revision.");
            ValidateDestinationBoundary(setup.OutputPath);
        }

        private static void ValidateDestinationBoundary(string outputPath)
        {
            if (File.Exists(outputPath) || Directory.Exists(outputPath))
                throw new IOException("The Review.xlsx destination now exists. Choose a new output filename and retry.");
        }

        private static int Scale(int completed, int total, int start, int end)
        {
            if (total <= 0) return start;
            return start + (int)Math.Round((end - start) * (completed / (double)total));
        }

        private static IReadOnlyList<ComparisonInputSheetBounds> BuildInputBounds(
            ComparisonProjectArtifactInfo artifact,
            IReadOnlyList<OpenXmlWorksheetCompaction> worksheets)
        {
            var result = new List<ComparisonInputSheetBounds>(worksheets.Count);
            foreach (OpenXmlWorksheetCompaction sheet in worksheets)
            {
                result.Add(new ComparisonInputSheetBounds(
                    artifact,
                    sheet.SheetName,
                    sheet.Range));
            }
            return new ReadOnlyCollection<ComparisonInputSheetBounds>(result);
        }

        private static string BuildAlignmentEnvelopeError(
            string projectId,
            ComparisonProjectArtifactInfo crossingArtifact,
            IReadOnlyList<ComparisonInputSheetBounds> sheetBounds,
            long aggregateStoredCells,
            long aggregateGridCells,
            ComparisonAlignmentEnvelopeLimit exceededLimit)
        {
            long retainedStoredCells = sheetBounds.Sum(item => item.Range.RetainedStoredCellCount);
            long contentCells = sheetBounds.Sum(item => item.Range.ContentCellCount);
            long ignoredOutsideRange = sheetBounds.Sum(item => item.Range.IgnoredOutsideRangeCellCount);
            long invalidCoordinates = sheetBounds.Sum(item => item.Range.InvalidCoordinateCellCount);
            string projectToken = string.IsNullOrEmpty(projectId)
                ? "UNKNOWN"
                : projectId.Substring(0, Math.Min(12, projectId.Length));
            var message = new StringBuilder(4096);
            message.Append("Issue ID: COMPARE-ENVELOPE-").AppendLine(projectToken)
                .AppendLine("Source: ComparisonProjectCommands.LoadInputsAsync / content-aware-envelope-v3")
                .AppendLine()
                .AppendLine("The comparison is too large for safe in-memory alignment.")
                .Append("Limit reached: ").AppendLine(DescribeAlignmentEnvelopeLimit(exceededLimit))
                .AppendLine("The revision below is where that limit was crossed; earlier revisions may also contribute.")
                .Append("Crossing revision: ").AppendLine(crossingArtifact.Path)
                .AppendLine()
                .AppendLine("Cumulative evidence read so far:")
                .Append("- Physically stored cells: ").AppendLine(
                    aggregateStoredCells.ToString("N0", CultureInfo.InvariantCulture))
                .Append("- Cells with stored values, saved calculated values, or formulas: ").AppendLine(
                    contentCells.ToString("N0", CultureInfo.InvariantCulture))
                .Append("- Stored cells retained inside useful ranges: ").AppendLine(
                    retainedStoredCells.ToString("N0", CultureInfo.InvariantCulture))
                .Append("- Stored blank cells ignored outside useful ranges: ").AppendLine(
                    ignoredOutsideRange.ToString("N0", CultureInfo.InvariantCulture))
                .Append("- Effective content-bounded grid cells: ").AppendLine(
                    aggregateGridCells.ToString("N0", CultureInfo.InvariantCulture));
            if (invalidCoordinates > 0)
                message.Append("- Invalid canonical coordinates: ").AppendLine(
                    invalidCoordinates.ToString("N0", CultureInfo.InvariantCulture));

            message.AppendLine()
                .AppendLine("Current audit and alignment limits:")
                .Append("- ").Append(ComparisonAlignmentEnvelopePolicy.MaximumWorkbookStoredCells.ToString(
                    "N0", CultureInfo.InvariantCulture)).AppendLine(" physically stored cells in any one revision")
                .Append("- ").Append(ComparisonAlignmentEnvelopePolicy.MaximumAggregateStoredCells.ToString(
                    "N0", CultureInfo.InvariantCulture)).AppendLine(" physically stored cells audited across all revisions")
                .Append("- ").Append(ComparisonAlignmentEnvelopePolicy.MaximumAggregateRetainedStoredCells.ToString(
                    "N0", CultureInfo.InvariantCulture)).AppendLine(" stored cells retained inside useful ranges")
                .Append("- ").Append(ComparisonAlignmentEnvelopePolicy.MaximumAggregateGridCells.ToString(
                    "N0", CultureInfo.InvariantCulture)).AppendLine(" effective grid cells")
                .AppendLine()
                .AppendLine("Useful-range policy:")
                .AppendLine("Each sheet is bounded by the first and last rows and columns containing a stored value, saved calculated value, or formula. Physically stored blank or formatting-only cells outside that rectangle are parsed under the per-input OOXML package safety limits, counted only against the larger audit ceilings, and then released before alignment. They do not consume the retained-evidence or effective-grid ceilings. Blank-versus-missing storage inside the rectangle is retained for alignment safety but is not reported as a content difference.")
                .AppendLine()
                .AppendLine("Largest effective sheets read so far:");

            List<ComparisonInputSheetBounds> ranked = sheetBounds
                .Where(item => item.Range.TotalStoredCellCount > 0 || !item.Range.IsEmpty)
                .OrderByDescending(item => item.Range.GridCellCount)
                .ThenByDescending(item => item.Range.IgnoredOutsideRangeCellCount)
                .ThenBy(item => item.Artifact.BaseName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.SheetName, StringComparer.OrdinalIgnoreCase)
                .ToList();
            const int detailLimit = 24;
            foreach (ComparisonInputSheetBounds item in ranked.Take(detailLimit))
            {
                CanonicalUsefulRange range = item.Range;
                message.Append("- ").Append(OneLine(item.Artifact.BaseName))
                    .Append(" | ").Append(OneLine(item.SheetName))
                    .Append(" | useful ").Append(range.DisplayAddress)
                    .Append(" | ").Append(range.RowCount.ToString("N0", CultureInfo.InvariantCulture))
                    .Append(" x ").Append(range.ColumnCount.ToString("N0", CultureInfo.InvariantCulture))
                    .Append(" = ").Append(range.GridCellCount.ToString("N0", CultureInfo.InvariantCulture))
                    .Append(" grid | stored ").Append(range.TotalStoredCellCount.ToString(
                        "N0", CultureInfo.InvariantCulture))
                    .Append(" | comparison-content cells ").Append(range.ContentCellCount.ToString(
                        "N0", CultureInfo.InvariantCulture))
                    .Append(" | retained ").Append(range.RetainedStoredCellCount.ToString(
                        "N0", CultureInfo.InvariantCulture))
                    .Append(" | ignored blank tail ").Append(range.IgnoredOutsideRangeCellCount.ToString(
                        "N0", CultureInfo.InvariantCulture));
                if (range.InvalidCoordinateCellCount > 0)
                    message.Append(" | invalid coordinates ").Append(range.InvalidCoordinateCellCount.ToString(
                        "N0", CultureInfo.InvariantCulture));
                message.AppendLine();
            }
            if (ranked.Count > detailLimit)
                message.Append("- ").Append(ranked.Count - detailLimit).AppendLine(
                    " additional non-empty/physically-stored sheet(s) omitted from this dialog.");

            message.AppendLine()
                .AppendLine("Corrective actions:")
                .AppendLine("1. Inspect the reported useful ranges for a stray nonblank cell far below or to the right of the real table.")
                .AppendLine("2. Remove unintended stored values or formulas from a copy of the affected revision, save it, and retry.")
                .AppendLine("3. If the reported content bounds are genuine, split the revisions into smaller comparison projects for this alpha build.");
            return message.ToString().TrimEnd();
        }

        private static string DescribeAlignmentEnvelopeLimit(
            ComparisonAlignmentEnvelopeLimit exceededLimit)
        {
            switch (exceededLimit)
            {
                case ComparisonAlignmentEnvelopeLimit.WorkbookStoredCells:
                    return "physically stored cells in one revision";
                case ComparisonAlignmentEnvelopeLimit.AggregateStoredCells:
                    return "physically stored cells audited across all revisions";
                case ComparisonAlignmentEnvelopeLimit.AggregateRetainedStoredCells:
                    return "stored cells retained inside useful ranges";
                case ComparisonAlignmentEnvelopeLimit.AggregateGridCells:
                    return "effective content-bounded grid cells";
                default:
                    return "unknown alignment limit";
            }
        }

        private static string BuildUsefulRangeNotice(ComparisonInputBoundsSummary bounds)
        {
            if (bounds == null || bounds.IgnoredOutsideRangeCellCount == 0) return "";

            var message = new StringBuilder(768);
            message.AppendLine().AppendLine()
                .AppendLine("Useful-range notice:")
                .Append(bounds.IgnoredOutsideRangeCellCount.ToString(
                    "N0", CultureInfo.CurrentCulture))
                .AppendLine(" physically stored blank/formatting-only cell(s) outside the comparison-content boundary were ignored for alignment. No stored value, saved calculated value, or formula was excluded. Blank-versus-missing storage inside each useful range was retained for alignment safety, not reported as content.")
                .Append("Effective comparison grid: ")
                .Append(bounds.GridCellCount.ToString("N0", CultureInfo.CurrentCulture))
                .Append(" / ").Append(ComparisonAlignmentEnvelopePolicy.MaximumAggregateGridCells.ToString(
                    "N0", CultureInfo.CurrentCulture)).AppendLine(" cells.");

            foreach (ComparisonInputSheetBounds item in bounds.Sheets
                         .Where(candidate => candidate.Range.IgnoredOutsideRangeCellCount > 0)
                         .OrderByDescending(candidate => candidate.Range.IgnoredOutsideRangeCellCount)
                         .Take(5))
            {
                message.Append("- ").Append(OneLine(item.Artifact.BaseName))
                    .Append(" | ").Append(OneLine(item.SheetName))
                    .Append(": useful ").Append(item.Range.DisplayAddress)
                    .Append("; ignored ").Append(item.Range.IgnoredOutsideRangeCellCount.ToString(
                        "N0", CultureInfo.CurrentCulture)).AppendLine(" stored blank cell(s)");
            }
            return message.ToString().TrimEnd();
        }

        private static string BuildProgressUsefulRangeSuffix(ComparisonInputBoundsSummary bounds)
        {
            if (bounds == null || bounds.IgnoredOutsideRangeCellCount == 0) return "";
            return " " + bounds.IgnoredOutsideRangeCellCount.ToString(
                       "N0", CultureInfo.InvariantCulture) +
                   " stored blank/formatting-only cell(s) outside useful ranges were ignored for alignment.";
        }

        private static string OneLine(string value) =>
            string.IsNullOrEmpty(value)
                ? "(unnamed)"
                : value.Replace('\r', ' ').Replace('\n', ' ');

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

        private static void ShowFailure(string heading, Exception exception)
        {
            Dialogs.Info(heading + "\n\n" + ComparisonCommands.DescribeException(exception));
        }

        private sealed class PreviewProgressAdapter : IProgress<ComparisonPreviewProgress>
        {
            private readonly IProgress<ComparisonBuildProgressUpdate> _inner;

            internal PreviewProgressAdapter(IProgress<ComparisonBuildProgressUpdate> inner)
            {
                _inner = inner;
            }

            public void Report(ComparisonPreviewProgress value)
            {
                if (value == null) return;
                _inner?.Report(new ComparisonBuildProgressUpdate(
                    value.Phase,
                    value.Detail,
                    value.Percent,
                    value.CanCancel));
            }
        }

        private sealed class ProjectEngineProgressAdapter : IProgress<OpenXmlComparisonProjectProgress>
        {
            private readonly IProgress<ComparisonBuildProgressUpdate> _inner;

            internal ProjectEngineProgressAdapter(IProgress<ComparisonBuildProgressUpdate> inner)
            {
                _inner = inner;
            }

            public void Report(OpenXmlComparisonProjectProgress value)
            {
                if (value == null) return;
                int percent = 60 + (int)Math.Floor(
                    24D * value.CompletedWorkItems / Math.Max(1, value.TotalWorkItems));
                string detail = value.Detail;
                if (!string.IsNullOrWhiteSpace(value.SheetName))
                    detail += " (" + value.CompletedWorkItems.ToString(CultureInfo.InvariantCulture) + "/" +
                              value.TotalWorkItems.ToString(CultureInfo.InvariantCulture) + " sheets completed)";
                _inner?.Report(new ComparisonBuildProgressUpdate(
                    "Aligning revision history",
                    detail,
                    Math.Max(60, Math.Min(84, percent)),
                    true));
            }
        }
    }

    internal sealed class ComparisonInputSheetBounds
    {
        internal ComparisonInputSheetBounds(
            ComparisonProjectArtifactInfo artifact,
            string sheetName,
            CanonicalUsefulRange range)
        {
            Artifact = artifact ?? throw new ArgumentNullException(nameof(artifact));
            SheetName = sheetName ?? "";
            Range = range ?? throw new ArgumentNullException(nameof(range));
        }

        internal ComparisonProjectArtifactInfo Artifact { get; }
        internal string SheetName { get; }
        internal CanonicalUsefulRange Range { get; }
    }

    internal sealed class ComparisonInputBoundsSummary
    {
        private readonly ReadOnlyCollection<ComparisonInputSheetBounds> _sheets;

        internal ComparisonInputBoundsSummary(IEnumerable<ComparisonInputSheetBounds> sheets)
        {
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            _sheets = new ReadOnlyCollection<ComparisonInputSheetBounds>(sheets.ToList());
            StoredCellCount = _sheets.Sum(item => item.Range.TotalStoredCellCount);
            ContentCellCount = _sheets.Sum(item => item.Range.ContentCellCount);
            RetainedStoredCellCount = _sheets.Sum(item => item.Range.RetainedStoredCellCount);
            IgnoredOutsideRangeCellCount = _sheets.Sum(item => item.Range.IgnoredOutsideRangeCellCount);
            GridCellCount = _sheets.Sum(item => item.Range.GridCellCount);
        }

        internal IReadOnlyList<ComparisonInputSheetBounds> Sheets => _sheets;
        internal long StoredCellCount { get; }
        internal long ContentCellCount { get; }
        internal long RetainedStoredCellCount { get; }
        internal long IgnoredOutsideRangeCellCount { get; }
        internal long GridCellCount { get; }
    }

    /// <summary>
    /// Immutable, command-lifetime evidence retained between the load and analysis phases. Absolute source paths
    /// remain transient here and are never added to persisted review state.
    /// </summary>
    internal sealed class ComparisonProjectBuildSession
    {
        private readonly ReadOnlyDictionary<string, OpenXmlWorkbookSnapshot> _snapshots;
        private readonly ReadOnlyCollection<OpenXmlProjectRevision> _projectRevisions;
        private readonly ReadOnlyCollection<ReviewRevisionDisplayMetadata> _revisionMetadata;

        internal ComparisonProjectBuildSession(
            ComparisonProjectPreparationResult preparation,
            ComparisonProjectSetupResult setup,
            IDictionary<string, OpenXmlWorkbookSnapshot> snapshots,
            IEnumerable<OpenXmlProjectRevision> projectRevisions,
            IEnumerable<ReviewRevisionDisplayMetadata> revisionMetadata,
            ComparisonInputBoundsSummary inputBounds)
        {
            Preparation = preparation ?? throw new ArgumentNullException(nameof(preparation));
            Setup = setup ?? throw new ArgumentNullException(nameof(setup));
            InputBounds = inputBounds ?? throw new ArgumentNullException(nameof(inputBounds));
            if (snapshots == null) throw new ArgumentNullException(nameof(snapshots));

            _snapshots = new ReadOnlyDictionary<string, OpenXmlWorkbookSnapshot>(
                new Dictionary<string, OpenXmlWorkbookSnapshot>(snapshots, StringComparer.Ordinal));
            _projectRevisions = new ReadOnlyCollection<OpenXmlProjectRevision>(
                new List<OpenXmlProjectRevision>(
                    projectRevisions ?? throw new ArgumentNullException(nameof(projectRevisions))));
            _revisionMetadata = new ReadOnlyCollection<ReviewRevisionDisplayMetadata>(
                new List<ReviewRevisionDisplayMetadata>(
                    revisionMetadata ?? throw new ArgumentNullException(nameof(revisionMetadata))));

            int expectedCount = setup.Mapping.OrderedRevisionIds.Count;
            if (_snapshots.Count != expectedCount ||
                _projectRevisions.Count != expectedCount ||
                _revisionMetadata.Count != expectedCount)
            {
                throw new ArgumentException(
                    "Loaded snapshots, project revisions, and revision metadata must cover the confirmed lineage exactly.");
            }

            for (int index = 0; index < expectedCount; index++)
            {
                string revisionId = setup.Mapping.OrderedRevisionIds[index];
                OpenXmlProjectRevision revision = _projectRevisions[index];
                ReviewRevisionDisplayMetadata metadata = _revisionMetadata[index];
                if (!_snapshots.TryGetValue(revisionId, out OpenXmlWorkbookSnapshot snapshot) ||
                    !string.Equals(revision.RevisionId, revisionId, StringComparison.Ordinal) ||
                    !ReferenceEquals(revision.Snapshot, snapshot) ||
                    !string.Equals(metadata.RevisionId, revisionId, StringComparison.Ordinal) ||
                    metadata.Order != index)
                {
                    throw new ArgumentException(
                        "The loaded session does not preserve the confirmed revision order and snapshot identity.");
                }
            }
        }

        internal ComparisonProjectPreparationResult Preparation { get; }
        internal ComparisonProjectSetupResult Setup { get; }
        internal IReadOnlyDictionary<string, OpenXmlWorkbookSnapshot> Snapshots => _snapshots;
        internal IReadOnlyList<OpenXmlProjectRevision> ProjectRevisions => _projectRevisions;
        internal IReadOnlyList<ReviewRevisionDisplayMetadata> RevisionMetadata => _revisionMetadata;
        internal ComparisonInputBoundsSummary InputBounds { get; }
    }

    /// <summary>One engine pass over a retained build session; publication consumes this accepted result once.</summary>
    internal sealed class ComparisonProjectAnalysisResult
    {
        internal ComparisonProjectAnalysisResult(
            ComparisonProjectBuildSession session,
            AlignmentCorrectionGeneration corrections,
            OpenXmlComparisonProjectResult projectResult)
        {
            Session = session ?? throw new ArgumentNullException(nameof(session));
            AlignmentCorrections = corrections ?? throw new ArgumentNullException(nameof(corrections));
            ProjectResult = projectResult ?? throw new ArgumentNullException(nameof(projectResult));
            if (!string.Equals(
                    ProjectResult.ProjectId,
                    Session.Preparation.ProjectId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    ProjectResult.MappingGenerationId,
                    Session.Setup.Mapping.GenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    AlignmentCorrections.BaseMappingGenerationId,
                    Session.Setup.Mapping.GenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    ProjectResult.AlignmentCorrectionGenerationId,
                    AlignmentCorrections.GenerationId,
                    StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "The analysis result does not belong to the retained project and mapping generation.",
                    nameof(projectResult));
            }
        }

        internal ComparisonProjectBuildSession Session { get; }
        internal AlignmentCorrectionGeneration AlignmentCorrections { get; }
        internal OpenXmlComparisonProjectResult ProjectResult { get; }
    }

    internal sealed class ComparisonProjectBuildResult
    {
        internal ComparisonProjectBuildResult(
            string outputPath,
            string projectId,
            bool isComplete,
            int subjectCount,
            int transitionCount,
            int findingCount,
            ComparisonInputBoundsSummary inputBounds)
        {
            OutputPath = outputPath;
            ProjectId = projectId;
            IsComplete = isComplete;
            SubjectCount = subjectCount;
            TransitionCount = transitionCount;
            FindingCount = findingCount;
            InputBounds = inputBounds ?? throw new ArgumentNullException(nameof(inputBounds));
        }

        internal string OutputPath { get; }
        internal string ProjectId { get; }
        internal bool IsComplete { get; }
        internal int SubjectCount { get; }
        internal int TransitionCount { get; }
        internal int FindingCount { get; }
        internal ComparisonInputBoundsSummary InputBounds { get; }
    }
}
