using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.Project
{
    /// <summary>
    /// Builds logical sheet, row, and column identity over a rooted OOXML revision tree. Exact history can be
    /// projected for a safely aligned diagnostic subset, but <see cref="OpenXmlComparisonProjectResult.IsComplete"/>
    /// remains false until the full tree is resolved. Ambiguities are retained as findings and never converted into
    /// guessed physical matches.
    /// </summary>
    public static class OpenXmlComparisonProjectEngine
    {
        public static OpenXmlComparisonProjectResult Build(
            string projectId,
            string mappingGenerationId,
            IEnumerable<OpenXmlProjectRevision> revisions,
            OpenXmlComparisonProjectOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken),
            IProgress<OpenXmlComparisonProjectProgress> progress = null)
        {
            Require(projectId, nameof(projectId));
            Require(mappingGenerationId, nameof(mappingGenerationId));
            List<OpenXmlProjectRevision> ordered = ValidateAndOrder(revisions);
            options = options ?? new OpenXmlComparisonProjectOptions();
            ValidateOptions(options);
            cancellationToken.ThrowIfCancellationRequested();
            AlignmentCorrectionGeneration corrections = options.AlignmentCorrections ??
                                                        AlignmentCorrectionEngine.Create(mappingGenerationId);
            if (!string.Equals(
                    corrections.BaseMappingGenerationId,
                    mappingGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The alignment correction generation belongs to a different lineage mapping generation.");
            ValidateAlignmentCorrections(ordered, corrections, options, cancellationToken);
            string effectiveMappingGenerationId = StableId.Compute(
                "ooxml-project-effective-mapping-generation-v1",
                mappingGenerationId,
                corrections.GenerationId);

            var progressReporter = new ProjectProgressReporter(ordered, progress, cancellationToken);
            progressReporter.Report(
                OpenXmlComparisonProjectProgressStage.Preparing,
                ordered[0],
                "",
                "Preparing the confirmed revision tree for logical alignment.");

            var findings = new List<OpenXmlProjectFinding>();
            var revisionResults = new List<OpenXmlAlignedProjectRevision>(ordered.Count);
            var edges = new List<OpenXmlProjectEdgeResult>(Math.Max(0, ordered.Count - 1));
            var byRevision = new Dictionary<string, OpenXmlAlignedProjectRevision>(StringComparer.Ordinal);

            OpenXmlProjectRevision root = ordered.Single(item => item.ParentRevisionId.Length == 0);
            OpenXmlAlignedProjectRevision rootResult = BuildRootRevision(
                projectId,
                root,
                options,
                findings,
                cancellationToken,
                progressReporter);
            revisionResults.Add(rootResult);
            byRevision.Add(root.RevisionId, rootResult);

            foreach (OpenXmlProjectRevision child in ordered)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (child.ParentRevisionId.Length == 0) continue;

                OpenXmlAlignedProjectRevision parent = byRevision[child.ParentRevisionId];
                OpenXmlAlignedProjectRevision childResult = BuildChildRevision(
                    projectId,
                    parent,
                    child,
                    options,
                    findings,
                    edges,
                    corrections,
                    cancellationToken,
                    progressReporter);
                revisionResults.Add(childResult);
                byRevision.Add(child.RevisionId, childResult);
            }

            revisionResults.Sort((left, right) => left.Source.Order.CompareTo(right.Source.Order));
            progressReporter.Report(
                OpenXmlComparisonProjectProgressStage.CheckingBranches,
                ordered[ordered.Count - 1],
                "",
                "Checking independent branches across the revision tree for plausible shared insertions.");
            HashSet<string> ambiguousBranchInsertions = DetectSiblingInsertionAmbiguities(
                projectId,
                revisionResults,
                edges,
                findings,
                cancellationToken);
            if (ambiguousBranchInsertions.Count > 0)
                InvalidateAmbiguousBranches(revisionResults, edges, ambiguousBranchInsertions);

            bool complete = revisionResults.All(item => item.IsComplete) &&
                            edges.All(item => item.IsComplete) &&
                            findings.All(item => item.Severity < DiagnosticSeverity.Correctable);
            LogicalHistoryResult history = null;
            IReadOnlyList<HistoryRevision> historyRevisions = BuildReviewHistoryRevisions(
                revisionResults,
                edges,
                ambiguousBranchInsertions.Count > 0,
                cancellationToken);
            if (historyRevisions != null)
            {
                try
                {
                    progressReporter.Report(
                        OpenXmlComparisonProjectProgressStage.BuildingHistory,
                        ordered[ordered.Count - 1],
                        "",
                        complete
                            ? "Building history of stored values, saved calculated values, and formulas across " +
                              "the aligned revision tree."
                            : "Building partial history of stored values, saved calculated values, and formulas " +
                              "for safely aligned sheets.");
                    history = LogicalHistoryEngine.Build(
                        projectId,
                        effectiveMappingGenerationId,
                        historyRevisions,
                        options.ExactComparison,
                        cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    findings.Add(Finding(
                        projectId,
                        OpenXmlProjectFindingCodes.HistoryBuildFailed,
                        DiagnosticSeverity.ProjectBlocking,
                        "",
                        "",
                        "",
                        "",
                        "Logical history could not be built: " + exception.Message,
                        "Review the revision tree and aligned evidence, then retry.",
                        new[] { exception.GetType().FullName },
                        Array.Empty<string>()));
                    complete = false;
                }
            }

            progressReporter.Report(
                OpenXmlComparisonProjectProgressStage.Complete,
                ordered[ordered.Count - 1],
                "",
                complete && history != null
                    ? "Logical alignment and history are complete."
                    : "Logical alignment completed with retained findings.");

            return new OpenXmlComparisonProjectResult(
                projectId,
                mappingGenerationId,
                corrections.GenerationId,
                effectiveMappingGenerationId,
                revisionResults,
                edges,
                findings,
                history,
                complete && history != null);
        }

        /// <summary>
        /// Creates history inputs for the exact subset whose physical-to-logical alignment is proven. This permits
        /// a diagnostic report to show useful stored formula/value changes while the project remains incomplete and
        /// merge-disabled. Any unresolved sheet mapping or cross-branch insertion ambiguity suppresses partial
        /// history entirely because omission could otherwise be misreported as a deletion.
        /// </summary>
        private static IReadOnlyList<HistoryRevision> BuildReviewHistoryRevisions(
            IReadOnlyList<OpenXmlAlignedProjectRevision> revisions,
            IReadOnlyList<OpenXmlProjectEdgeResult> edges,
            bool hasAmbiguousBranchInsertion,
            CancellationToken cancellationToken)
        {
            if (hasAmbiguousBranchInsertion ||
                edges.Any(edge => edge.SheetMapping.Mapping.UnresolvedBeforeSheets.Count != 0))
                return null;

            var revisionById = revisions.ToDictionary(item => item.Source.RevisionId, StringComparer.Ordinal);
            var excludedLogicalSheets = new HashSet<string>(StringComparer.Ordinal);
            foreach (OpenXmlProjectEdgeResult edge in edges)
            {
                cancellationToken.ThrowIfCancellationRequested();
                OpenXmlAlignedProjectRevision parent = revisionById[edge.ParentRevisionId];
                OpenXmlAlignedProjectRevision child = revisionById[edge.ChildRevisionId];
                foreach (ResolvedSheetMapping mapping in edge.SheetMapping.Mapping.Mappings)
                {
                    OpenXmlAlignedSheetRevision parentSheet = parent.Sheets.FirstOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, mapping.Before.SheetKey, StringComparison.Ordinal));
                    OpenXmlAlignedSheetRevision childSheet = child.Sheets.FirstOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, mapping.After.SheetKey, StringComparison.Ordinal));
                    OpenXmlProjectSheetAlignment alignment = edge.SheetAlignments.FirstOrDefault(item =>
                        string.Equals(item.BeforeSheetKey, mapping.Before.SheetKey, StringComparison.Ordinal) &&
                        string.Equals(item.AfterSheetKey, mapping.After.SheetKey, StringComparison.Ordinal));
                    if (parentSheet != null && childSheet != null && alignment?.IsComplete == true) continue;
                    if (parentSheet != null) excludedLogicalSheets.Add(parentSheet.LogicalSheetId);
                    if (childSheet != null) excludedLogicalSheets.Add(childSheet.LogicalSheetId);
                }
            }

            bool hasReviewableSheet = revisions.Any(revision => revision.Sheets.Any(sheet =>
                !excludedLogicalSheets.Contains(sheet.LogicalSheetId)));
            if (!hasReviewableSheet) return null;

            var result = new List<HistoryRevision>(revisions.Count);
            foreach (OpenXmlAlignedProjectRevision revision in revisions.OrderBy(item => item.Source.Order))
            {
                cancellationToken.ThrowIfCancellationRequested();
                RevisionSnapshot snapshot = CreateAlignedWorkbookSnapshot(
                    revision.Source.RevisionId,
                    revision.Sheets.Where(sheet => !excludedLogicalSheets.Contains(sheet.LogicalSheetId)),
                    cancellationToken);
                result.Add(new HistoryRevision(
                    revision.Source.RevisionId,
                    revision.Source.ParentRevisionId,
                    revision.Source.Order,
                    snapshot,
                    revision.Source.IsMaster,
                    revision.Source.IsTrustedBaseline));
            }
            return new ReadOnlyCollection<HistoryRevision>(result);
        }

        private static OpenXmlAlignedProjectRevision BuildRootRevision(
            string projectId,
            OpenXmlProjectRevision revision,
            OpenXmlComparisonProjectOptions options,
            ICollection<OpenXmlProjectFinding> findings,
            CancellationToken cancellationToken,
            ProjectProgressReporter progress)
        {
            int findingStart = findings.Count;
            progress.Report(
                OpenXmlComparisonProjectProgressStage.FingerprintingSheets,
                revision,
                "",
                "Fingerprinting sheets in revision '" + revision.RevisionId + "'.");
            OpenXmlWorkbookSheetFingerprintSet evidenceSet = OpenXmlSheetFingerprintBuilder.Build(
                revision.Snapshot,
                cancellationToken);
            var sheets = new List<OpenXmlAlignedSheetRevision>();
            foreach (OpenXmlSheetFingerprintEvidence evidence in evidenceSet.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress.Report(
                    OpenXmlComparisonProjectProgressStage.ProjectingGrid,
                    revision,
                    evidence.Fingerprint.Name,
                    "Projecting the bounded grid for root sheet '" + evidence.Fingerprint.Name + "'.");
                if (!ValidateEvidence(projectId, revision.RevisionId, "", "", evidence, findings))
                {
                    progress.CompleteSheet(revision, evidence.Fingerprint.SheetKey, evidence.Fingerprint.Name);
                    continue;
                }

                CanonicalGridProjection grid;
                try
                {
                    grid = CanonicalGridProjectionBuilder.Build(
                        evidence.Source.Snapshot,
                        options.GridProjection,
                        cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    findings.Add(GridFailure(projectId, "", revision.RevisionId, "", evidence, exception));
                    progress.CompleteSheet(revision, evidence.Fingerprint.SheetKey, evidence.Fingerprint.Name);
                    continue;
                }

                string logicalSheetId = StableId.Compute(
                    "logical-sheet-root-v1",
                    projectId,
                    revision.RevisionId,
                    evidence.Fingerprint.SheetKey);
                LogicalAxisIdentityProjection rows = LogicalAxisIdentityProjector.CreateRoot(
                    projectId,
                    logicalSheetId,
                    revision.RevisionId,
                    LogicalAxisKind.Row,
                    grid.Table.RowCount);
                LogicalAxisIdentityProjection columns = LogicalAxisIdentityProjector.CreateRoot(
                    projectId,
                    logicalSheetId,
                    revision.RevisionId,
                    LogicalAxisKind.Column,
                    grid.Table.ColumnCount);
                sheets.Add(new OpenXmlAlignedSheetRevision(
                    evidence.Fingerprint.SheetKey,
                    evidence.Fingerprint.Name,
                    logicalSheetId,
                    grid,
                    rows,
                    columns));
                progress.CompleteSheet(revision, evidence.Fingerprint.SheetKey, evidence.Fingerprint.Name);
            }
            progress.CompleteRevision(revision);

            bool complete = sheets.Count == evidenceSet.Sheets.Count &&
                            !findings.Skip(findingStart)
                                .Any(item => item.Severity >= DiagnosticSeverity.Correctable);
            RevisionSnapshot aligned = complete
                ? CreateAlignedWorkbookSnapshot(revision.RevisionId, sheets, cancellationToken)
                : null;
            return new OpenXmlAlignedProjectRevision(revision, sheets, aligned, complete);
        }

        private static OpenXmlAlignedProjectRevision BuildChildRevision(
            string projectId,
            OpenXmlAlignedProjectRevision parent,
            OpenXmlProjectRevision child,
            OpenXmlComparisonProjectOptions options,
            ICollection<OpenXmlProjectFinding> findings,
            ICollection<OpenXmlProjectEdgeResult> edges,
            AlignmentCorrectionGeneration corrections,
            CancellationToken cancellationToken,
            ProjectProgressReporter progress)
        {
            int findingStart = findings.Count;
            progress.Report(
                OpenXmlComparisonProjectProgressStage.MappingSheets,
                child,
                "",
                "Mapping sheets from revision '" + parent.Source.RevisionId + "' to '" + child.RevisionId + "'.");
            OpenXmlWorkbookSheetMappingResult mapping = OpenXmlWorkbookSheetMappingInference.Infer(
                parent.Source.Snapshot,
                child.Snapshot,
                options.SheetMapping,
                cancellationToken);
            mapping = ApplyConfirmedSheetMappings(
                mapping,
                corrections.GetSheetPairs(parent.Source.RevisionId, child.RevisionId),
                corrections.GetUnmatchedSheets(parent.Source.RevisionId, child.RevisionId));
            foreach (AlignmentKeyColumnsConfirmation keyConfirmation in corrections.KeyColumns.Where(item =>
                         string.Equals(item.ParentRevisionId, parent.Source.RevisionId, StringComparison.Ordinal) &&
                         string.Equals(item.ChildRevisionId, child.RevisionId, StringComparison.Ordinal)))
            {
                if (!mapping.Mapping.Mappings.Any(item =>
                        string.Equals(item.Before.SheetKey, keyConfirmation.BeforeSheetKey, StringComparison.Ordinal) &&
                        string.Equals(item.After.SheetKey, keyConfirmation.AfterSheetKey, StringComparison.Ordinal)))
                    throw new InvalidOperationException(
                        "A key-column confirmation does not match an accepted exact sheet pair on edge '" +
                        parent.Source.RevisionId + "' -> '" + child.RevisionId +
                        "'. Confirm the sheet pair first; the engine will not redirect the key by name or position.");
            }
            var alignments = new List<OpenXmlProjectSheetAlignment>();
            var childSheets = new List<OpenXmlAlignedSheetRevision>();
            var parentByPhysicalKey = parent.Sheets.ToDictionary(item => item.PhysicalSheetKey, StringComparer.Ordinal);
            var beforeEvidence = mapping.Before.Sheets.ToDictionary(
                item => item.Fingerprint.SheetKey,
                StringComparer.Ordinal);
            var afterEvidence = mapping.After.Sheets.ToDictionary(
                item => item.Fingerprint.SheetKey,
                StringComparer.Ordinal);

            foreach (UnresolvedSheetMapping unresolved in mapping.Mapping.UnresolvedBeforeSheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.SheetMappingAmbiguous,
                    DiagnosticSeverity.Correctable,
                    parent.Source.RevisionId,
                    child.RevisionId,
                    ParentLogicalSheetId(parentByPhysicalKey, unresolved.Before.SheetKey),
                    unresolved.Before.Name,
                    "Sheet '" + unresolved.Before.Name + "' has no safe one-to-one mapping in revision '" +
                    child.RevisionId + "'.",
                    "Confirm the exact source/target sheet pair. If the source was deleted and the target is a new " +
                    "sheet, choose the explicit deleted + new option instead.",
                    unresolved.ReasonCodes,
                    unresolved.Candidates.Select(candidate => candidate.After.Name)));
            }

            foreach (ResolvedSheetMapping resolved in mapping.Mapping.Mappings)
            {
                cancellationToken.ThrowIfCancellationRequested();
                OpenXmlSheetFingerprintEvidence before = beforeEvidence[resolved.Before.SheetKey];
                OpenXmlSheetFingerprintEvidence after = afterEvidence[resolved.After.SheetKey];
                if (!parentByPhysicalKey.TryGetValue(resolved.Before.SheetKey, out OpenXmlAlignedSheetRevision parentSheet))
                {
                    findings.Add(Finding(
                        projectId,
                        OpenXmlProjectFindingCodes.ParentSheetStateUnavailable,
                        DiagnosticSeverity.ProjectBlocking,
                        parent.Source.RevisionId,
                        child.RevisionId,
                        "",
                        resolved.Before.Name,
                        "The resolved sheet mapping points to a parent sheet whose logical state is unavailable.",
                        "Resolve the parent revision's blocking findings before extending this branch.",
                        resolved.ReasonCodes,
                        new[] { resolved.After.Name }));
                    continue;
                }
                if (!ValidateEvidence(
                        projectId,
                        child.RevisionId,
                        parent.Source.RevisionId,
                        parentSheet.LogicalSheetId,
                        after,
                        findings))
                    continue;

                OpenXmlProjectSheetAlignment alignment = AlignMappedSheet(
                    projectId,
                    parent,
                    child,
                    parentSheet,
                    before,
                    after,
                    options,
                    corrections,
                    findings,
                    cancellationToken,
                    progress);
                alignments.Add(alignment);
                if (alignment.IsComplete)
                {
                    CanonicalGridProjection childGrid = CanonicalGridProjectionBuilder.Build(
                        after.Source.Snapshot,
                        options.GridProjection,
                        cancellationToken);
                    childSheets.Add(new OpenXmlAlignedSheetRevision(
                        after.Fingerprint.SheetKey,
                        after.Fingerprint.Name,
                        parentSheet.LogicalSheetId,
                        childGrid,
                        alignment.Rows,
                        alignment.Columns));
                }
                progress.CompleteSheet(child, after.Fingerprint.SheetKey, after.Fingerprint.Name);
            }

            // A target unclaimed by any source is safe to call new only when no unresolved source could still be
            // its renamed predecessor.
            if (mapping.Mapping.UnresolvedBeforeSheets.Count == 0)
            {
                foreach (WorkbookSheetFingerprint unmatched in mapping.Mapping.UnmatchedAfterSheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    OpenXmlSheetFingerprintEvidence evidence = afterEvidence[unmatched.SheetKey];
                    if (!ValidateEvidence(projectId, child.RevisionId, parent.Source.RevisionId, "", evidence, findings))
                        continue;
                    CanonicalGridProjection grid;
                    try
                    {
                        grid = CanonicalGridProjectionBuilder.Build(
                            evidence.Source.Snapshot,
                            options.GridProjection,
                            cancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        findings.Add(GridFailure(
                            projectId,
                            parent.Source.RevisionId,
                            child.RevisionId,
                            "",
                            evidence,
                            exception));
                        continue;
                    }

                    string logicalSheetId = StableId.Compute(
                        "logical-sheet-inserted-v1",
                        projectId,
                        child.RevisionId,
                        evidence.Fingerprint.SheetKey);
                    childSheets.Add(new OpenXmlAlignedSheetRevision(
                        evidence.Fingerprint.SheetKey,
                        evidence.Fingerprint.Name,
                        logicalSheetId,
                        grid,
                        LogicalAxisIdentityProjector.CreateRoot(
                            projectId,
                            logicalSheetId,
                            child.RevisionId,
                            LogicalAxisKind.Row,
                            grid.Table.RowCount),
                        LogicalAxisIdentityProjector.CreateRoot(
                            projectId,
                            logicalSheetId,
                            child.RevisionId,
                            LogicalAxisKind.Column,
                            grid.Table.ColumnCount)));
                    progress.CompleteSheet(child, evidence.Fingerprint.SheetKey, evidence.Fingerprint.Name);
                }
            }

            progress.CompleteRevision(child);

            bool edgeComplete = parent.IsComplete &&
                                !findings.Skip(findingStart)
                                    .Any(item => item.Severity >= DiagnosticSeverity.Correctable) &&
                                childSheets.Count == mapping.After.Sheets.Count &&
                                alignments.All(item => item.IsComplete);
            RevisionSnapshot aligned = edgeComplete
                ? CreateAlignedWorkbookSnapshot(child.RevisionId, childSheets, cancellationToken)
                : null;
            var result = new OpenXmlAlignedProjectRevision(child, childSheets, aligned, edgeComplete);
            edges.Add(new OpenXmlProjectEdgeResult(
                parent.Source.RevisionId,
                child.RevisionId,
                mapping,
                alignments,
                edgeComplete));
            return result;
        }

        private static HashSet<string> DetectSiblingInsertionAmbiguities(
            string projectId,
            IReadOnlyList<OpenXmlAlignedProjectRevision> revisions,
            IReadOnlyList<OpenXmlProjectEdgeResult> edges,
            ICollection<OpenXmlProjectFinding> findings,
            CancellationToken cancellationToken)
        {
            var byRevision = revisions.ToDictionary(item => item.Source.RevisionId, StringComparer.Ordinal);
            var candidates = new List<SiblingInsertionCandidate>();
            var sheetCandidates = new List<IndependentSheetInsertionCandidate>();
            foreach (OpenXmlProjectEdgeResult edge in edges)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!byRevision.TryGetValue(edge.ParentRevisionId, out OpenXmlAlignedProjectRevision parent) ||
                    !byRevision.TryGetValue(edge.ChildRevisionId, out OpenXmlAlignedProjectRevision child))
                    continue;

                foreach (OpenXmlProjectSheetAlignment alignment in edge.SheetAlignments)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!alignment.IsComplete || alignment.Rows == null || alignment.Columns == null) continue;
                    OpenXmlAlignedSheetRevision parentSheet = parent.Sheets.SingleOrDefault(item =>
                        string.Equals(item.LogicalSheetId, alignment.LogicalSheetId, StringComparison.Ordinal));
                    OpenXmlAlignedSheetRevision childSheet = child.Sheets.SingleOrDefault(item =>
                        string.Equals(item.PhysicalSheetKey, alignment.AfterSheetKey, StringComparison.Ordinal));
                    if (parentSheet == null || childSheet == null) continue;

                    AddSiblingRowInsertionCandidates(edge, alignment, parentSheet, childSheet, candidates);
                    AddSiblingColumnInsertionCandidates(edge, alignment, childSheet, candidates);
                }
                AddIndependentSheetInsertionCandidates(edge, child, sheetCandidates);
            }

            // A token is project-ambiguous only when it was independently introduced on incomparable lineages.
            // Ordinary propagation (or even a later reinsertion) along one ancestor/descendant chain is not a
            // cross-branch identity claim and therefore remains outside this detector.
            var affected = new HashSet<string>(StringComparer.Ordinal);
            ReportAxisBranchAmbiguities(projectId, candidates, byRevision, findings, affected);
            ReportSheetBranchAmbiguities(projectId, sheetCandidates, byRevision, findings, affected);
            return affected;
        }

        private static void ReportAxisBranchAmbiguities(
            string projectId,
            IEnumerable<SiblingInsertionCandidate> candidates,
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisions,
            ICollection<OpenXmlProjectFinding> findings,
            ISet<string> affected)
        {
            foreach (IGrouping<string, SiblingInsertionCandidate> group in candidates
                         .GroupBy(item => item.GroupKey, StringComparer.Ordinal)
                         .OrderBy(item => item.Key, StringComparer.Ordinal))
            {
                List<SiblingInsertionCandidate> distinctChildren = group
                    .GroupBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                    .Select(item => item.OrderBy(candidate => candidate.AfterIndex).First())
                    .OrderBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                    .ToList();
                if (distinctChildren.Count < 2) continue;

                foreach (SiblingInsertionCandidate candidate in distinctChildren)
                {
                    List<SiblingInsertionCandidate> independentPeers = distinctChildren
                        .Where(item => !string.Equals(
                                           item.ChildRevisionId,
                                           candidate.ChildRevisionId,
                                           StringComparison.Ordinal) &&
                                       AreOnIndependentLineages(
                                           candidate.ChildRevisionId,
                                           item.ChildRevisionId,
                                           revisions))
                        .OrderBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                        .ToList();
                    if (independentPeers.Count == 0) continue;

                    affected.Add(candidate.ChildRevisionId);
                    bool row = candidate.Axis == LogicalAxisKind.Row;
                    string axisName = row ? "row" : "column";
                    string evidence = row
                        ? "the same strong/confirmed exact key tuple"
                        : "the same unique accepted header label";
                    string peerIds = string.Join(
                        ", ",
                        independentPeers.Select(item => "'" + item.ChildRevisionId + "'"));
                    string commonAncestor = NearestCommonAncestor(
                        new[] { candidate.ChildRevisionId }
                            .Concat(independentPeers.Select(item => item.ChildRevisionId)),
                        revisions);
                    findings.Add(Finding(
                        projectId,
                        row
                            ? OpenXmlProjectFindingCodes.SiblingRowInsertionAmbiguous
                            : OpenXmlProjectFindingCodes.SiblingColumnInsertionAmbiguous,
                        DiagnosticSeverity.Correctable,
                        candidate.ParentRevisionId,
                        candidate.ChildRevisionId,
                        candidate.LogicalSheetId,
                        candidate.SheetName,
                        "Revision '" + candidate.ChildRevisionId + "' independently inserts " + axisName +
                        " at zero-based aligned index " + candidate.AfterIndex.ToString(CultureInfo.InvariantCulture) +
                        "; non-ancestor branch revision(s) " + peerIds + " contain " + evidence +
                        ". Their nearest common ancestor is '" + commonAncestor +
                        "'. The engine did not guess whether they share one logical identity.",
                        "Compare the sibling branches separately or make the intended " + axisName +
                        " identity unambiguous in comparison copies before rebuilding the project.",
                        new[]
                        {
                            row
                                ? "sibling-insertion.same-strong-key-tuple"
                                : "sibling-insertion.same-unique-header-label",
                            "branch-insertion.non-ancestor-lineages",
                            "sibling-insertion.no-automatic-reconciliation",
                        },
                        independentPeers.Select(item => item.SheetName)
                            .Concat(new[] { candidate.SheetName })));
                }
            }
        }

        private static void ReportSheetBranchAmbiguities(
            string projectId,
            IEnumerable<IndependentSheetInsertionCandidate> candidates,
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisions,
            ICollection<OpenXmlProjectFinding> findings,
            ISet<string> affected)
        {
            foreach (IGrouping<string, IndependentSheetInsertionCandidate> group in candidates
                         .GroupBy(item => item.SheetName, StringComparer.OrdinalIgnoreCase)
                         .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                List<IndependentSheetInsertionCandidate> distinctChildren = group
                    .GroupBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                    .Select(item => item.OrderBy(candidate => candidate.PhysicalSheetKey, StringComparer.Ordinal).First())
                    .OrderBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                    .ToList();
                if (distinctChildren.Count < 2) continue;

                foreach (IndependentSheetInsertionCandidate candidate in distinctChildren)
                {
                    List<IndependentSheetInsertionCandidate> independentPeers = distinctChildren
                        .Where(item => !string.Equals(
                                           item.ChildRevisionId,
                                           candidate.ChildRevisionId,
                                           StringComparison.Ordinal) &&
                                       AreOnIndependentLineages(
                                           candidate.ChildRevisionId,
                                           item.ChildRevisionId,
                                           revisions))
                        .OrderBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                        .ToList();
                    if (independentPeers.Count == 0) continue;

                    affected.Add(candidate.ChildRevisionId);
                    string peerIds = string.Join(
                        ", ",
                        independentPeers.Select(item => "'" + item.ChildRevisionId + "'"));
                    string commonAncestor = NearestCommonAncestor(
                        new[] { candidate.ChildRevisionId }
                            .Concat(independentPeers.Select(item => item.ChildRevisionId)),
                        revisions);
                    findings.Add(Finding(
                        projectId,
                        OpenXmlProjectFindingCodes.IndependentSheetInsertionAmbiguous,
                        DiagnosticSeverity.Correctable,
                        candidate.ParentRevisionId,
                        candidate.ChildRevisionId,
                        candidate.LogicalSheetId,
                        candidate.SheetName,
                        "Revision '" + candidate.ChildRevisionId + "' independently inserts sheet '" +
                        candidate.SheetName + "'; non-ancestor branch revision(s) " + peerIds +
                        " independently insert the same case-insensitive sheet name. Their nearest common ancestor " +
                        "is '" + commonAncestor + "'. The engine did not guess whether they share one logical sheet.",
                        "Rename one independently created sheet in a comparison copy, compare the branches " +
                        "separately, or wait for explicit logical-sheet reconciliation before rebuilding the project.",
                        new[]
                        {
                            "branch-sheet-insertion.same-case-insensitive-name",
                            "branch-insertion.non-ancestor-lineages",
                            "branch-sheet-insertion.no-automatic-reconciliation",
                        },
                        independentPeers.Select(item => item.SheetName)
                            .Concat(new[] { candidate.SheetName })));
                }
            }
        }

        private static bool AreOnIndependentLineages(
            string leftRevisionId,
            string rightRevisionId,
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisions) =>
            !IsAncestor(leftRevisionId, rightRevisionId, revisions) &&
            !IsAncestor(rightRevisionId, leftRevisionId, revisions);

        private static bool IsAncestor(
            string possibleAncestorRevisionId,
            string revisionId,
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisions)
        {
            string cursor = revisionId;
            while (cursor.Length > 0)
            {
                if (string.Equals(cursor, possibleAncestorRevisionId, StringComparison.Ordinal)) return true;
                cursor = revisions[cursor].Source.ParentRevisionId;
            }
            return false;
        }

        private static string NearestCommonAncestor(
            IEnumerable<string> revisionIds,
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisions)
        {
            List<List<string>> lineages = revisionIds
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .Select(item => LineageFromRoot(item, revisions))
                .ToList();
            if (lineages.Count == 0) return "";

            string common = lineages[0][0];
            int limit = lineages.Min(item => item.Count);
            for (int index = 0; index < limit; index++)
            {
                string candidate = lineages[0][index];
                if (lineages.Any(item => !string.Equals(item[index], candidate, StringComparison.Ordinal))) break;
                common = candidate;
            }
            return common;
        }

        private static List<string> LineageFromRoot(
            string revisionId,
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisions)
        {
            var lineage = new List<string>();
            string cursor = revisionId;
            while (cursor.Length > 0)
            {
                lineage.Add(cursor);
                cursor = revisions[cursor].Source.ParentRevisionId;
            }
            lineage.Reverse();
            return lineage;
        }

        private static void AddSiblingRowInsertionCandidates(
            OpenXmlProjectEdgeResult edge,
            OpenXmlProjectSheetAlignment alignment,
            OpenXmlAlignedSheetRevision parentSheet,
            OpenXmlAlignedSheetRevision childSheet,
            ICollection<SiblingInsertionCandidate> candidates)
        {
            KeyInferenceResult key = alignment.KeyInference;
            if (key == null || !key.HasUsableKey || key.SelectedCandidate == null) return;

            var keySchema = new List<string>();
            foreach (int beforeColumn in key.SelectedCandidate.BeforeColumns)
            {
                if (!parentSheet.Columns.IdentityByPhysicalIndex.TryGetValue(beforeColumn, out string identity))
                    return;
                keySchema.Add(identity);
            }

            foreach (LogicalAxisStructuralEvent insertion in alignment.Rows.StructuralEvents
                         .Where(item => item.Kind == LogicalAxisStructuralEventKind.Insertion && item.AfterIndex.HasValue)
                         .OrderBy(item => item.AfterIndex.Value))
            {
                int rowIndex = insertion.AfterIndex.Value;
                if (!HasCompleteExactTuple(
                        childSheet.Grid.Table,
                        rowIndex,
                        key.SelectedCandidate.AfterColumns,
                        out string exactTuple))
                    continue;

                var tokenComponents = new List<string>(keySchema) { exactTuple };
                string evidenceToken = StableId.Compute(
                    "sibling-inserted-row-key-token-v1",
                    tokenComponents.ToArray());
                candidates.Add(new SiblingInsertionCandidate(
                    edge.ParentRevisionId,
                    edge.ChildRevisionId,
                    alignment.LogicalSheetId,
                    alignment.AfterSheetName,
                    LogicalAxisKind.Row,
                    rowIndex,
                    evidenceToken));
            }
        }

        private static void AddSiblingColumnInsertionCandidates(
            OpenXmlProjectEdgeResult edge,
            OpenXmlProjectSheetAlignment alignment,
            OpenXmlAlignedSheetRevision childSheet,
            ICollection<SiblingInsertionCandidate> candidates)
        {
            HeaderDetectionResult header = alignment.AfterHeader;
            if (header == null || !header.HasAcceptedHeader || !header.AcceptedHeaderRowIndex.HasValue) return;
            AlignmentTable table = childSheet.Grid.Table;
            int headerRow = header.AcceptedHeaderRowIndex.Value;

            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int column = 0; column < table.ColumnCount; column++)
            {
                string normalized = AlignmentText.NormalizeLabel(table.GetCellOrNull(headerRow, column));
                if (normalized.Length == 0) continue;
                counts.TryGetValue(normalized, out int count);
                counts[normalized] = count + 1;
            }

            foreach (LogicalAxisStructuralEvent insertion in alignment.Columns.StructuralEvents
                         .Where(item => item.Kind == LogicalAxisStructuralEventKind.Insertion && item.AfterIndex.HasValue)
                         .OrderBy(item => item.AfterIndex.Value))
            {
                int columnIndex = insertion.AfterIndex.Value;
                string normalized = AlignmentText.NormalizeLabel(table.GetCellOrNull(headerRow, columnIndex));
                if (normalized.Length == 0 || !counts.TryGetValue(normalized, out int count) || count != 1) continue;
                string evidenceToken = StableId.Compute("sibling-inserted-column-header-token-v1", normalized);
                candidates.Add(new SiblingInsertionCandidate(
                    edge.ParentRevisionId,
                    edge.ChildRevisionId,
                    alignment.LogicalSheetId,
                    alignment.AfterSheetName,
                    LogicalAxisKind.Column,
                    columnIndex,
                    evidenceToken));
            }
        }

        private static void AddIndependentSheetInsertionCandidates(
            OpenXmlProjectEdgeResult edge,
            OpenXmlAlignedProjectRevision child,
            ICollection<IndependentSheetInsertionCandidate> candidates)
        {
            // Mapping-unmatched sheets are the only sheets treated as independently created. A mapped sheet,
            // including one inherited from an earlier insertion on the same lineage, retains its logical identity.
            foreach (WorkbookSheetFingerprint unmatched in edge.SheetMapping.Mapping.UnmatchedAfterSheets
                         .OrderBy(item => item.SheetKey, StringComparer.Ordinal))
            {
                OpenXmlAlignedSheetRevision inserted = child.Sheets.SingleOrDefault(item =>
                    string.Equals(item.PhysicalSheetKey, unmatched.SheetKey, StringComparison.Ordinal));
                if (inserted == null || string.IsNullOrEmpty(unmatched.Name)) continue;
                candidates.Add(new IndependentSheetInsertionCandidate(
                    edge.ParentRevisionId,
                    edge.ChildRevisionId,
                    inserted.LogicalSheetId,
                    unmatched.SheetKey,
                    unmatched.Name));
            }
        }

        private static bool HasCompleteExactTuple(
            AlignmentTable table,
            int rowIndex,
            IReadOnlyList<int> columns,
            out string exactTuple)
        {
            exactTuple = "";
            if (rowIndex < 0 || rowIndex >= table.RowCount || columns == null || columns.Count == 0) return false;
            foreach (int column in columns)
            {
                if (!table.TryGetCell(rowIndex, column, out string value) || AlignmentText.IsBlank(value))
                    return false;
            }
            exactTuple = AlignmentText.ExactTuple(table, rowIndex, columns);
            return true;
        }

        private static void InvalidateAmbiguousBranches(
            IList<OpenXmlAlignedProjectRevision> revisions,
            IList<OpenXmlProjectEdgeResult> edges,
            IEnumerable<string> directlyAffectedRevisionIds)
        {
            var invalid = new HashSet<string>(directlyAffectedRevisionIds, StringComparer.Ordinal);
            bool added;
            do
            {
                added = false;
                foreach (OpenXmlAlignedProjectRevision revision in revisions)
                    if (revision.Source.ParentRevisionId.Length > 0 &&
                        invalid.Contains(revision.Source.ParentRevisionId) &&
                        invalid.Add(revision.Source.RevisionId))
                        added = true;
            } while (added);

            for (int index = 0; index < revisions.Count; index++)
            {
                OpenXmlAlignedProjectRevision revision = revisions[index];
                if (!invalid.Contains(revision.Source.RevisionId)) continue;
                revisions[index] = new OpenXmlAlignedProjectRevision(
                    revision.Source,
                    revision.Sheets,
                    null,
                    false);
            }
            for (int index = 0; index < edges.Count; index++)
            {
                OpenXmlProjectEdgeResult edge = edges[index];
                if (!invalid.Contains(edge.ChildRevisionId)) continue;
                edges[index] = new OpenXmlProjectEdgeResult(
                    edge.ParentRevisionId,
                    edge.ChildRevisionId,
                    edge.SheetMapping,
                    edge.SheetAlignments,
                    false);
            }
        }

        private sealed class SiblingInsertionCandidate
        {
            internal SiblingInsertionCandidate(
                string parentRevisionId,
                string childRevisionId,
                string logicalSheetId,
                string sheetName,
                LogicalAxisKind axis,
                int afterIndex,
                string evidenceToken)
            {
                ParentRevisionId = parentRevisionId;
                ChildRevisionId = childRevisionId;
                LogicalSheetId = logicalSheetId;
                SheetName = sheetName ?? "";
                Axis = axis;
                AfterIndex = afterIndex;
                GroupKey = StableId.Compute(
                    "independent-branch-axis-insertion-group-v2",
                    logicalSheetId,
                    axis == LogicalAxisKind.Row ? "row" : "column",
                    evidenceToken);
            }

            internal string ParentRevisionId { get; }
            internal string ChildRevisionId { get; }
            internal string LogicalSheetId { get; }
            internal string SheetName { get; }
            internal LogicalAxisKind Axis { get; }
            internal int AfterIndex { get; }
            internal string GroupKey { get; }
        }

        private sealed class IndependentSheetInsertionCandidate
        {
            internal IndependentSheetInsertionCandidate(
                string parentRevisionId,
                string childRevisionId,
                string logicalSheetId,
                string physicalSheetKey,
                string sheetName)
            {
                ParentRevisionId = parentRevisionId;
                ChildRevisionId = childRevisionId;
                LogicalSheetId = logicalSheetId;
                PhysicalSheetKey = physicalSheetKey;
                SheetName = sheetName;
            }

            internal string ParentRevisionId { get; }
            internal string ChildRevisionId { get; }
            internal string LogicalSheetId { get; }
            internal string PhysicalSheetKey { get; }
            internal string SheetName { get; }
        }

        private sealed class ProjectProgressReporter
        {
            private readonly IProgress<OpenXmlComparisonProjectProgress> _progress;
            private readonly CancellationToken _cancellationToken;
            private readonly Dictionary<string, int> _revisionTargets;
            private readonly HashSet<string> _completedSheets = new HashSet<string>(StringComparer.Ordinal);
            private int _completedWorkItems;

            internal ProjectProgressReporter(
                IReadOnlyList<OpenXmlProjectRevision> revisions,
                IProgress<OpenXmlComparisonProjectProgress> progress,
                CancellationToken cancellationToken)
            {
                _progress = progress;
                _cancellationToken = cancellationToken;
                _revisionTargets = new Dictionary<string, int>(StringComparer.Ordinal);
                int total = 0;
                foreach (OpenXmlProjectRevision revision in revisions)
                {
                    total = checked(total + Math.Max(1, revision.Snapshot.Sheets.Count));
                    _revisionTargets.Add(revision.RevisionId, total);
                }
                TotalWorkItems = total;
            }

            internal int TotalWorkItems { get; }

            internal void Report(
                OpenXmlComparisonProjectProgressStage stage,
                OpenXmlProjectRevision revision,
                string sheetName,
                string detail)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                _progress?.Report(new OpenXmlComparisonProjectProgress(
                    stage,
                    _completedWorkItems,
                    TotalWorkItems,
                    revision?.RevisionId ?? "",
                    sheetName,
                    detail));
                // Inline progress implementations may cancel synchronously. Honor that before the expensive step.
                _cancellationToken.ThrowIfCancellationRequested();
            }

            internal void CompleteSheet(
                OpenXmlProjectRevision revision,
                string physicalSheetKey,
                string sheetName)
            {
                string key = revision.RevisionId + "\u0000" + (physicalSheetKey ?? "");
                if (_completedSheets.Add(key) && _completedWorkItems < TotalWorkItems)
                    _completedWorkItems++;
                Report(
                    OpenXmlComparisonProjectProgressStage.PropagatingIdentities,
                    revision,
                    sheetName,
                    "Completed logical identity work for sheet '" + sheetName + "'.");
            }

            internal void CompleteRevision(OpenXmlProjectRevision revision)
            {
                if (_revisionTargets.TryGetValue(revision.RevisionId, out int target))
                    _completedWorkItems = Math.Max(_completedWorkItems, target);
                Report(
                    OpenXmlComparisonProjectProgressStage.PropagatingIdentities,
                    revision,
                    "",
                    "Completed logical identity work for revision '" + revision.RevisionId + "'.");
            }
        }

        private static OpenXmlProjectSheetAlignment AlignMappedSheet(
            string projectId,
            OpenXmlAlignedProjectRevision parent,
            OpenXmlProjectRevision child,
            OpenXmlAlignedSheetRevision parentSheet,
            OpenXmlSheetFingerprintEvidence before,
            OpenXmlSheetFingerprintEvidence after,
            OpenXmlComparisonProjectOptions options,
            AlignmentCorrectionGeneration corrections,
            ICollection<OpenXmlProjectFinding> findings,
            CancellationToken cancellationToken,
            ProjectProgressReporter progress)
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress.Report(
                OpenXmlComparisonProjectProgressStage.ProjectingGrid,
                child,
                after.Fingerprint.Name,
                "Projecting the bounded grid for sheet '" + after.Fingerprint.Name + "'.");
            CanonicalGridProjection afterGrid;
            try
            {
                afterGrid = CanonicalGridProjectionBuilder.Build(
                    after.Source.Snapshot,
                    options.GridProjection,
                    cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                findings.Add(GridFailure(
                    projectId,
                    parent.Source.RevisionId,
                    child.RevisionId,
                    parentSheet.LogicalSheetId,
                    after,
                    exception));
                return EmptyAlignment(parentSheet, before, after);
            }

            AlignmentTable beforeTable = parentSheet.Grid.Table;
            AlignmentTable afterTable = afterGrid.Table;
            bool hasExactPositionalEvidence = HasExactPositionalCellEvidence(
                before,
                parentSheet.Grid,
                after,
                afterGrid,
                cancellationToken);
            progress.Report(
                OpenXmlComparisonProjectProgressStage.DetectingHeaders,
                child,
                after.Fingerprint.Name,
                "Detecting comparable headers in sheet '" + after.Fingerprint.Name + "'.");
            bool hasBeforeHeaderCorrection = corrections.TryGetHeaderRow(
                parent.Source.RevisionId,
                before.Fingerprint.SheetKey,
                out AlignmentHeaderRowConfirmation beforeHeaderCorrection);
            bool hasAfterHeaderCorrection = corrections.TryGetHeaderRow(
                child.RevisionId,
                after.Fingerprint.SheetKey,
                out AlignmentHeaderRowConfirmation afterHeaderCorrection);
            HeaderDetectionResult beforeHeader = HeaderDetector.Detect(
                beforeTable,
                afterTable,
                HeaderOptions(options.HeaderDetection, beforeHeaderCorrection),
                cancellationToken);
            HeaderDetectionResult afterHeader = HeaderDetector.Detect(
                afterTable,
                beforeTable,
                HeaderOptions(options.HeaderDetection, afterHeaderCorrection),
                cancellationToken);
            progress.Report(
                OpenXmlComparisonProjectProgressStage.InferringKeys,
                child,
                after.Fingerprint.Name,
                "Inferring conservative row keys in sheet '" + after.Fingerprint.Name + "'.");
            bool hasKeyCorrection = corrections.TryGetKeyColumns(
                parent.Source.RevisionId,
                child.RevisionId,
                before.Fingerprint.SheetKey,
                after.Fingerprint.SheetKey,
                out AlignmentKeyColumnsConfirmation keyCorrection);
            KeyInferenceResult key = KeyColumnInference.Infer(
                beforeTable,
                afterTable,
                beforeHeader,
                afterHeader,
                KeyOptions(options.KeyInference, keyCorrection),
                cancellationToken);

            // Exact canonical evidence at the same physical coordinates is stronger than inferred keys, header
            // labels, or content-occurrence anchors. In particular, duplicate rows and duplicate header labels
            // cannot make an unchanged sheet ambiguous. Active reviewer corrections still take precedence so an
            // explicit semantic choice is never silently bypassed by this automatic fast path.
            bool useExactPositionalEvidence = hasExactPositionalEvidence &&
                                              !hasBeforeHeaderCorrection &&
                                              !hasAfterHeaderCorrection &&
                                              !hasKeyCorrection;
            bool headerShapeSafe = true;
            if (!useExactPositionalEvidence &&
                beforeHeader.HasAcceptedHeader != afterHeader.HasAcceptedHeader)
            {
                headerShapeSafe = false;
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.HeaderMismatch,
                    DiagnosticSeverity.Correctable,
                    parent.Source.RevisionId,
                    child.RevisionId,
                    parentSheet.LogicalSheetId,
                    after.Fingerprint.Name,
                    "A header was accepted on only one side of the sheet mapping.",
                    "Use the comparison corrections dialog to confirm Header row or No header independently for " +
                    "both exact sheet occurrences, including the first data row when subheaders are present.",
                    HeaderReasonCodes(beforeHeader, afterHeader),
                    Array.Empty<string>()));
            }
            else if (!useExactPositionalEvidence &&
                     beforeHeader.HasAcceptedHeader &&
                     (beforeHeader.AcceptedHeaderRowIndex.Value != 0 ||
                      afterHeader.AcceptedHeaderRowIndex.Value != 0 ||
                      beforeHeader.DataStartRowIndex != beforeHeader.AcceptedHeaderRowIndex.Value + 1 ||
                      afterHeader.DataStartRowIndex != afterHeader.AcceptedHeaderRowIndex.Value + 1) &&
                     !CanLinkNonLeadingHeaderAutomatically(
                         beforeTable,
                         afterTable,
                         beforeHeader,
                         afterHeader))
            {
                headerShapeSafe = false;
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.NonLeadingHeader,
                    DiagnosticSeverity.Correctable,
                    parent.Source.RevisionId,
                    child.RevisionId,
                    parentSheet.LogicalSheetId,
                    after.Fingerprint.Name,
                    "The detected non-leading header or separated first-data row does not have enough matching " +
                    "structure to retain every pre-data row identity safely.",
                    "Confirm the exact header and first data row on both revisions. Unequal pre-data counts, " +
                    "changed preamble rows, or an automatically detected blank first-data row remain blocked " +
                    "because the engine will not guess their row identities.",
                    HeaderReasonCodes(beforeHeader, afterHeader)
                        .Concat(new[] { "project.nonleading-predata-identity-unproven" }),
                    Array.Empty<string>()));
            }

            if (!headerShapeSafe)
                return new OpenXmlProjectSheetAlignment(
                    parentSheet.LogicalSheetId,
                    before.Fingerprint.SheetKey,
                    after.Fingerprint.SheetKey,
                    before.Fingerprint.Name,
                    after.Fingerprint.Name,
                    beforeHeader,
                    afterHeader,
                    key,
                    null,
                    null,
                    null,
                    null,
                    false);

            AxisAlignmentResult rows;
            AxisAlignmentResult columns;
            if (useExactPositionalEvidence)
            {
                progress.Report(
                    OpenXmlComparisonProjectProgressStage.AligningRows,
                    child,
                    after.Fingerprint.Name,
                    "Preserving exact positional row identity in unchanged sheet '" +
                    after.Fingerprint.Name + "'.");
                rows = CreateExactPositionalAlignment(
                    beforeTable.RowCount,
                    LogicalAxisKind.Row,
                    cancellationToken);
                progress.Report(
                    OpenXmlComparisonProjectProgressStage.AligningColumns,
                    child,
                    after.Fingerprint.Name,
                    "Preserving exact positional column identity in unchanged sheet '" +
                    after.Fingerprint.Name + "'.");
                columns = CreateExactPositionalAlignment(
                    beforeTable.ColumnCount,
                    LogicalAxisKind.Column,
                    cancellationToken);
            }
            else
            {
                progress.Report(
                    OpenXmlComparisonProjectProgressStage.AligningRows,
                    child,
                    after.Fingerprint.Name,
                    "Aligning logical rows in sheet '" + after.Fingerprint.Name + "'.");
                rows = TabularAlignmentEngine.AlignRows(
                    beforeTable,
                    afterTable,
                    beforeHeader,
                    afterHeader,
                    key,
                    options.Alignment,
                    cancellationToken);
                if (beforeHeader.HasAcceptedHeader)
                    rows = IncludeHeaderAndPreDataRows(rows, beforeHeader, afterHeader);
                progress.Report(
                    OpenXmlComparisonProjectProgressStage.AligningColumns,
                    child,
                    after.Fingerprint.Name,
                    "Aligning logical columns in sheet '" + after.Fingerprint.Name + "'.");
                columns = TabularAlignmentEngine.AlignColumns(
                    beforeTable,
                    afterTable,
                    beforeHeader,
                    afterHeader,
                    cancellationToken);
            }

            bool safe = InspectAxis(
                            projectId,
                            parent.Source.RevisionId,
                            child.RevisionId,
                            parentSheet.LogicalSheetId,
                            after.Fingerprint.Name,
                            LogicalAxisKind.Row,
                            rows,
                            options,
                            findings) &
                        InspectAxis(
                            projectId,
                            parent.Source.RevisionId,
                            child.RevisionId,
                            parentSheet.LogicalSheetId,
                            after.Fingerprint.Name,
                            LogicalAxisKind.Column,
                            columns,
                            options,
                            findings);

            cancellationToken.ThrowIfCancellationRequested();
            progress.Report(
                OpenXmlComparisonProjectProgressStage.PropagatingIdentities,
                child,
                after.Fingerprint.Name,
                "Propagating stable row and column identities in sheet '" + after.Fingerprint.Name + "'.");
            LogicalAxisIdentityProjection rowIdentities = LogicalAxisIdentityProjector.Propagate(
                parentSheet.Rows,
                child.RevisionId,
                afterTable.RowCount,
                rows);
            cancellationToken.ThrowIfCancellationRequested();
            LogicalAxisIdentityProjection columnIdentities = LogicalAxisIdentityProjector.Propagate(
                parentSheet.Columns,
                child.RevisionId,
                afterTable.ColumnCount,
                columns);
            cancellationToken.ThrowIfCancellationRequested();
            if (!rowIdentities.IsFullyResolved || !columnIdentities.IsFullyResolved)
            {
                safe = false;
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.AxisIdentityIncomplete,
                    DiagnosticSeverity.Correctable,
                    parent.Source.RevisionId,
                    child.RevisionId,
                    parentSheet.LogicalSheetId,
                    after.Fingerprint.Name,
                    "One or more physical rows or columns have no confirmed logical identity.",
                    "Confirm the exact sheet, header/data region, and key columns in the comparison corrections " +
                    "dialog. Remaining duplicate or competing axis blocks require a comparison copy with unique evidence.",
                    rowIdentities.Ambiguities.SelectMany(item => item.Reasons.Select(reason => reason.Code))
                        .Concat(columnIdentities.Ambiguities.SelectMany(item => item.Reasons.Select(reason => reason.Code))),
                    Array.Empty<string>()));
            }

            return new OpenXmlProjectSheetAlignment(
                parentSheet.LogicalSheetId,
                before.Fingerprint.SheetKey,
                after.Fingerprint.SheetKey,
                before.Fingerprint.Name,
                after.Fingerprint.Name,
                beforeHeader,
                afterHeader,
                key,
                rows,
                columns,
                rowIdentities,
                columnIdentities,
                safe && rowIdentities.IsFullyResolved && columnIdentities.IsFullyResolved);
        }

        private static bool HasExactPositionalCellEvidence(
            OpenXmlSheetFingerprintEvidence before,
            CanonicalGridProjection beforeGrid,
            OpenXmlSheetFingerprintEvidence after,
            CanonicalGridProjection afterGrid,
            CancellationToken cancellationToken)
        {
            if (!before.CanUseCanonicalCellEvidence ||
                !after.CanUseCanonicalCellEvidence ||
                !before.Source.CanProjectExactStoredFormulaValueEvidence ||
                !after.Source.CanProjectExactStoredFormulaValueEvidence)
                return false;

            AlignmentTable beforeTable = beforeGrid.Table;
            AlignmentTable afterTable = afterGrid.Table;
            if (beforeGrid.FirstPhysicalRow != afterGrid.FirstPhysicalRow ||
                beforeGrid.FirstPhysicalColumn != afterGrid.FirstPhysicalColumn ||
                beforeTable.RowCount != afterTable.RowCount ||
                beforeTable.ColumnCount != afterTable.ColumnCount)
                return false;

            long beforeStoredCells = 0;
            int processed = 0;
            foreach (CanonicalCellState beforeCell in beforeGrid.Source.Cells.Values)
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(
                        beforeCell.SubjectKey,
                        out int row,
                        out int column))
                    return false;
                if (!IsInsideProjection(beforeGrid, row, column)) continue;

                beforeStoredCells++;
                if (!afterGrid.TryGetPhysicalCell(row, column, out CanonicalCellState afterCell) ||
                    !HasSameExactCellEvidence(beforeCell, afterCell))
                    return false;
            }

            long afterStoredCells = 0;
            processed = 0;
            foreach (CanonicalCellState afterCell in afterGrid.Source.Cells.Values)
            {
                if ((processed++ & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(
                        afterCell.SubjectKey,
                        out int row,
                        out int column))
                    return false;
                if (IsInsideProjection(afterGrid, row, column)) afterStoredCells++;
            }

            cancellationToken.ThrowIfCancellationRequested();
            return beforeStoredCells == afterStoredCells;
        }

        private static bool IsInsideProjection(CanonicalGridProjection grid, int row, int column)
        {
            if (grid.Table.RowCount == 0 || grid.Table.ColumnCount == 0) return false;
            return row >= grid.FirstPhysicalRow &&
                   row < grid.FirstPhysicalRow + grid.Table.RowCount &&
                   column >= grid.FirstPhysicalColumn &&
                   column < grid.FirstPhysicalColumn + grid.Table.ColumnCount;
        }

        private static bool HasSameExactCellEvidence(
            CanonicalCellState before,
            CanonicalCellState after)
        {
            if (before.HasFormula != after.HasFormula ||
                !string.Equals(before.FormulaText, after.FormulaText, StringComparison.Ordinal) ||
                !string.Equals(before.FormulaMetadata, after.FormulaMetadata, StringComparison.Ordinal) ||
                !before.ConstantValue.Equals(after.ConstantValue))
                return false;

            if (ReferenceEquals(before.CachedResult, after.CachedResult)) return true;
            return before.CachedResult != null && before.CachedResult.Equals(after.CachedResult);
        }

        private static AxisAlignmentResult CreateExactPositionalAlignment(
            int count,
            LogicalAxisKind axis,
            CancellationToken cancellationToken)
        {
            string axisName = axis == LogicalAxisKind.Row ? "row" : "column";
            var reason = new AlignmentReason(
                "axis.exact-evidence-position",
                "The " + axisName +
                " stayed at the same physical position inside a grid with identical stored values, saved " +
                "calculated values, and formulas.");
            var links = new List<AxisAlignmentLink>(count);
            for (int index = 0; index < count; index++)
            {
                if ((index & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                links.Add(new AxisAlignmentLink(
                    index,
                    index,
                    AxisAlignmentKind.Matched,
                    AlignmentConfidence.Strong,
                    AxisAlignmentStrategy.Positional,
                    new[] { reason }));
            }

            cancellationToken.ThrowIfCancellationRequested();
            return new AxisAlignmentResult(
                AxisAlignmentStrategy.Positional,
                links,
                Array.Empty<AxisAlignmentAmbiguity>(),
                new[] { reason });
        }

        private static bool InspectAxis(
            string projectId,
            string beforeRevisionId,
            string afterRevisionId,
            string logicalSheetId,
            string sheetName,
            LogicalAxisKind axis,
            AxisAlignmentResult alignment,
            OpenXmlComparisonProjectOptions options,
            ICollection<OpenXmlProjectFinding> findings)
        {
            bool safe = true;
            if (alignment.HasAmbiguity)
            {
                safe = false;
                findings.Add(Finding(
                    projectId,
                    axis == LogicalAxisKind.Row
                        ? OpenXmlProjectFindingCodes.RowAlignmentAmbiguous
                        : OpenXmlProjectFindingCodes.ColumnAlignmentAmbiguous,
                    DiagnosticSeverity.Correctable,
                    beforeRevisionId,
                    afterRevisionId,
                    logicalSheetId,
                    sheetName,
                    "Duplicate or competing " + AxisName(axis) + " evidence produced " +
                    alignment.Ambiguities.Count.ToString(CultureInfo.InvariantCulture) +
                    " unresolved alignment block(s).",
                    "Confirm the exact header/data region and key columns in the comparison corrections dialog. " +
                    "Duplicate key tuples or competing " + AxisName(axis) + " blocks remain blocked and require " +
                    "unique evidence in a comparison copy or the exact two-file preview.",
                    alignment.Ambiguities.SelectMany(item => item.Reasons.Select(reason => reason.Code)),
                    Array.Empty<string>()));
            }

            if (options.RequireAtLeastMediumAxisConfidence &&
                alignment.Links.Any(item => item.Kind == AxisAlignmentKind.Matched &&
                                            item.Confidence < AlignmentConfidence.Medium))
            {
                safe = false;
                findings.Add(Finding(
                    projectId,
                    axis == LogicalAxisKind.Row
                        ? OpenXmlProjectFindingCodes.WeakRowAlignment
                        : OpenXmlProjectFindingCodes.WeakColumnAlignment,
                    DiagnosticSeverity.Correctable,
                    beforeRevisionId,
                    afterRevisionId,
                    logicalSheetId,
                    sheetName,
                    "At least one " + AxisName(axis) + " match relies only on weak positional evidence.",
                    "Confirm the exact header/data region and key columns in the comparison corrections dialog. " +
                    "If the remaining match is still only positional, provide stronger evidence in a comparison copy " +
                    "or use the exact two-file preview.",
                    alignment.Links
                        .Where(item => item.Kind == AxisAlignmentKind.Matched &&
                                       item.Confidence < AlignmentConfidence.Medium)
                        .SelectMany(item => item.Reasons.Select(reason => reason.Code)),
                    Array.Empty<string>()));
            }
            return safe;
        }

        private static AxisAlignmentResult IncludeHeaderAndPreDataRows(
            AxisAlignmentResult dataRows,
            HeaderDetectionResult beforeHeader,
            HeaderDetectionResult afterHeader)
        {
            if (!beforeHeader.AcceptedHeaderRowIndex.HasValue ||
                !afterHeader.AcceptedHeaderRowIndex.HasValue)
                throw new InvalidOperationException("Both header rows must be accepted before structural links are added.");
            if (beforeHeader.AcceptedHeaderRowIndex.Value != afterHeader.AcceptedHeaderRowIndex.Value ||
                beforeHeader.DataStartRowIndex != afterHeader.DataStartRowIndex)
                throw new InvalidOperationException(
                    "Confirmed preamble/subheader row counts must match before positional identities are created.");

            AlignmentConfidence confidence = beforeHeader.Decision == HeaderDecision.ConfirmedHeader &&
                                             afterHeader.Decision == HeaderDecision.ConfirmedHeader
                ? AlignmentConfidence.Confirmed
                : AlignmentConfidence.Strong;
            bool confirmed = confidence == AlignmentConfidence.Confirmed;
            int headerIndex = beforeHeader.AcceptedHeaderRowIndex.Value;
            int dataStart = beforeHeader.DataStartRowIndex;
            var links = new List<AxisAlignmentLink>(dataRows.Links.Count + dataStart);
            for (int index = 0; index < dataStart; index++)
            {
                string role = index < headerIndex
                    ? "preamble"
                    : index == headerIndex
                        ? "header"
                        : "subheader";
                links.Add(new AxisAlignmentLink(
                    index,
                    index,
                    AxisAlignmentKind.Matched,
                    confidence,
                    AxisAlignmentStrategy.Positional,
                    new[]
                    {
                        new AlignmentReason(
                            "project." + (confirmed ? "confirmed-" : "strong-evidence-") +
                            role + "-row-match",
                            confirmed
                                ? "The equal confirmed pre-data shape retains this " + role +
                                  " row at the same table-relative position on both revisions."
                                : "Matching strong header detection and exact preamble evidence retain this " +
                                  role + " row at the same table-relative position on both revisions."),
                    }));
            }
            links.AddRange(dataRows.Links);
            return new AxisAlignmentResult(dataRows.Strategy, links, dataRows.Ambiguities, dataRows.Reasons);
        }

        private static bool CanLinkNonLeadingHeaderAutomatically(
            AlignmentTable before,
            AlignmentTable after,
            HeaderDetectionResult beforeHeader,
            HeaderDetectionResult afterHeader)
        {
            if (!beforeHeader.HasAcceptedHeader || !afterHeader.HasAcceptedHeader ||
                beforeHeader.AcceptedHeaderRowIndex.Value != afterHeader.AcceptedHeaderRowIndex.Value ||
                beforeHeader.DataStartRowIndex != afterHeader.DataStartRowIndex)
                return false;

            if (beforeHeader.Decision == HeaderDecision.ConfirmedHeader &&
                afterHeader.Decision == HeaderDecision.ConfirmedHeader)
                return true;

            int headerIndex = beforeHeader.AcceptedHeaderRowIndex.Value;
            return HasEqualExactPreambleEvidence(before, after, headerIndex) &&
                   HasPopulatedRow(before, beforeHeader.DataStartRowIndex) &&
                   HasPopulatedRow(after, afterHeader.DataStartRowIndex);
        }

        private static bool HasEqualExactPreambleEvidence(
            AlignmentTable before,
            AlignmentTable after,
            int headerIndex)
        {
            int columnCount = Math.Max(before.ColumnCount, after.ColumnCount);
            for (int row = 0; row < headerIndex; row++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    bool beforePresent = before.TryGetCell(row, column, out string beforeValue);
                    bool afterPresent = after.TryGetCell(row, column, out string afterValue);
                    if (beforePresent != afterPresent ||
                        !string.Equals(beforeValue, afterValue, StringComparison.Ordinal))
                        return false;
                }
            }
            return true;
        }

        private static bool HasPopulatedRow(AlignmentTable table, int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= table.RowCount) return false;
            return table.Rows[rowIndex].Values.Any(value => !string.IsNullOrWhiteSpace(value));
        }

        private static RevisionSnapshot CreateAlignedWorkbookSnapshot(
            string revisionId,
            IEnumerable<OpenXmlAlignedSheetRevision> sheets,
            CancellationToken cancellationToken)
        {
            var cells = new List<CanonicalCellState>();
            foreach (OpenXmlAlignedSheetRevision sheet in sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyDictionary<int, string> rows = ToPhysicalIdentityMap(sheet.Grid, sheet.Rows, true);
                IReadOnlyDictionary<int, string> columns = ToPhysicalIdentityMap(sheet.Grid, sheet.Columns, false);
                cells.AddRange(sheet.Grid.CreateAlignedCells(
                        sheet.LogicalSheetId,
                        rows,
                        columns,
                        cancellationToken)
                    .Select(item => item.Cell));
            }
            cancellationToken.ThrowIfCancellationRequested();
            return new RevisionSnapshot(revisionId, cells);
        }

        private static IReadOnlyDictionary<int, string> ToPhysicalIdentityMap(
            CanonicalGridProjection grid,
            LogicalAxisIdentityProjection identities,
            bool rows)
        {
            var result = new Dictionary<int, string>();
            foreach (KeyValuePair<int, string> pair in identities.IdentityByPhysicalIndex)
            {
                int physical = rows
                    ? grid.PhysicalRowForTableIndex(pair.Key)
                    : grid.PhysicalColumnForTableIndex(pair.Key);
                result.Add(physical, pair.Value);
            }
            return new ReadOnlyDictionary<int, string>(result);
        }

        private static HeaderDetectionOptions HeaderOptions(
            HeaderDetectionOptions source,
            AlignmentHeaderRowConfirmation correction)
        {
            if (correction == null) return source;
            return new HeaderDetectionOptions
            {
                MaxScanRows = source.MaxScanRows,
                MaxEvidenceRows = source.MaxEvidenceRows,
                CommentPrefix = source.CommentPrefix,
                StrongScoreThreshold = source.StrongScoreThreshold,
                StrongRunnerUpMargin = source.StrongRunnerUpMargin,
                AmbiguousScoreThreshold = source.AmbiguousScoreThreshold,
                ConfirmedNoHeader = correction.Decision == AlignmentHeaderConfirmationDecision.NoHeader,
                ConfirmedHeaderRowIndex = correction.Row?.TableIndex,
                ConfirmedDataStartRowIndex = correction.FirstDataRow?.TableIndex,
            };
        }

        private static KeyInferenceOptions KeyOptions(
            KeyInferenceOptions source,
            AlignmentKeyColumnsConfirmation correction)
        {
            if (correction == null) return source;
            return new KeyInferenceOptions
            {
                MaxCompositeColumns = Math.Max(source.MaxCompositeColumns, correction.BeforeColumns.Count),
                MaxCandidateHeaderColumns = source.MaxCandidateHeaderColumns,
                StrongScoreThreshold = source.StrongScoreThreshold,
                StrongRunnerUpMargin = source.StrongRunnerUpMargin,
                MinimumCoverage = source.MinimumCoverage,
                MinimumUniqueness = source.MinimumUniqueness,
                MinimumCrossRevisionOverlap = source.MinimumCrossRevisionOverlap,
                ConfirmedBeforeColumns = new ReadOnlyCollection<int>(
                    correction.BeforeColumns.Select(item => item.TableIndex).ToList()),
                ConfirmedAfterColumns = new ReadOnlyCollection<int>(
                    correction.AfterColumns.Select(item => item.TableIndex).ToList()),
            };
        }

        private static void ValidateAlignmentCorrections(
            IReadOnlyList<OpenXmlProjectRevision> revisions,
            AlignmentCorrectionGeneration corrections,
            OpenXmlComparisonProjectOptions options,
            CancellationToken cancellationToken)
        {
            var revisionsById = revisions.ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            var fingerprintSets = new Dictionary<string, OpenXmlWorkbookSheetFingerprintSet>(StringComparer.Ordinal);
            var grids = new Dictionary<string, CanonicalGridProjection>(StringComparer.Ordinal);

            OpenXmlWorkbookSheetFingerprintSet Fingerprints(string revisionId)
            {
                if (!revisionsById.TryGetValue(revisionId, out OpenXmlProjectRevision revision))
                    throw new InvalidOperationException(
                        "An alignment confirmation references stale or unknown revision occurrence '" +
                        revisionId + "'.");
                if (!fingerprintSets.TryGetValue(revisionId, out OpenXmlWorkbookSheetFingerprintSet result))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    result = OpenXmlSheetFingerprintBuilder.Build(revision.Snapshot, cancellationToken);
                    fingerprintSets.Add(revisionId, result);
                }
                return result;
            }

            OpenXmlSheetFingerprintEvidence Sheet(string revisionId, string physicalSheetKey)
            {
                OpenXmlSheetFingerprintEvidence evidence = Fingerprints(revisionId).Sheets.SingleOrDefault(item =>
                    string.Equals(item.Fingerprint.SheetKey, physicalSheetKey, StringComparison.Ordinal));
                if (evidence == null)
                    throw new InvalidOperationException(
                        "An alignment confirmation references stale physical sheet key '" + physicalSheetKey +
                        "' in revision occurrence '" + revisionId + "'.");
                return evidence;
            }

            CanonicalGridProjection Grid(string revisionId, string physicalSheetKey)
            {
                string key = revisionId.Length.ToString(CultureInfo.InvariantCulture) + ":" + revisionId +
                             physicalSheetKey;
                if (!grids.TryGetValue(key, out CanonicalGridProjection grid))
                {
                    OpenXmlSheetFingerprintEvidence evidence = Sheet(revisionId, physicalSheetKey);
                    cancellationToken.ThrowIfCancellationRequested();
                    grid = CanonicalGridProjectionBuilder.Build(
                        evidence.Source.Snapshot,
                        options.GridProjection,
                        cancellationToken);
                    grids.Add(key, grid);
                }
                return grid;
            }

            void RequireEdge(string parentRevisionId, string childRevisionId)
            {
                if (!revisionsById.TryGetValue(childRevisionId, out OpenXmlProjectRevision child) ||
                    !revisionsById.ContainsKey(parentRevisionId) ||
                    !string.Equals(child.ParentRevisionId, parentRevisionId, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "An alignment confirmation references stale or non-edge revision occurrences '" +
                        parentRevisionId + "' -> '" + childRevisionId + "'.");
            }

            foreach (AlignmentSheetPairConfirmation confirmation in corrections.SheetPairs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RequireEdge(confirmation.ParentRevisionId, confirmation.ChildRevisionId);
                Sheet(confirmation.ParentRevisionId, confirmation.BeforeSheetKey);
                Sheet(confirmation.ChildRevisionId, confirmation.AfterSheetKey);
            }

            foreach (AlignmentSheetUnmatchedConfirmation confirmation in corrections.UnmatchedSheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RequireEdge(confirmation.ParentRevisionId, confirmation.ChildRevisionId);
                Sheet(confirmation.ParentRevisionId, confirmation.BeforeSheetKey);
            }

            foreach (AlignmentHeaderRowConfirmation confirmation in corrections.HeaderRows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CanonicalGridProjection grid = Grid(confirmation.RevisionId, confirmation.PhysicalSheetKey);
                if (confirmation.Decision == AlignmentHeaderConfirmationDecision.NoHeader) continue;
                ValidateGridRowReference(grid, confirmation.Row, false, "header");
                if (confirmation.FirstDataRow != null)
                    ValidateGridRowReference(grid, confirmation.FirstDataRow, true, "first data");
            }

            foreach (AlignmentKeyColumnsConfirmation confirmation in corrections.KeyColumns)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RequireEdge(confirmation.ParentRevisionId, confirmation.ChildRevisionId);
                CanonicalGridProjection before = Grid(
                    confirmation.ParentRevisionId,
                    confirmation.BeforeSheetKey);
                CanonicalGridProjection after = Grid(
                    confirmation.ChildRevisionId,
                    confirmation.AfterSheetKey);
                foreach (AlignmentGridIndexReference column in confirmation.BeforeColumns)
                    ValidateGridColumnReference(before, column, "Before key");
                foreach (AlignmentGridIndexReference column in confirmation.AfterColumns)
                    ValidateGridColumnReference(after, column, "After key");
            }
        }

        private static void ValidateGridRowReference(
            CanonicalGridProjection grid,
            AlignmentGridIndexReference reference,
            bool allowJustAfter,
            string role)
        {
            int upperBound = allowJustAfter ? grid.Table.RowCount : grid.Table.RowCount - 1;
            if (reference.TableIndex < 0 || reference.TableIndex > upperBound)
                throw new InvalidOperationException(
                    "The confirmed " + role + " table row is outside the exact sheet grid.");
            int expected = checked(grid.FirstPhysicalRow + reference.TableIndex);
            if (reference.PhysicalIndex != expected)
                throw new InvalidOperationException(
                    "The confirmed " + role + " row provenance is stale: table index " +
                    reference.TableIndex.ToString(CultureInfo.InvariantCulture) + " now maps to physical row " +
                    expected.ToString(CultureInfo.InvariantCulture) + ", not " +
                    reference.PhysicalIndex.ToString(CultureInfo.InvariantCulture) + ".");
        }

        private static void ValidateGridColumnReference(
            CanonicalGridProjection grid,
            AlignmentGridIndexReference reference,
            string role)
        {
            if (reference.TableIndex < 0 || reference.TableIndex >= grid.Table.ColumnCount)
                throw new InvalidOperationException(
                    "The confirmed " + role + " table column is outside the exact sheet grid.");
            int expected = checked(grid.FirstPhysicalColumn + reference.TableIndex);
            if (reference.PhysicalIndex != expected)
                throw new InvalidOperationException(
                    "The confirmed " + role + " column provenance is stale: table index " +
                    reference.TableIndex.ToString(CultureInfo.InvariantCulture) + " now maps to physical column " +
                    expected.ToString(CultureInfo.InvariantCulture) + ", not " +
                    reference.PhysicalIndex.ToString(CultureInfo.InvariantCulture) + ".");
        }

        internal static OpenXmlWorkbookSheetMappingResult ApplyConfirmedSheetMappings(
            OpenXmlWorkbookSheetMappingResult source,
            IReadOnlyList<AlignmentSheetPairConfirmation> confirmations) =>
            ApplyConfirmedSheetMappings(
                source,
                confirmations,
                Array.Empty<AlignmentSheetUnmatchedConfirmation>());

        internal static OpenXmlWorkbookSheetMappingResult ApplyConfirmedSheetMappings(
            OpenXmlWorkbookSheetMappingResult source,
            IReadOnlyList<AlignmentSheetPairConfirmation> confirmations,
            IReadOnlyList<AlignmentSheetUnmatchedConfirmation> unmatchedConfirmations)
        {
            confirmations = confirmations ?? Array.Empty<AlignmentSheetPairConfirmation>();
            unmatchedConfirmations = unmatchedConfirmations ?? Array.Empty<AlignmentSheetUnmatchedConfirmation>();
            if (confirmations.Count == 0 && unmatchedConfirmations.Count == 0) return source;

            var beforeByKey = source.Before.Fingerprints.ToDictionary(item => item.SheetKey, StringComparer.Ordinal);
            var afterByKey = source.After.Fingerprints.ToDictionary(item => item.SheetKey, StringComparer.Ordinal);
            var confirmedBefore = new HashSet<string>(
                confirmations.Select(item => item.BeforeSheetKey),
                StringComparer.Ordinal);
            var confirmedAfter = new HashSet<string>(
                confirmations.Select(item => item.AfterSheetKey),
                StringComparer.Ordinal);
            var confirmedUnmatchedBefore = new HashSet<string>(
                unmatchedConfirmations.Select(item => item.BeforeSheetKey),
                StringComparer.Ordinal);
            if (confirmedBefore.Overlaps(confirmedUnmatchedBefore))
                throw new InvalidOperationException(
                    "One source sheet cannot be both explicitly paired and explicitly unmatched on the same edge.");
            foreach (AlignmentSheetPairConfirmation confirmation in confirmations)
            {
                if (!beforeByKey.ContainsKey(confirmation.BeforeSheetKey) ||
                    !afterByKey.ContainsKey(confirmation.AfterSheetKey))
                    throw new InvalidOperationException(
                        "A sheet-pair confirmation references a stale exact sheet occurrence.");
            }
            foreach (AlignmentSheetUnmatchedConfirmation confirmation in unmatchedConfirmations)
                if (!beforeByKey.ContainsKey(confirmation.BeforeSheetKey))
                    throw new InvalidOperationException(
                        "An unmatched-sheet confirmation references a stale exact source occurrence.");

            var claimedBefore = new HashSet<string>(confirmedBefore, StringComparer.Ordinal);
            claimedBefore.UnionWith(confirmedUnmatchedBefore);
            var accepted = new List<ResolvedSheetMapping>();
            foreach (ResolvedSheetMapping inferred in source.Mapping.Mappings)
            {
                if (claimedBefore.Contains(inferred.Before.SheetKey) ||
                    confirmedAfter.Contains(inferred.After.SheetKey))
                    continue;
                accepted.Add(inferred);
            }
            foreach (AlignmentSheetPairConfirmation confirmation in confirmations)
                accepted.Add(new ResolvedSheetMapping(
                    beforeByKey[confirmation.BeforeSheetKey],
                    afterByKey[confirmation.AfterSheetKey],
                    SheetMappingKind.UserConfirmed,
                    1D,
                    new[] { SheetMappingReasonCodes.UserConfirmed }));

            var acceptedBefore = new HashSet<string>(
                accepted.Select(item => item.Before.SheetKey),
                StringComparer.Ordinal);
            var acceptedAfter = new HashSet<string>(
                accepted.Select(item => item.After.SheetKey),
                StringComparer.Ordinal);
            var originalUnresolved = source.Mapping.UnresolvedBeforeSheets.ToDictionary(
                item => item.Before.SheetKey,
                StringComparer.Ordinal);
            var originalResolved = source.Mapping.Mappings.ToDictionary(
                item => item.Before.SheetKey,
                StringComparer.Ordinal);
            var unresolved = new List<UnresolvedSheetMapping>();
            foreach (WorkbookSheetFingerprint before in source.Before.Fingerprints
                         .Where(item => !acceptedBefore.Contains(item.SheetKey) &&
                                        !confirmedUnmatchedBefore.Contains(item.SheetKey)))
            {
                var candidates = new List<SheetMappingCandidate>();
                var reasons = new List<string>();
                if (originalUnresolved.TryGetValue(before.SheetKey, out UnresolvedSheetMapping prior))
                {
                    candidates.AddRange(prior.Candidates.Where(item =>
                        !acceptedAfter.Contains(item.After.SheetKey)));
                    reasons.AddRange(prior.ReasonCodes);
                }
                if (originalResolved.TryGetValue(before.SheetKey, out ResolvedSheetMapping displaced) &&
                    !acceptedAfter.Contains(displaced.After.SheetKey))
                    candidates.Add(new SheetMappingCandidate(
                        displaced.Before,
                        displaced.After,
                        displaced.Confidence,
                        displaced.ReasonCodes));
                reasons.Add(SheetMappingReasonCodes.ConfirmedTargetClaimed);
                unresolved.Add(new UnresolvedSheetMapping(
                    before,
                    candidates
                        .GroupBy(item => item.After.SheetKey, StringComparer.Ordinal)
                        .Select(group => group.OrderByDescending(item => item.Confidence).First())
                        .OrderByDescending(item => item.Confidence)
                        .ThenBy(item => item.After.Name, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(item => item.After.SheetKey, StringComparer.Ordinal)
                        .ToList(),
                    reasons.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()));
            }

            accepted.Sort((left, right) => string.CompareOrdinal(left.Before.SheetKey, right.Before.SheetKey));
            unresolved.Sort((left, right) => string.CompareOrdinal(left.Before.SheetKey, right.Before.SheetKey));
            var unmatchedBefore = confirmedUnmatchedBefore
                .Select(key => beforeByKey[key])
                .OrderBy(item => item.SheetKey, StringComparer.Ordinal)
                .ToList();
            var unmatchedAfter = source.After.Fingerprints
                .Where(item => !acceptedAfter.Contains(item.SheetKey))
                .OrderBy(item => item.SheetKey, StringComparer.Ordinal)
                .ToList();
            return new OpenXmlWorkbookSheetMappingResult(
                source.Before,
                source.After,
                new WorkbookSheetMappingResult(accepted, unresolved, unmatchedBefore, unmatchedAfter));
        }

        private static bool ValidateEvidence(
            string projectId,
            string revisionId,
            string beforeRevisionId,
            string logicalSheetId,
            OpenXmlSheetFingerprintEvidence evidence,
            ICollection<OpenXmlProjectFinding> findings)
        {
            if (evidence.Source.Kind != OpenXmlPhysicalSheetKind.Worksheet ||
                evidence.Fingerprint.Kind != WorkbookSheetKind.Worksheet)
            {
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.UnsupportedSheet,
                    DiagnosticSeverity.ArtifactBlocking,
                    beforeRevisionId,
                    revisionId,
                    logicalSheetId,
                    evidence.Fingerprint.Name,
                    "The physical sheet kind is not supported by canonical worksheet alignment.",
                    "Convert the sheet to a worksheet or exclude it from the comparison project.",
                    ExactCellEvidenceReasonCodes(evidence),
                    Array.Empty<string>()));
                return false;
            }
            if (!evidence.CanUseCanonicalCellEvidence)
            {
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.IncompleteCellEvidence,
                    DiagnosticSeverity.ArtifactBlocking,
                    beforeRevisionId,
                    revisionId,
                    logicalSheetId,
                    evidence.Fingerprint.Name,
                    "Power Tools cannot model complete stored-value, saved-calculated-value, and formula evidence " +
                    "for this sheet.",
                    "Review the listed coverage codes and the source sheet directly. Do not rely on this diagnostic " +
                    "for cell-level completeness until the unsupported construct is removed or supported.",
                    ExactCellEvidenceReasonCodes(evidence),
                    Array.Empty<string>()));
                return false;
            }
            if (!evidence.Source.CanAssertExactCellEquality)
            {
                string[] dynamicCoverage = evidence.Source.CoverageMessages
                    .Where(item => item.AffectsExactCellComparison)
                    .Select(item => item.Code)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(item => item, StringComparer.Ordinal)
                    .ToArray();
                findings.Add(Finding(
                    projectId,
                    OpenXmlProjectFindingCodes.IncompleteCellEvidence,
                    DiagnosticSeverity.Warning,
                    beforeRevisionId,
                    revisionId,
                    logicalSheetId,
                    evidence.Fingerprint.Name,
                    "Stored values, saved calculated values, and exact formulas are available for review. " +
                    "Spill metadata is retained only to warn that dynamic-array recalculation and spill semantics " +
                    "are not modeled.",
                    "Review the exact stored changes in this diagnostic report. Merge remains blocked until Power " +
                    "Tools can validate the dynamic-array spill semantics.",
                    dynamicCoverage,
                    Array.Empty<string>()));
            }
            return true;
        }

        private static string[] ExactCellEvidenceReasonCodes(OpenXmlSheetFingerprintEvidence evidence)
        {
            var codes = evidence.Source.CoverageMessages
                .Where(item => item.AffectsExactCellComparison)
                .Select(item => item.Code)
                .ToList();
            if (evidence.CoverageCodes.Contains(
                    OpenXmlSheetFingerprintCoverageCodes.InvalidCanonicalCellReference,
                    StringComparer.Ordinal))
            {
                codes.Add(OpenXmlSheetFingerprintCoverageCodes.InvalidCanonicalCellReference);
            }
            return codes
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .ToArray();
        }

        private static OpenXmlProjectFinding GridFailure(
            string projectId,
            string beforeRevisionId,
            string afterRevisionId,
            string logicalSheetId,
            OpenXmlSheetFingerprintEvidence evidence,
            Exception exception) =>
            Finding(
                projectId,
                OpenXmlProjectFindingCodes.GridProjectionFailed,
                DiagnosticSeverity.ArtifactBlocking,
                beforeRevisionId,
                afterRevisionId,
                logicalSheetId,
                evidence.Fingerprint.Name,
                "The bounded alignment grid could not be created: " + exception.Message,
                "Reduce the comparable used range or configure larger safe projection limits.",
                new[] { exception.GetType().FullName },
                Array.Empty<string>());

        private static OpenXmlProjectSheetAlignment EmptyAlignment(
            OpenXmlAlignedSheetRevision parentSheet,
            OpenXmlSheetFingerprintEvidence before,
            OpenXmlSheetFingerprintEvidence after) =>
            new OpenXmlProjectSheetAlignment(
                parentSheet.LogicalSheetId,
                before.Fingerprint.SheetKey,
                after.Fingerprint.SheetKey,
                before.Fingerprint.Name,
                after.Fingerprint.Name,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                false);

        private static OpenXmlProjectFinding Finding(
            string projectId,
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
            string findingId = StableId.Compute(
                "ooxml-project-finding-v1",
                projectId,
                code,
                beforeRevisionId ?? "",
                afterRevisionId ?? "",
                logicalSheetId ?? "",
                sheetName ?? "",
                message ?? "");
            return new OpenXmlProjectFinding(
                findingId,
                code,
                severity,
                beforeRevisionId,
                afterRevisionId,
                logicalSheetId,
                sheetName,
                message,
                suggestedAction,
                Distinct(reasonCodes),
                Distinct(candidateSheetNames));
        }

        private static IEnumerable<string> HeaderReasonCodes(
            HeaderDetectionResult before,
            HeaderDetectionResult after) =>
            before.Evidence.Select(item => item.Code)
                .Concat(after.Evidence.Select(item => item.Code));

        private static IReadOnlyList<string> Distinct(IEnumerable<string> values) =>
            new ReadOnlyCollection<string>(
                (values ?? Array.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToList());

        private static string ParentLogicalSheetId(
            IReadOnlyDictionary<string, OpenXmlAlignedSheetRevision> parentSheets,
            string physicalSheetKey) =>
            parentSheets.TryGetValue(physicalSheetKey, out OpenXmlAlignedSheetRevision sheet)
                ? sheet.LogicalSheetId
                : "";

        private static string AxisName(LogicalAxisKind axis) =>
            axis == LogicalAxisKind.Row ? "row" : "column";

        private static void ValidateOptions(OpenXmlComparisonProjectOptions options)
        {
            if (options.SheetMapping == null) throw new ArgumentNullException(nameof(options.SheetMapping));
            if (options.HeaderDetection == null) throw new ArgumentNullException(nameof(options.HeaderDetection));
            if (options.KeyInference == null) throw new ArgumentNullException(nameof(options.KeyInference));
            if (options.Alignment == null) throw new ArgumentNullException(nameof(options.Alignment));
            if (options.GridProjection == null) throw new ArgumentNullException(nameof(options.GridProjection));
            if (options.ExactComparison == null) throw new ArgumentNullException(nameof(options.ExactComparison));
        }

        private static List<OpenXmlProjectRevision> ValidateAndOrder(
            IEnumerable<OpenXmlProjectRevision> revisions)
        {
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));
            var ordered = new List<OpenXmlProjectRevision>(revisions);
            if (ordered.Count < 2)
                throw new ArgumentException("At least two revisions are required.", nameof(revisions));
            if (ordered.Any(item => item == null))
                throw new ArgumentException("The revision collection cannot contain null entries.", nameof(revisions));
            if (ordered.Select(item => item.RevisionId).Distinct(StringComparer.Ordinal).Count() != ordered.Count)
                throw new ArgumentException("Revision occurrence IDs must be unique.", nameof(revisions));
            if (ordered.Select(item => item.Order).Distinct().Count() != ordered.Count)
                throw new ArgumentException("Revision orders must be unique.", nameof(revisions));
            if (ordered.Count(item => item.ParentRevisionId.Length == 0) != 1)
                throw new ArgumentException("The revision tree must contain exactly one root.", nameof(revisions));
            if (ordered.Count(item => item.IsMaster) != 1)
                throw new ArgumentException("The revision tree must contain exactly one master.", nameof(revisions));
            if (ordered.Count(item => item.IsTrustedBaseline) > 1)
                throw new ArgumentException("At most one trusted baseline is allowed.", nameof(revisions));

            ordered.Sort((left, right) => left.Order.CompareTo(right.Order));
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (OpenXmlProjectRevision revision in ordered)
            {
                if (revision.ParentRevisionId.Length > 0 && !seen.Contains(revision.ParentRevisionId))
                    throw new ArgumentException(
                        "Each parent must exist and precede its child in revision order.",
                        nameof(revisions));
                seen.Add(revision.RevisionId);
            }
            return ordered;
        }

        private static void Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A value is required.", name);
        }
    }
}
