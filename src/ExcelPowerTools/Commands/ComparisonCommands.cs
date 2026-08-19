using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.Comparison.Adapters;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Inventory;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Interop;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Bounded, exact two-file comparison preview. This is intentionally separate from the future comparison
    /// project service: it exposes the trustworthy pairwise engine without implying that alignment, grouping,
    /// multi-revision history, snapshots, decisions, or merge are already available.
    /// </summary>
    public static class ComparisonCommands
    {
        internal const string SupportedInputFilter =
            "Comparison files|*.xlsx;*.xlsm;*.xlsb;*.csv;*.tsv;*.txt|" +
            "Excel workbooks|*.xlsx;*.xlsm;*.xlsb|" +
            "Delimited files|*.csv;*.tsv;*.txt|All files|*.*";

        internal const string ReviewWorkbookFilter =
            "Power Tools review workbooks|*.xlsx|Excel workbooks|*.xlsx";

        // The current OOXML adapter owns an in-memory copy, so the UI must not admit inputs beyond its
        // measured safety envelope. The same ceiling is applied to delimited preview inputs until a streaming
        // comparison path exists.
        private const long MaxPreviewInputBytes = OpenXmlWorkbookSnapshotAdapter.MaxSourcePackageBytes;

        [ExcelCommand(
            Name = "PT_CompareExactPreview",
            MenuName = "Power Tools",
            MenuText = "Compare files (exact preview)...")]
        public static void StartExactPreview()
        {
            string[] selected = Dialogs.OpenFiles(SupportedInputFilter);
            if (selected == null || selected.Length == 0) return;
            if (selected.Length != 2)
            {
                Dialogs.Info(
                    "Exact Compare Preview currently compares exactly two files.\n\n" +
                    "Select the earlier file and the later file together. You can correct their Before/After " +
                    "order in the next window.");
                return;
            }

            ExactPreviewResult result;
            using (var form = new ComparisonSetupForm(selected[0], selected[1]))
            {
                if (form.ShowDialog(ExcelWindow.Current) != DialogResult.OK || form.Result == null) return;
                result = form.Result;
            }

            try
            {
                OpenValidatedReviewInExcel(result.OutputPath);
            }
            catch (Exception exception)
            {
                Dialogs.Info(
                    "The review workbook was created and verified, but Excel could not open it.\n\n" +
                    result.OutputPath + "\n\n" + DescribeException(exception));
            }
        }

        [ExcelCommand(
            Name = "PT_OpenComparisonReview",
            MenuName = "Power Tools",
            MenuText = "Open comparison review workbook...")]
        public static void OpenReviewWorkbook()
        {
            string path = Dialogs.OpenFile(ReviewWorkbookFilter);
            if (path == null) return;

            try
            {
                OpenValidatedReviewInExcel(path);
            }
            catch (Exception exception)
            {
                Dialogs.Info(
                    "This file could not be opened as a Power Tools comparison review. Nothing was changed.\n\n" +
                    DescribeException(exception));
            }
        }

        /// <summary>
        /// Inventories, reads, compares, publishes, and revalidates a new Review.xlsx. All parsing and package
        /// work resumes off the caller's synchronization context after the first asynchronous inventory read.
        /// </summary>
        public static async Task<ExactPreviewResult> GenerateExactPreviewAsync(
            string beforePath,
            string afterPath,
            string outputPath,
            IProgress<ComparisonPreviewProgress> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            string before = NormalizeExistingInput(beforePath, "Before");
            string after = NormalizeExistingInput(afterPath, "After");
            string output = NormalizeNewOutput(outputPath, before, after);

            if (string.Equals(before, after, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Before and After resolve to the same file. Choose two distinct files; byte-identical copies " +
                    "at different paths are supported.");

            var inventory = new ArtifactInventoryService();
            progress?.Report(new ComparisonPreviewProgress(
                "Inventorying Before", "Hashing and identifying " + Path.GetFileName(before), 0, true));
            ArtifactInventoryRecord beforeRecord = await inventory.InventoryAsync(
                before,
                new InventoryProgressForwarder(progress, "Before", 0, 18),
                cancellationToken).ConfigureAwait(false);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonPreviewProgress(
                "Inventorying After", "Hashing and identifying " + Path.GetFileName(after), 18, true));
            ArtifactInventoryRecord afterRecord = await inventory.InventoryAsync(
                after,
                new InventoryProgressForwarder(progress, "After", 18, 36),
                cancellationToken).ConfigureAwait(false);

            ExactPreviewFamily family = ValidatePair(beforeRecord, afterRecord);
            EnsurePreviewSize(beforeRecord);
            EnsurePreviewSize(afterRecord);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonPreviewProgress(
                "Reading immutable inputs", "Reading Before from disk and confirming its inventory hash", 38, true));
            byte[] beforeBytes = ReadAndVerifyBytes(beforeRecord, 38, 47, progress, cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonPreviewProgress(
                "Reading immutable inputs", "Reading After from disk and confirming its inventory hash", 47, true));
            byte[] afterBytes = ReadAndVerifyBytes(afterRecord, 47, 56, progress, cancellationToken);

            string projectId = StableId.Compute(
                "exact-preview-project-v1",
                beforeRecord.ContentSha256,
                Path.GetFileName(beforeRecord.FullNormalizedPath),
                afterRecord.ContentSha256,
                Path.GetFileName(afterRecord.FullNormalizedPath));
            string beforeRevisionId = StableId.Compute(
                "exact-preview-revision-v1", projectId, "before", beforeRecord.ContentSha256);
            string afterRevisionId = StableId.Compute(
                "exact-preview-revision-v1", projectId, "after", afterRecord.ContentSha256);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonPreviewProgress(
                "Extracting formula/value evidence", "Parsing both sources without starting Excel", 58, true));

            var projections = new List<PreviewChange>();
            var diagnostics = new List<PreviewDiagnostic>();
            bool byteIdentical;
            bool canAssertExactCellEquality;
            string status;
            int atomicChangeCount;

            if (family == ExactPreviewFamily.OpenXmlWorkbook)
            {
                OpenXmlWorkbookSnapshot beforeSnapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                    beforeRevisionId, beforeBytes, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                OpenXmlWorkbookSnapshot afterSnapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                    afterRevisionId, afterBytes, cancellationToken);

                CollectCoverage("Before", beforeSnapshot, diagnostics);
                CollectCoverage("After", afterSnapshot, diagnostics);

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new ComparisonPreviewProgress(
                    "Comparing exact cells", "Comparing same-name sheets, formulas, typed values, and cached results", 73, true));
                ExactWorkbookComparisonResult comparison = ExactWorkbookComparisonEngine.Compare(
                    beforeSnapshot, afterSnapshot, cancellationToken: cancellationToken);

                byteIdentical = comparison.ByteIdentical;
                canAssertExactCellEquality = comparison.CanAssertExactCellEquality;
                status = comparison.Status.ToString();
                atomicChangeCount = comparison.AtomicChangeCount;
                ProjectWorkbookChanges(projectId, comparison, projections, cancellationToken);
            }
            else
            {
                SeparatedValueSnapshot beforeSnapshot = SeparatedValueSnapshotAdapter.FromBytes(
                    beforeRevisionId, beforeBytes, beforeRecord.FullNormalizedPath);
                cancellationToken.ThrowIfCancellationRequested();
                SeparatedValueSnapshot afterSnapshot = SeparatedValueSnapshotAdapter.FromBytes(
                    afterRevisionId, afterBytes, afterRecord.FullNormalizedPath);

                progress?.Report(new ComparisonPreviewProgress(
                    "Comparing exact fields", "Comparing positional records and fields", 73, true));
                IReadOnlyList<AtomicCellChange> changes = ExactComparisonEngine.Compare(
                    beforeSnapshot.Snapshot,
                    afterSnapshot.Snapshot,
                    cancellationToken: cancellationToken);

                byteIdentical = string.Equals(
                    beforeSnapshot.ByteContentId,
                    afterSnapshot.ByteContentId,
                    StringComparison.Ordinal);
                canAssertExactCellEquality = true;
                status = changes.Count == 0 ? ExactComparisonStatus.Equal.ToString() :
                    ExactComparisonStatus.Different.ToString();
                atomicChangeCount = changes.Count;
                ProjectDelimitedChanges(projectId, changes, projections, cancellationToken);
            }

            if (projections.Count > ReviewWorkbookFile.MaxProjectedChangeRows)
                throw new InvalidDataException(
                    "The exact preview produced " + projections.Count.ToString("N0", CultureInfo.InvariantCulture) +
                    " visible findings, exceeding Excel's one-sheet review limit of " +
                    ReviewWorkbookFile.MaxProjectedChangeRows.ToString("N0", CultureInfo.InvariantCulture) +
                    ". Use a smaller pair or wait for the sharded comparison-project workflow.");

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ComparisonPreviewProgress(
                "Building review state", "Binding visible rows to the embedded exact evidence", 84, true));

            string stateJson = BuildStateJson(
                projectId,
                family,
                beforeRecord,
                afterRecord,
                status,
                byteIdentical,
                canAssertExactCellEquality,
                atomicChangeCount,
                projections,
                diagnostics);

            int stateBytes = new UTF8Encoding(false, true).GetByteCount(stateJson);
            if (stateBytes > ReviewWorkbookFile.MaxEmbeddedStateBytes)
                throw new InvalidDataException(
                    "The exact preview evidence requires " + stateBytes.ToString("N0", CultureInfo.InvariantCulture) +
                    " embedded bytes, exceeding the current Review.xlsx limit of " +
                    ReviewWorkbookFile.MaxEmbeddedStateBytes.ToString("N0", CultureInfo.InvariantCulture) +
                    ". Use a smaller pair or wait for archive-backed comparison projects.");

            var reviewState = new ReviewWorkbookState(projectId, stateJson);
            for (int i = 0; i < projections.Count; i++)
            {
                PreviewChange change = projections[i];
                reviewState.ChangeRows.Add(new ReviewChangeProjectionRow(
                    change.ChangeId,
                    change.Sheet,
                    change.Address,
                    change.Aspect,
                    change.Before,
                    change.After));
            }

            PreviewChange detail = FirstCharacterDetail(projections);
            if (detail != null)
            {
                reviewState.DetailBefore = detail.Before;
                reviewState.DetailAfter = detail.After;
            }

            // Publication is intentionally non-cancellable once it begins: ReviewWorkbookFile writes a staged
            // package, verifies it, flushes it, and atomically publishes a brand-new destination.
            progress?.Report(new ComparisonPreviewProgress(
                "Writing Review.xlsx", "Publishing and integrity-checking a new workbook", 90, false));
            ReviewWorkbookFile.CreateNew(output, reviewState);

            try
            {
                progress?.Report(new ComparisonPreviewProgress(
                    "Verifying Review.xlsx", "Reopening the closed package and checking state/projection hashes", 98, false));
                ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(output);
                if (!string.Equals(read.ProjectId, projectId, StringComparison.Ordinal) ||
                    !string.Equals(read.StateJson, stateJson, StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The closed review workbook did not retain the exact project state that was written.");
            }
            catch
            {
                TryDeleteGeneratedOutput(output);
                throw;
            }

            progress?.Report(new ComparisonPreviewProgress(
                "Complete", BuildCompletionText(status, atomicChangeCount, projections.Count, diagnostics.Count), 100, false));
            return new ExactPreviewResult(
                output,
                projectId,
                status,
                byteIdentical,
                canAssertExactCellEquality,
                atomicChangeCount,
                projections.Count,
                diagnostics.Count);
        }

        internal static string SuggestOutputPath(string beforePath, string afterPath)
        {
            string after = string.IsNullOrWhiteSpace(afterPath) ? null : TryFullPath(afterPath);
            string before = string.IsNullOrWhiteSpace(beforePath) ? null : TryFullPath(beforePath);
            string directory = after == null ? null : Path.GetDirectoryName(after);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                directory = before == null ? null : Path.GetDirectoryName(before);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                directory = Path.GetTempPath();

            string afterName = SafeBaseName(after, "after");
            string beforeName = SafeBaseName(before, "before");
            string stem = afterName + "-vs-" + beforeName + "-review";
            string candidate = Path.Combine(directory, stem + ".xlsx");
            for (int suffix = 2; File.Exists(candidate) || Directory.Exists(candidate); suffix++)
                candidate = Path.Combine(directory, stem + "-" + suffix.ToString(CultureInfo.InvariantCulture) + ".xlsx");
            return candidate;
        }

        internal static string DescribeException(Exception exception)
        {
            if (exception == null) return "Unknown error.";
            var messages = new List<string>();
            Exception current = exception;
            while (current != null && messages.Count < 6)
            {
                string message = (current.Message ?? "").Trim();
                if (message.Length > 0 && !messages.Contains(message)) messages.Add(message);
                current = current.InnerException;
            }
            return exception.GetType().FullName + (messages.Count == 0 ? "" : ": " + string.Join("\n", messages));
        }

        internal static void OpenValidatedReviewInExcel(string workbookPath)
        {
            OpenValidatedReviewInExcelCore(Path.GetFullPath(workbookPath), null);
        }

        /// <summary>
        /// Revalidates the exact occurrence retained by the comparison workspace, binds the private launch copy
        /// to its complete package identity, and returns the refreshed read-only session. No decision mutation is
        /// performed or enabled by this operation.
        /// </summary>
        internal static ReviewWorkbookHydrationSession OpenValidatedReviewInExcel(
            ReviewWorkbookHydrationSession expectedSession)
        {
            if (expectedSession == null) throw new ArgumentNullException(nameof(expectedSession));
            ReviewWorkbookHydrationSession current = expectedSession.RevalidateCurrentOccurrence();
            return OpenValidatedReviewInExcelCore(current.FullPath, current);
        }

        private static ReviewWorkbookHydrationSession OpenValidatedReviewInExcelCore(
            string workbookPath,
            ReviewWorkbookHydrationSession expectedSession)
        {
            string path = Path.GetFullPath(workbookPath);
            ReviewWorkbookLaunchCopy launch = null;
            dynamic opened = null;
            bool registered = false;
            try
            {
                // Copy and validate through retained handles before Excel sees any path. This closes the
                // validate-then-open race on the user-owned report and keeps the canonical artifact untouched.
                launch = ReviewWorkbookLaunchCopy.Create(path);
                if (expectedSession != null)
                {
                    EnsureLaunchMatchesHydratedSession(launch, expectedSession);
                    // Confirm that the path still names the same occurrence after the protected copy was made.
                    // Excel receives only the already-protected private copy after this boundary.
                    expectedSession = expectedSession.RevalidateCurrentOccurrence();
                    EnsureLaunchMatchesHydratedSession(launch, expectedSession);
                }
                dynamic app = ExcelApp.App;

                // msoAutomationSecurityForceDisable = 3. Fail closed if Excel cannot accept or report the
                // setting; opening active content is never an acceptable fallback. UpdateLinks:=0 and
                // ReadOnly:=true are explicit at the Workbooks.Open boundary.
                object previousAutomationSecurity = app.AutomationSecurity;
                try
                {
                    app.AutomationSecurity = 3;
                    int applied = Convert.ToInt32(app.AutomationSecurity, CultureInfo.InvariantCulture);
                    if (applied != 3)
                        throw new InvalidOperationException(
                            "Excel did not confirm that automation security was forced off for the review launch.");
                    opened = app.Workbooks.Open(launch.LaunchPath, 0, true);
                }
                finally
                {
                    app.AutomationSecurity = previousAutomationSecurity;
                }

                string openedPath = TryGetOpenWorkbookPath(opened);
                if (!string.Equals(openedPath, launch.LaunchPath, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "Excel returned a workbook whose path does not match the protected review launch copy.");
                if (!Convert.ToBoolean(opened.ReadOnly, CultureInfo.InvariantCulture))
                    throw new InvalidOperationException(
                        "Excel did not open the protected review launch copy read-only.");

                // Revalidate both the passive schema identity and whole-package digest after Workbooks.Open
                // returns, while the no-write/no-delete protection handle is still retained.
                launch.VerifyUnchanged();
                ComparisonReviewLaunchManager.Register(launch);
                launch = null; // lifetime ownership transfers to the manager until close or Save As
                registered = true;

                opened.Activate();
                Dialogs.Info(
                    "Opened a protected, validated read-only copy.\n\n" +
                    "Original Review.xlsx (share or reopen this file):\n" + path + "\n\n" +
                    "Excel cannot save this window back to the original. File > Save As creates a separate " +
                    "working copy; it does not replace the comparison project artifact above.");
                return expectedSession;
            }
            catch
            {
                // Only close the new private launch workbook, and only when Excel confirms its exact path.
                // Never close a pre-existing or path-mismatched user workbook.
                if (!registered) TryCloseOwnedLaunchWorkbook(opened, launch == null ? null : launch.LaunchPath);
                throw;
            }
            finally
            {
                if (launch != null) launch.Dispose();
            }
        }

        private static void EnsureLaunchMatchesHydratedSession(
            ReviewWorkbookLaunchCopy launch,
            ReviewWorkbookHydrationSession expectedSession)
        {
            if (launch == null) throw new ArgumentNullException(nameof(launch));
            if (expectedSession == null) throw new ArgumentNullException(nameof(expectedSession));
            if (launch.ByteLength != expectedSession.FileByteLength ||
                !string.Equals(
                    launch.PackageSha256,
                    expectedSession.PackageSha256,
                    StringComparison.OrdinalIgnoreCase) ||
                !ReviewWorkbookLaunchCopy.HasSameValidatedIdentity(
                    expectedSession.ReadResult,
                    launch.Validation))
                throw new InvalidDataException(
                    "The Review workbook occurrence changed before its protected launch copy was finalized.");
        }

        private static string TryGetOpenWorkbookPath(dynamic workbook)
        {
            if (workbook == null) return null;
            try
            {
                string fullName = (string)workbook.FullName;
                return string.IsNullOrWhiteSpace(fullName) ? null : Path.GetFullPath(fullName);
            }
            catch { return null; }
        }

        private static void TryCloseOwnedLaunchWorkbook(dynamic workbook, string launchPath)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(launchPath)) return;
            string openedPath = TryGetOpenWorkbookPath(workbook);
            if (!string.Equals(openedPath, launchPath, StringComparison.OrdinalIgnoreCase)) return;
            try { workbook.Close(false); }
            catch { }
        }

        private static string NormalizeExistingInput(string path, string role)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new InvalidOperationException(role + " file is required.");
            string full;
            try { full = Path.GetFullPath(path); }
            catch (Exception exception)
            {
                throw new InvalidOperationException(role + " path is invalid: " + exception.Message, exception);
            }
            if (!File.Exists(full))
                throw new FileNotFoundException(role + " file was not found. Choose an existing local file.", full);
            return full;
        }

        private static string NormalizeNewOutput(string path, string before, string after)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new InvalidOperationException("A destination for the new Review.xlsx is required.");
            string full;
            try { full = Path.GetFullPath(path); }
            catch (Exception exception)
            {
                throw new InvalidOperationException("The review destination is invalid: " + exception.Message, exception);
            }
            if (!string.Equals(Path.GetExtension(full), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The review destination must end in .xlsx.");
            string directory = Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException("The review destination folder does not exist.");
            if (File.Exists(full) || Directory.Exists(full))
                throw new IOException(
                    "The review destination already exists. Choose a new filename; comparison never overwrites an artifact.");
            if (string.Equals(full, before, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(full, after, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The review destination cannot replace either input file.");
            return full;
        }

        private static ExactPreviewFamily ValidatePair(
            ArtifactInventoryRecord before,
            ArtifactInventoryRecord after)
        {
            ValidateInventory("Before", before);
            ValidateInventory("After", after);

            ArtifactFormatKind beforeFormat = before.Format.EffectiveFormat;
            ArtifactFormatKind afterFormat = after.Format.EffectiveFormat;
            if (beforeFormat == ArtifactFormatKind.Xlsb || afterFormat == ArtifactFormatKind.Xlsb)
                throw new NotSupportedException(
                    "XLSB was detected, but safe XLSB formula extraction is not part of Exact Compare Preview yet. " +
                    "Save a comparison copy as .xlsx/.xlsm, or wait for the isolated XLSB adapter.");

            bool beforeWorkbook = IsOpenXmlWorkbook(beforeFormat);
            bool afterWorkbook = IsOpenXmlWorkbook(afterFormat);
            bool beforeDelimited = IsDelimited(beforeFormat);
            bool afterDelimited = IsDelimited(afterFormat);

            if (beforeWorkbook && afterWorkbook) return ExactPreviewFamily.OpenXmlWorkbook;
            if (beforeDelimited && afterDelimited) return ExactPreviewFamily.Delimited;
            if ((beforeWorkbook && afterDelimited) || (beforeDelimited && afterWorkbook))
                throw new NotSupportedException(
                    "The selected files belong to different comparison families. Compare two .xlsx/.xlsm " +
                    "workbooks or two .csv/.tsv/.txt separated-value files.");

            throw new NotSupportedException(
                "This pair is not supported by Exact Compare Preview. Compare two .xlsx/.xlsm workbooks or " +
                "two .csv/.tsv/.txt separated-value files.");
        }

        private static void ValidateInventory(string role, ArtifactInventoryRecord record)
        {
            if (record.Format.EffectiveFormat == ArtifactFormatKind.Xlsb) return;
            if (record.Format.IsSupportedForComparison) return;
            throw new InvalidDataException(
                role + " could not be accepted as a comparison source. " + record.Format.Message);
        }

        private static bool IsOpenXmlWorkbook(ArtifactFormatKind format) =>
            format == ArtifactFormatKind.Xlsx || format == ArtifactFormatKind.Xlsm;

        private static bool IsDelimited(ArtifactFormatKind format) =>
            format == ArtifactFormatKind.Csv ||
            format == ArtifactFormatKind.Tsv ||
            format == ArtifactFormatKind.DelimitedText;

        private static void EnsurePreviewSize(ArtifactInventoryRecord record)
        {
            if (record.ByteLength <= MaxPreviewInputBytes) return;
            throw new InvalidDataException(
                "'" + Path.GetFileName(record.FullNormalizedPath) + "' is " +
                FormatBytes(record.ByteLength) + ", exceeding the current in-memory preview limit of " +
                FormatBytes(MaxPreviewInputBytes) + ". The file was inventoried but was not parsed.");
        }

        private static byte[] ReadAndVerifyBytes(
            ArtifactInventoryRecord record,
            int progressStart,
            int progressEnd,
            IProgress<ComparisonPreviewProgress> progress,
            CancellationToken cancellationToken)
        {
            if (record.ByteLength > int.MaxValue)
                throw new InvalidDataException("The source is too large for the bounded in-memory preview.");
            var bytes = new byte[(int)record.ByteLength];
            int offset = 0;
            using (var stream = new FileStream(
                record.FullNormalizedPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                128 * 1024,
                FileOptions.SequentialScan))
            {
                while (offset < bytes.Length)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read = stream.Read(bytes, offset, Math.Min(128 * 1024, bytes.Length - offset));
                    if (read == 0)
                        throw new IOException(
                            "The source became shorter after inventory. Wait for all writers or sync clients, then retry.");
                    offset += read;
                    if (progress != null && (offset == bytes.Length || offset % (1024 * 1024) < read))
                    {
                        int percent = bytes.Length == 0
                            ? progressEnd
                            : progressStart + (int)((long)(progressEnd - progressStart) * offset / bytes.Length);
                        progress.Report(new ComparisonPreviewProgress(
                            "Reading immutable inputs",
                            Path.GetFileName(record.FullNormalizedPath) + "  " +
                            FormatBytes(offset) + " / " + FormatBytes(bytes.Length),
                            percent,
                            true));
                    }
                }

                if (stream.ReadByte() != -1)
                    throw new IOException(
                        "The source became longer after inventory. Wait for all writers or sync clients, then retry.");
            }

            string actualHash = StableId.ComputeByteContentId(bytes);
            if (!string.Equals(actualHash, record.ContentSha256, StringComparison.OrdinalIgnoreCase))
                throw new IOException(
                    "The source bytes changed between inventory and parsing. Nothing was compared. Wait for all " +
                    "writers or sync clients, then retry.");
            return bytes;
        }

        private static void CollectCoverage(
            string role,
            OpenXmlWorkbookSnapshot workbook,
            IList<PreviewDiagnostic> diagnostics)
        {
            for (int i = 0; i < workbook.CoverageMessages.Count; i++)
                AddCoverage(role, workbook.CoverageMessages[i], diagnostics);
            for (int sheetIndex = 0; sheetIndex < workbook.Sheets.Count; sheetIndex++)
            {
                OpenXmlWorksheetSnapshot sheet = workbook.Sheets[sheetIndex];
                for (int i = 0; i < sheet.CoverageMessages.Count; i++)
                    AddCoverage(role, sheet.CoverageMessages[i], diagnostics);
            }
        }

        private static void AddCoverage(
            string role,
            OpenXmlCoverageMessage message,
            IList<PreviewDiagnostic> diagnostics)
        {
            diagnostics.Add(new PreviewDiagnostic(
                role,
                message.Code,
                message.Message,
                message.AffectsExactCellComparison,
                message.SheetName,
                message.CellReference,
                message.PartUri));
        }

        private static void ProjectWorkbookChanges(
            string projectId,
            ExactWorkbookComparisonResult result,
            IList<PreviewChange> projections,
            CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < result.Sheets.Count; sheetIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ExactSheetComparison sheet = result.Sheets[sheetIndex];
                if (sheet.MatchKind != ExactSheetMatchKind.ExactName)
                {
                    string aspect;
                    string before;
                    string after;
                    switch (sheet.MatchKind)
                    {
                        case ExactSheetMatchKind.OnlyBefore:
                            aspect = "Sheet removed or renamed";
                            before = "present";
                            after = "missing";
                            break;
                        case ExactSheetMatchKind.OnlyAfter:
                            aspect = "Sheet added or renamed";
                            before = "missing";
                            after = "present";
                            break;
                        default:
                            aspect = "Ambiguous sheet name";
                            before = "ambiguous";
                            after = "ambiguous";
                            break;
                    }
                    AddProjection(projectId, sheet.SheetName, "", aspect, before, after, projections);
                }

                if (sheet.VisibilityChanged)
                {
                    AddProjection(
                        projectId,
                        sheet.SheetName,
                        "",
                        "Sheet visibility",
                        sheet.Before.Visibility.ToString(),
                        sheet.After.Visibility.ToString(),
                        projections);
                }

                for (int changeIndex = 0; changeIndex < sheet.Changes.Count; changeIndex++)
                {
                    AtomicCellChange change = sheet.Changes[changeIndex];
                    AddProjection(
                        projectId,
                        sheet.SheetName,
                        change.SubjectKey,
                        DisplayAspect(change.Aspect) + " - " + change.Kind,
                        change.BeforeText,
                        change.AfterText,
                        projections);
                }
            }
        }

        private static void ProjectDelimitedChanges(
            string projectId,
            IReadOnlyList<AtomicCellChange> changes,
            IList<PreviewChange> projections,
            CancellationToken cancellationToken)
        {
            for (int i = 0; i < changes.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                AtomicCellChange change = changes[i];
                AddProjection(
                    projectId,
                    "Delimited data",
                    DelimitedAddress(change.SubjectKey),
                    DisplayAspect(change.Aspect) + " - " + change.Kind,
                    change.BeforeText,
                    change.AfterText,
                    projections);
            }
        }

        private static void AddProjection(
            string projectId,
            string sheet,
            string address,
            string aspect,
            string before,
            string after,
            IList<PreviewChange> projections)
        {
            string changeId = StableId.Compute(
                "exact-preview-change-v1",
                projectId,
                projections.Count.ToString(CultureInfo.InvariantCulture),
                sheet ?? "",
                address ?? "",
                aspect ?? "",
                before,
                after);
            projections.Add(new PreviewChange(
                changeId,
                sheet ?? "",
                address ?? "",
                aspect ?? "",
                before ?? "",
                after ?? ""));
        }

        private static string DisplayAspect(CellChangeAspect aspect)
        {
            switch (aspect)
            {
                case CellChangeAspect.Formula: return "Formula";
                case CellChangeAspect.FormulaStructure: return "Formula structure";
                case CellChangeAspect.Value: return "Value";
                case CellChangeAspect.CachedResult: return "Cached result";
                default: return "Cell presence";
            }
        }

        private static string DelimitedAddress(string subjectKey)
        {
            const string rowMarker = "/r:";
            const string columnMarker = "/c:";
            if (string.IsNullOrEmpty(subjectKey)) return "";
            int rowAt = subjectKey.IndexOf(rowMarker, StringComparison.Ordinal);
            int columnAt = subjectKey.IndexOf(columnMarker, StringComparison.Ordinal);
            if (rowAt < 0 || columnAt <= rowAt) return subjectKey;

            int row;
            int column;
            string rowText = subjectKey.Substring(
                rowAt + rowMarker.Length,
                columnAt - rowAt - rowMarker.Length);
            string columnText = subjectKey.Substring(columnAt + columnMarker.Length);
            if (!int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out row) ||
                !int.TryParse(columnText, NumberStyles.None, CultureInfo.InvariantCulture, out column) ||
                row <= 0 || column <= 0)
                return subjectKey;
            return ExcelColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
        }

        private static string ExcelColumnName(int oneBasedColumn)
        {
            var result = new StringBuilder();
            int value = oneBasedColumn;
            while (value > 0)
            {
                value--;
                result.Insert(0, (char)('A' + value % 26));
                value /= 26;
            }
            return result.ToString();
        }

        private static PreviewChange FirstCharacterDetail(IList<PreviewChange> projections)
        {
            for (int i = 0; i < projections.Count; i++)
            {
                PreviewChange change = projections[i];
                if (change.Address.Length > 0 && (change.Before.Length > 0 || change.After.Length > 0))
                    return change;
            }
            return null;
        }

        private static string BuildStateJson(
            string projectId,
            ExactPreviewFamily family,
            ArtifactInventoryRecord before,
            ArtifactInventoryRecord after,
            string status,
            bool byteIdentical,
            bool canAssertExactCellEquality,
            int atomicChangeCount,
            IList<PreviewChange> changes,
            IList<PreviewDiagnostic> diagnostics)
        {
            int estimatedRows = Math.Min(changes.Count, 5000);
            var json = new StringBuilder(Math.Min(1024 * 1024, 4096 + estimatedRows * 180));
            json.Append('{');
            JsonProperty(json, "schema", "exact-preview-v1", true);
            JsonProperty(json, "mode", "bounded-exact-pairwise-preview", false);
            JsonProperty(json, "projectId", projectId, false);
            JsonProperty(json, "generatedUtc", DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture), false);
            JsonProperty(json, "family", family == ExactPreviewFamily.OpenXmlWorkbook ? "openxml-workbook" : "delimited", false);

            json.Append(",\"limitations\":[");
            string[] limitations =
            {
                "Exactly two revisions",
                "Positional comparison only",
                "Only exact same-name workbook sheets are paired",
                "No row or column alignment",
                "No grouping, snapshots, review decisions, merge, or XLSB extraction",
            };
            for (int i = 0; i < limitations.Length; i++)
            {
                if (i > 0) json.Append(',');
                AppendJsonString(json, limitations[i]);
            }
            json.Append(']');

            json.Append(",\"before\":");
            AppendArtifactJson(json, before);
            json.Append(",\"after\":");
            AppendArtifactJson(json, after);

            json.Append(",\"comparison\":{");
            JsonProperty(json, "status", status, true);
            JsonBooleanProperty(json, "byteIdentical", byteIdentical, false);
            JsonBooleanProperty(json, "canAssertExactCellEquality", canAssertExactCellEquality, false);
            JsonNumberProperty(json, "atomicChangeCount", atomicChangeCount, false);
            JsonNumberProperty(json, "projectedFindingCount", changes.Count, false);
            json.Append('}');

            json.Append(",\"changes\":[");
            for (int i = 0; i < changes.Count; i++)
            {
                if (i > 0) json.Append(',');
                PreviewChange change = changes[i];
                json.Append('{');
                JsonProperty(json, "changeId", change.ChangeId, true);
                JsonProperty(json, "sheet", change.Sheet, false);
                JsonProperty(json, "address", change.Address, false);
                JsonProperty(json, "aspect", change.Aspect, false);
                JsonProperty(json, "before", change.Before, false);
                JsonProperty(json, "after", change.After, false);
                json.Append('}');
            }
            json.Append(']');

            json.Append(",\"diagnostics\":[");
            for (int i = 0; i < diagnostics.Count; i++)
            {
                if (i > 0) json.Append(',');
                PreviewDiagnostic diagnostic = diagnostics[i];
                json.Append('{');
                JsonProperty(json, "revisionRole", diagnostic.RevisionRole, true);
                JsonProperty(json, "code", diagnostic.Code, false);
                JsonProperty(json, "message", diagnostic.Message, false);
                JsonBooleanProperty(json, "blocksExactCellEquality", diagnostic.BlocksExactCellEquality, false);
                JsonProperty(json, "sheet", diagnostic.Sheet, false);
                JsonProperty(json, "cell", diagnostic.Cell, false);
                JsonProperty(json, "part", diagnostic.Part, false);
                json.Append('}');
            }
            json.Append("]}");
            return json.ToString();
        }

        private static void AppendArtifactJson(StringBuilder json, ArtifactInventoryRecord artifact)
        {
            json.Append('{');
            // Deliberately retain only the basename. The exact content hash and weak timestamps identify the
            // imported occurrence without leaking the user's directory or account name into a shared report.
            JsonProperty(json, "fileName", Path.GetFileName(artifact.FullNormalizedPath), true);
            JsonProperty(json, "contentSha256", artifact.ContentSha256, false);
            JsonNumberProperty(json, "byteLength", artifact.ByteLength, false);
            JsonProperty(json, "creationTimeUtc", artifact.WeakCreationTimeUtc.ToString("O", CultureInfo.InvariantCulture), false);
            JsonProperty(json, "lastWriteTimeUtc", artifact.WeakLastWriteTimeUtc.ToString("O", CultureInfo.InvariantCulture), false);
            JsonProperty(json, "format", artifact.Format.EffectiveFormat.ToString(), false);
            JsonProperty(json, "formatMatch", artifact.Format.MatchKind.ToString(), false);
            json.Append('}');
        }

        private static void JsonProperty(StringBuilder json, string name, string value, bool first)
        {
            if (!first) json.Append(',');
            AppendJsonString(json, name);
            json.Append(':');
            AppendJsonString(json, value ?? "");
        }

        private static void JsonBooleanProperty(StringBuilder json, string name, bool value, bool first)
        {
            if (!first) json.Append(',');
            AppendJsonString(json, name);
            json.Append(value ? ":true" : ":false");
        }

        private static void JsonNumberProperty(StringBuilder json, string name, long value, bool first)
        {
            if (!first) json.Append(',');
            AppendJsonString(json, name);
            json.Append(':').Append(value.ToString(CultureInfo.InvariantCulture));
        }

        private static void AppendJsonString(StringBuilder json, string value)
        {
            json.Append('"');
            value = value ?? "";
            for (int i = 0; i < value.Length; i++)
            {
                char character = value[i];
                switch (character)
                {
                    case '"': json.Append("\\\""); break;
                    case '\\': json.Append("\\\\"); break;
                    case '\b': json.Append("\\b"); break;
                    case '\f': json.Append("\\f"); break;
                    case '\n': json.Append("\\n"); break;
                    case '\r': json.Append("\\r"); break;
                    case '\t': json.Append("\\t"); break;
                    default:
                        if (character < 0x20 || character == '\u2028' || character == '\u2029' ||
                            char.IsSurrogate(character))
                        {
                            json.Append("\\u").Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            json.Append(character);
                        }
                        break;
                }
            }
            json.Append('"');
        }

        private static string BuildCompletionText(
            string status,
            int atomicChangeCount,
            int projectedCount,
            int diagnosticCount)
        {
            return status + ": " + atomicChangeCount.ToString("N0", CultureInfo.InvariantCulture) +
                   " exact cell aspect change(s), " + projectedCount.ToString("N0", CultureInfo.InvariantCulture) +
                   " visible finding(s), " + diagnosticCount.ToString("N0", CultureInfo.InvariantCulture) +
                   " coverage diagnostic(s).";
        }

        private static void TryDeleteGeneratedOutput(string output)
        {
            try { if (File.Exists(output)) File.Delete(output); }
            catch { }
        }

        private static string TryFullPath(string path)
        {
            try { return Path.GetFullPath(path); }
            catch { return path; }
        }

        private static string SafeBaseName(string path, string fallback)
        {
            if (string.IsNullOrWhiteSpace(path)) return fallback;
            string value;
            try { value = Path.GetFileNameWithoutExtension(path); }
            catch { value = fallback; }
            if (string.IsNullOrWhiteSpace(value)) return fallback;
            foreach (char invalid in Path.GetInvalidFileNameChars()) value = value.Replace(invalid, '_');
            const int maxComponentLength = 70;
            return value.Length > maxComponentLength ? value.Substring(0, maxComponentLength) : value;
        }

        private static string FormatBytes(long value)
        {
            string[] units = { "B", "KiB", "MiB", "GiB" };
            double display = value;
            int unit = 0;
            while (display >= 1024 && unit < units.Length - 1)
            {
                display /= 1024;
                unit++;
            }
            return display.ToString(unit == 0 ? "N0" : "N1", CultureInfo.InvariantCulture) + " " + units[unit];
        }

        private enum ExactPreviewFamily
        {
            OpenXmlWorkbook,
            Delimited,
        }

        private sealed class InventoryProgressForwarder : IProgress<ArtifactInventoryProgress>
        {
            private readonly IProgress<ComparisonPreviewProgress> _outer;
            private readonly string _role;
            private readonly int _start;
            private readonly int _end;

            internal InventoryProgressForwarder(
                IProgress<ComparisonPreviewProgress> outer,
                string role,
                int start,
                int end)
            {
                _outer = outer;
                _role = role;
                _start = start;
                _end = end;
            }

            public void Report(ArtifactInventoryProgress value)
            {
                if (_outer == null || value == null) return;
                int percent = _start + (int)Math.Round((_end - _start) * value.Fraction);
                _outer.Report(new ComparisonPreviewProgress(
                    "Inventorying " + _role,
                    value.Stage + " - " + Path.GetFileName(value.FullNormalizedPath) +
                    (value.TotalBytes <= 0 ? "" : "  " + FormatBytes(value.BytesRead) + " / " + FormatBytes(value.TotalBytes)),
                    percent,
                    true));
            }
        }

        private sealed class PreviewChange
        {
            internal PreviewChange(
                string changeId,
                string sheet,
                string address,
                string aspect,
                string before,
                string after)
            {
                ChangeId = changeId;
                Sheet = sheet;
                Address = address;
                Aspect = aspect;
                Before = before;
                After = after;
            }

            internal string ChangeId { get; }
            internal string Sheet { get; }
            internal string Address { get; }
            internal string Aspect { get; }
            internal string Before { get; }
            internal string After { get; }
        }

        private sealed class PreviewDiagnostic
        {
            internal PreviewDiagnostic(
                string revisionRole,
                string code,
                string message,
                bool blocksExactCellEquality,
                string sheet,
                string cell,
                string part)
            {
                RevisionRole = revisionRole ?? "";
                Code = code ?? "";
                Message = message ?? "";
                BlocksExactCellEquality = blocksExactCellEquality;
                Sheet = sheet ?? "";
                Cell = cell ?? "";
                Part = part ?? "";
            }

            internal string RevisionRole { get; }
            internal string Code { get; }
            internal string Message { get; }
            internal bool BlocksExactCellEquality { get; }
            internal string Sheet { get; }
            internal string Cell { get; }
            internal string Part { get; }
        }
    }

    public sealed class ComparisonPreviewProgress
    {
        public ComparisonPreviewProgress(string phase, string detail, int percent, bool canCancel)
        {
            Phase = phase ?? "";
            Detail = detail ?? "";
            Percent = Math.Max(0, Math.Min(100, percent));
            CanCancel = canCancel;
        }

        public string Phase { get; }
        public string Detail { get; }
        public int Percent { get; }
        public bool CanCancel { get; }
    }

    public sealed class ExactPreviewResult
    {
        internal ExactPreviewResult(
            string outputPath,
            string projectId,
            string status,
            bool byteIdentical,
            bool canAssertExactCellEquality,
            int atomicChangeCount,
            int projectedFindingCount,
            int diagnosticCount)
        {
            OutputPath = outputPath;
            ProjectId = projectId;
            Status = status;
            ByteIdentical = byteIdentical;
            CanAssertExactCellEquality = canAssertExactCellEquality;
            AtomicChangeCount = atomicChangeCount;
            ProjectedFindingCount = projectedFindingCount;
            DiagnosticCount = diagnosticCount;
        }

        public string OutputPath { get; }
        public string ProjectId { get; }
        public string Status { get; }
        public bool ByteIdentical { get; }
        public bool CanAssertExactCellEquality { get; }
        public int AtomicChangeCount { get; }
        public int ProjectedFindingCount { get; }
        public int DiagnosticCount { get; }
    }
}
