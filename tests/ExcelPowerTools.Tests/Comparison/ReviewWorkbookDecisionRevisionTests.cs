using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewWorkbookDecisionRevisionTests : IDisposable
    {
        private const string SubjectId = "subject-value";
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.DecisionRevision.Tests",
            Guid.NewGuid().ToString("N"));

        public ReviewWorkbookDecisionRevisionTests()
        {
            Directory.CreateDirectory(_directory);
        }

        [Fact]
        public void Creates_copy_on_write_revision_and_changes_only_decision_and_manifest_parts()
        {
            string source = CreateDecisionWorkbook("source.xlsx");
            string destination = Path.Combine(_directory, "revision-1.xlsx");
            byte[] sourceBytes = File.ReadAllBytes(source);
            ReviewWorkbookReadResult before = ReviewWorkbookFile.ReadAndValidate(source);
            PackageSnapshot sourcePackage = PackageSnapshot.Read(source);
            ReviewDecisionRequest request = Request(
                before.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 0, 0, TimeSpan.Zero),
                "run-first");

            ReviewWorkbookFile.CreateDecisionRevision(source, destination, request);

            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            ReviewWorkbookReadResult after = ReviewWorkbookFile.ReadAndValidate(destination);
            AssertRevisionIdentityPreserved(before, after);
            Assert.Single(after.DecisionGeneration!.Events);
            Assert.Equal(request.RequestId, after.DecisionGeneration.Events[0].RequestId);
            Assert.Equal(
                ReviewDecisionStatus.NeedsInvestigation,
                after.DecisionGeneration.GetSubjectState(SubjectId).Status);
            Assert.Equal(before.DecisionGeneration.GenerationId, after.DecisionGeneration.PriorGenerationId);

            PackageSnapshot revisedPackage = PackageSnapshot.Read(destination);
            AssertOnlyDecisionAndManifestPartsChanged(sourcePackage, revisedPackage);
            AssertManifestChangedOnlyAtDigest(sourcePackage, revisedPackage, before.DecisionSha256, after.DecisionSha256);
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Manifest_digest_scanner_ignores_attribute_name_literal_inside_project_id()
        {
            const string projectId = "project-with-decisionSha256-literal";
            string source = CreateDecisionWorkbook("literal-source.xlsx", projectId);
            string destination = Path.Combine(_directory, "literal-revision.xlsx");
            ReviewWorkbookReadResult before = ReviewWorkbookFile.ReadAndValidate(source);
            PackageSnapshot sourcePackage = PackageSnapshot.Read(source);
            Assert.Equal(projectId, before.ProjectId);
            Assert.True(
                CountOccurrences(
                    sourcePackage.Entries[sourcePackage.ManifestEntry],
                    Encoding.ASCII.GetBytes("decisionSha256")) >= 2);

            ReviewWorkbookFile.CreateDecisionRevision(
                source,
                destination,
                Request(
                    before.DecisionGeneration!,
                    ReviewDecisionStatus.NeedsInvestigation,
                    new DateTimeOffset(2026, 7, 17, 8, 2, 0, TimeSpan.Zero),
                    "literal-project"));

            ReviewWorkbookReadResult after = ReviewWorkbookFile.ReadAndValidate(destination);
            PackageSnapshot revisedPackage = PackageSnapshot.Read(destination);
            Assert.Equal(projectId, after.ProjectId);
            AssertRevisionIdentityPreserved(before, after);
            AssertOnlyDecisionAndManifestPartsChanged(sourcePackage, revisedPackage);
            AssertManifestChangedOnlyAtDigest(
                sourcePackage,
                revisedPackage,
                before.DecisionSha256,
                after.DecisionSha256);
        }

        [Fact]
        public void Rejects_legacy_source_without_sidecar_and_cleans_stage()
        {
            string source = Path.Combine(_directory, "legacy.xlsx");
            ReviewWorkbookFile.CreateNew(source, new ReviewWorkbookState("legacy-project", "{\"legacy\":true}"));
            byte[] sourceBytes = File.ReadAllBytes(source);
            string destination = Path.Combine(_directory, "legacy-revision.xlsx");
            ReviewDecisionGeneration unrelated = CreateDecisionRoot();

            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                ReviewWorkbookFile.CreateDecisionRevision(
                    source,
                    destination,
                    Request(
                        unrelated,
                        ReviewDecisionStatus.NeedsInvestigation,
                        new DateTimeOffset(2026, 7, 17, 8, 5, 0, TimeSpan.Zero),
                        "legacy")));

            Assert.Contains("sidecar", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            Assert.False(File.Exists(destination));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Rejects_stale_request_without_copying_or_mutating_source()
        {
            string source = CreateDecisionWorkbook("stale-source.xlsx");
            string destination = Path.Combine(_directory, "stale-destination.xlsx");
            byte[] sourceBytes = File.ReadAllBytes(source);
            ReviewWorkbookReadResult current = ReviewWorkbookFile.ReadAndValidate(source);
            ReviewDecisionRequest stale = ReviewDecisionRequest.ForSubject(
                current.ProjectId,
                current.DecisionGeneration!.BaseGenerationId,
                current.DecisionGeneration.MappingGenerationId,
                new string('f', 64),
                SubjectId,
                ReviewDecisionStatus.NeedsInvestigation,
                ReviewerSnapshot.Empty,
                new DateTimeOffset(2026, 7, 17, 8, 10, 0, TimeSpan.Zero));

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                ReviewWorkbookFile.CreateDecisionRevision(source, destination, stale));

            Assert.Contains("stale decision generation", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            Assert.False(File.Exists(destination));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Existing_destination_is_never_replaced()
        {
            string source = CreateDecisionWorkbook("existing-source.xlsx");
            string destination = Path.Combine(_directory, "existing-destination.xlsx");
            byte[] sentinel = { 4, 8, 15, 16, 23, 42 };
            File.WriteAllBytes(destination, sentinel);
            byte[] sourceBytes = File.ReadAllBytes(source);
            ReviewWorkbookReadResult current = ReviewWorkbookFile.ReadAndValidate(source);

            Assert.Throws<IOException>(() => ReviewWorkbookFile.CreateDecisionRevision(
                source,
                destination,
                Request(
                    current.DecisionGeneration!,
                    ReviewDecisionStatus.NeedsInvestigation,
                    new DateTimeOffset(2026, 7, 17, 8, 15, 0, TimeSpan.Zero),
                    "existing")));

            Assert.Equal(sentinel, File.ReadAllBytes(destination));
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Tampered_source_is_rejected_before_destination_creation()
        {
            string source = CreateDecisionWorkbook("tampered-source.xlsx");
            MutateManifestDigest(source, new string('0', 64));
            byte[] tamperedBytes = File.ReadAllBytes(source);
            string destination = Path.Combine(_directory, "tampered-destination.xlsx");
            ReviewDecisionGeneration root = CreateDecisionRoot();

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateDecisionRevision(
                source,
                destination,
                Request(
                    root,
                    ReviewDecisionStatus.NeedsInvestigation,
                    new DateTimeOffset(2026, 7, 17, 8, 20, 0, TimeSpan.Zero),
                    "tampered")));

            Assert.Equal(tamperedBytes, File.ReadAllBytes(source));
            Assert.False(File.Exists(destination));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Trailing_source_growth_is_rejected_and_stage_is_cleaned()
        {
            string source = CreateDecisionWorkbook("trailing-source.xlsx");
            using (var append = new FileStream(source, FileMode.Append, FileAccess.Write, FileShare.Read))
                append.Write(new byte[] { 0x54, 0x52, 0x41, 0x49, 0x4C }, 0, 5);
            byte[] sourceBytes = File.ReadAllBytes(source);
            string destination = Path.Combine(_directory, "trailing-destination.xlsx");
            ReviewDecisionGeneration root = CreateDecisionRoot();

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateDecisionRevision(
                source,
                destination,
                Request(
                    root,
                    ReviewDecisionStatus.NeedsInvestigation,
                    new DateTimeOffset(2026, 7, 17, 8, 22, 0, TimeSpan.Zero),
                    "trailing")));

            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            Assert.False(File.Exists(destination));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Bounded_copy_rejects_a_length_different_from_validated_snapshot()
        {
            MethodInfo copy = typeof(ReviewWorkbookFile).GetMethod(
                "CopyValidatedSourceBytes",
                BindingFlags.Static | BindingFlags.NonPublic)!;
            using var source = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 }, writable: false);
            using var stage = new MemoryStream();

            TargetInvocationException error = Assert.Throws<TargetInvocationException>(() =>
                copy.Invoke(null, new object[] { source, stage, 4L }));

            Assert.IsType<InvalidDataException>(error.InnerException);
            Assert.Equal(5, stage.Length);
        }

        [Fact]
        public async Task Racing_publishers_never_overwrite_destination_and_clean_losing_stage()
        {
            string source = CreateDecisionWorkbook("race-source.xlsx");
            string destination = Path.Combine(_directory, "race-destination.xlsx");
            byte[] sourceBytes = File.ReadAllBytes(source);
            ReviewWorkbookReadResult current = ReviewWorkbookFile.ReadAndValidate(source);
            ReviewDecisionRequest request = Request(
                current.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 25, 0, TimeSpan.Zero),
                "race");
            using var start = new ManualResetEventSlim(false);

            Task<Exception?> first = Task.Run(() => PublishAndCapture(start, source, destination, request));
            Task<Exception?> second = Task.Run(() => PublishAndCapture(start, source, destination, request));
            start.Set();
            Exception?[] results = await Task.WhenAll(first, second);

            Assert.Single(results, error => error == null);
            IOException losingError = Assert.IsAssignableFrom<IOException>(
                Assert.Single(results, error => error != null));
            Assert.NotNull(losingError);
            ReviewWorkbookReadResult revised = ReviewWorkbookFile.ReadAndValidate(destination);
            Assert.Single(revised.DecisionGeneration!.Events);
            Assert.Equal(request.RequestId, revised.DecisionGeneration.Events[0].RequestId);
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Two_successive_revisions_preserve_identifiers_and_append_exact_request_history()
        {
            string source = CreateDecisionWorkbook("successive-source.xlsx");
            string firstPath = Path.Combine(_directory, "successive-1.xlsx");
            string secondPath = Path.Combine(_directory, "successive-2.xlsx");
            byte[] sourceBytes = File.ReadAllBytes(source);
            ReviewWorkbookReadResult root = ReviewWorkbookFile.ReadAndValidate(source);
            ReviewDecisionRequest firstRequest = Request(
                root.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 30, 0, TimeSpan.Zero),
                "successive-first");
            ReviewWorkbookFile.CreateDecisionRevision(source, firstPath, firstRequest);
            byte[] firstBytes = File.ReadAllBytes(firstPath);
            ReviewWorkbookReadResult first = ReviewWorkbookFile.ReadAndValidate(firstPath);

            ReviewDecisionRequest secondRequest = ReviewDecisionRequest.ClearSubject(
                first.ProjectId,
                first.DecisionGeneration!.BaseGenerationId,
                first.DecisionGeneration.MappingGenerationId,
                first.DecisionGeneration.GenerationId,
                SubjectId,
                new ReviewerSnapshot("Reviewer Two", "Office Two", new[] { "confirmed" }),
                new DateTimeOffset(2026, 7, 17, 8, 35, 0, TimeSpan.Zero),
                runId: "successive-second",
                addInBuild: "test-build-2",
                note: "clear after review");
            ReviewWorkbookFile.CreateDecisionRevision(firstPath, secondPath, secondRequest);

            ReviewWorkbookReadResult second = ReviewWorkbookFile.ReadAndValidate(secondPath);
            AssertRevisionIdentityPreserved(root, first);
            AssertRevisionIdentityPreserved(first, second);
            Assert.Equal(2, second.DecisionGeneration!.Events.Count);
            Assert.Equal(
                new[] { firstRequest.RequestId, secondRequest.RequestId },
                second.DecisionGeneration.Events.Select(item => item.RequestId));
            Assert.Equal(first.DecisionGeneration.GenerationId, second.DecisionGeneration.PriorGenerationId);
            Assert.Equal(
                ReviewDecisionStatus.Unreviewed,
                second.DecisionGeneration.GetSubjectState(SubjectId).Status);
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            Assert.Equal(firstBytes, File.ReadAllBytes(firstPath));
            Assert.NotEqual(root.DecisionSha256, first.DecisionSha256);
            Assert.NotEqual(first.DecisionSha256, second.DecisionSha256);
            AssertNoTemporary(firstPath);
            AssertNoTemporary(secondPath);
        }

        [Fact]
        public void Batch_publishes_one_revision_with_an_exact_chained_event_sequence()
        {
            string source = CreateDecisionWorkbook("batch-source.xlsx");
            string destination = Path.Combine(_directory, "batch-revision.xlsx");
            byte[] sourceBytes = File.ReadAllBytes(source);
            ReviewWorkbookReadResult root = ReviewWorkbookFile.ReadAndValidate(source);
            PackageSnapshot sourcePackage = PackageSnapshot.Read(source);
            ReviewDecisionRequest first = Request(
                root.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 36, 0, TimeSpan.Zero),
                "batch-first");
            ReviewDecisionGeneration intermediate = ReviewDecisionEngine.Apply(
                root.DecisionGeneration,
                first);
            ReviewDecisionRequest second = ReviewDecisionRequest.ClearSubject(
                intermediate.ProjectId,
                intermediate.BaseGenerationId,
                intermediate.MappingGenerationId,
                intermediate.GenerationId,
                SubjectId,
                ReviewerSnapshot.Empty,
                new DateTimeOffset(2026, 7, 17, 8, 37, 0, TimeSpan.Zero),
                runId: "batch-second");

            ReviewWorkbookFile.CreateDecisionRevisionBatch(
                source,
                destination,
                new[] { first, second });

            ReviewWorkbookReadResult revised = ReviewWorkbookFile.ReadAndValidate(destination);
            AssertRevisionIdentityPreserved(root, revised);
            Assert.Equal(
                new[] { first.RequestId, second.RequestId },
                revised.DecisionGeneration!.Events.Select(item => item.RequestId));
            Assert.Equal(intermediate.GenerationId, revised.DecisionGeneration.PriorGenerationId);
            Assert.Equal(
                ReviewDecisionStatus.Unreviewed,
                revised.DecisionGeneration.GetSubjectState(SubjectId).Status);
            Assert.Equal(sourceBytes, File.ReadAllBytes(source));
            AssertOnlyDecisionAndManifestPartsChanged(
                sourcePackage,
                PackageSnapshot.Read(destination));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Hydrated_batch_derives_from_the_exact_retained_source_occurrence()
        {
            string source = CreateDecisionWorkbook("hydrated-source.xlsx");
            string destination = Path.Combine(_directory, "hydrated-revision.xlsx");
            ReviewWorkbookHydrationSession session =
                ReviewWorkbookHydrationService.Hydrate(source);
            ReviewDecisionRequest request = Request(
                session.ReadResult.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 38, 0, TimeSpan.Zero),
                "hydrated-source");

            ReviewWorkbookFile.CreateDecisionRevisionBatch(
                session,
                destination,
                new[] { request });

            ReviewWorkbookReadResult revised = ReviewWorkbookFile.ReadAndValidate(destination);
            Assert.Equal(request.RequestId, Assert.Single(revised.DecisionGeneration!.Events).RequestId);
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Hydrated_batch_rejects_a_byte_identical_path_replacement_on_windows()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

            string source = CreateDecisionWorkbook("hydrated-selected.xlsx");
            string destination = Path.Combine(_directory, "hydrated-replaced-revision.xlsx");
            ReviewWorkbookHydrationSession session =
                ReviewWorkbookHydrationService.Hydrate(source);
            ReviewDecisionRequest request = Request(
                session.ReadResult.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 39, 0, TimeSpan.Zero),
                "hydrated-replaced");
            string replacement = Path.Combine(_directory, "hydrated-identical-copy.xlsx");
            File.Copy(source, replacement);
            File.Delete(source);
            File.Move(replacement, source);

            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                ReviewWorkbookFile.CreateDecisionRevisionBatch(
                    session,
                    destination,
                    new[] { request }));

            Assert.Contains("source occurrence", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(destination));
            AssertNoTemporary(destination);
        }

        [Fact]
        public void Requires_xlsx_source_and_distinct_full_paths()
        {
            string source = CreateDecisionWorkbook("path-source.xlsx");
            ReviewWorkbookReadResult current = ReviewWorkbookFile.ReadAndValidate(source);
            ReviewDecisionRequest request = Request(
                current.DecisionGeneration!,
                ReviewDecisionStatus.NeedsInvestigation,
                new DateTimeOffset(2026, 7, 17, 8, 40, 0, TimeSpan.Zero),
                "paths");
            string nonXlsx = Path.Combine(_directory, "source.bin");
            File.Copy(source, nonXlsx);

            Assert.Throws<ArgumentException>(() => ReviewWorkbookFile.CreateDecisionRevision(
                nonXlsx,
                Path.Combine(_directory, "from-bin.xlsx"),
                request));
            Assert.Throws<ArgumentException>(() => ReviewWorkbookFile.CreateDecisionRevision(
                source,
                source,
                request));
            Assert.Equal(
                File.ReadAllBytes(source),
                File.ReadAllBytes(nonXlsx));
            AssertNoTemporary(Path.Combine(_directory, "from-bin.xlsx"));
        }

        public void Dispose()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        private string CreateDecisionWorkbook(
            string fileName,
            string projectId = "project-copy-on-write")
        {
            ReviewDecisionGeneration root = CreateDecisionRoot(projectId);
            string path = Path.Combine(_directory, fileName);
            ReviewWorkbookFile.CreateNew(
                path,
                new ReviewWorkbookState(root.ProjectId, "{\"state\":\"copy-on-write-α\",\"revision\":0}")
                {
                    DecisionGeneration = root,
                    DecisionEvidenceGenerationId = root.Catalog.EvidenceGenerationId,
                    DecisionMappingGenerationId = root.MappingGenerationId,
                },
                new ReviewWorkbookWriteOptions { StateChunkSizeBytes = 9 });
            return path;
        }

        private static ReviewDecisionGeneration CreateDecisionRoot(
            string projectId = "project-copy-on-write")
        {
            var evidence = new ReviewEvidenceGeneration(
                projectId,
                "mapping-copy-on-write",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("candidate", 0, new string('a', 64)),
                    new ReviewEvidenceRevisionIdentity("master", 1, new string('b', 64)),
                });
            var subject = new ReviewDecisionSubject(
                SubjectId,
                "logical-sheet:value:A1",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            Observation candidate = Observation(
                subject.SubjectId,
                "candidate",
                "candidate value",
                new string('1', 64));
            Observation master = Observation(
                subject.SubjectId,
                "master",
                "master value",
                new string('2', 64));
            var catalog = new ReviewDecisionCatalog(
                evidence,
                "master",
                new[] { subject },
                new[] { candidate, master });
            return ReviewDecisionEngine.Create(catalog);
        }

        private static Observation Observation(
            string subjectId,
            string revisionId,
            string value,
            string contentHash) =>
            new Observation(
                ComparisonIds.Observation(
                    subjectId,
                    revisionId,
                    ObservationExactType.String,
                    contentHash),
                subjectId,
                revisionId,
                ObservationExactType.String,
                contentHash,
                value: new CanonicalScalar(ObservationExactType.String, value),
                displayText: value,
                provenance: new ObservationProvenance(
                    revisionId + ".xlsx",
                    "Input",
                    "A1",
                    1,
                    1,
                    "logical-sheet"));

        private static ReviewDecisionRequest Request(
            ReviewDecisionGeneration generation,
            ReviewDecisionStatus status,
            DateTimeOffset createdUtc,
            string runId) =>
            ReviewDecisionRequest.ForSubject(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.GenerationId,
                SubjectId,
                status,
                new ReviewerSnapshot("Reviewer One", "Office One", new[] { "confirmed" }),
                createdUtc,
                runId: runId,
                addInBuild: "test-build-1",
                note: "copy-on-write decision");

        private static Exception? PublishAndCapture(
            ManualResetEventSlim start,
            string source,
            string destination,
            ReviewDecisionRequest request)
        {
            start.Wait();
            try
            {
                ReviewWorkbookFile.CreateDecisionRevision(source, destination, request);
                return null;
            }
            catch (Exception exception)
            {
                return exception;
            }
        }

        private static void AssertRevisionIdentityPreserved(
            ReviewWorkbookReadResult before,
            ReviewWorkbookReadResult after)
        {
            Assert.Equal(before.ProjectId, after.ProjectId);
            Assert.Equal(before.StateJson, after.StateJson);
            Assert.Equal(before.StateSha256, after.StateSha256, ignoreCase: true);
            Assert.Equal(before.ProjectionSha256, after.ProjectionSha256, ignoreCase: true);
            Assert.Equal(before.ManifestItemId, after.ManifestItemId);
            Assert.Equal(before.DecisionItemId, after.DecisionItemId);
            Assert.Equal(before.ChunkItemIds, after.ChunkItemIds);
            Assert.Equal(before.DecisionGeneration!.CatalogId, after.DecisionGeneration!.CatalogId);
            Assert.Equal(before.DecisionGeneration.BaseGenerationId, after.DecisionGeneration.BaseGenerationId);
            Assert.Equal(before.DecisionGeneration.MappingGenerationId, after.DecisionGeneration.MappingGenerationId);
        }

        private static void AssertOnlyDecisionAndManifestPartsChanged(
            PackageSnapshot before,
            PackageSnapshot after)
        {
            Assert.Equal(before.Entries.Keys.OrderBy(item => item), after.Entries.Keys.OrderBy(item => item));
            foreach (KeyValuePair<string, byte[]> entry in before.Entries)
            {
                if (string.Equals(entry.Key, before.DecisionEntry, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(entry.Key, before.ManifestEntry, StringComparison.OrdinalIgnoreCase))
                    continue;
                Assert.Equal(entry.Value, after.Entries[entry.Key]);
            }
            Assert.NotEqual(before.Entries[before.DecisionEntry], after.Entries[after.DecisionEntry]);
            Assert.NotEqual(before.Entries[before.ManifestEntry], after.Entries[after.ManifestEntry]);
        }

        private static void AssertManifestChangedOnlyAtDigest(
            PackageSnapshot before,
            PackageSnapshot after,
            string oldDigest,
            string newDigest)
        {
            byte[] expected = (byte[])before.Entries[before.ManifestEntry].Clone();
            byte[] oldValue = Encoding.ASCII.GetBytes(oldDigest);
            byte[] newValue = Encoding.ASCII.GetBytes(newDigest.ToUpperInvariant());
            int offset = FindUnique(expected, oldValue);
            Buffer.BlockCopy(newValue, 0, expected, offset, newValue.Length);
            Assert.Equal(expected, after.Entries[after.ManifestEntry]);
        }

        private static int FindUnique(byte[] bytes, byte[] value)
        {
            int found = -1;
            for (int offset = 0; offset <= bytes.Length - value.Length; offset++)
            {
                if (!bytes.AsSpan(offset, value.Length).SequenceEqual(value)) continue;
                Assert.Equal(-1, found);
                found = offset;
            }
            Assert.True(found >= 0);
            return found;
        }

        private static int CountOccurrences(byte[] bytes, byte[] value)
        {
            int count = 0;
            for (int offset = 0; offset <= bytes.Length - value.Length; offset++)
            {
                if (bytes.AsSpan(offset, value.Length).SequenceEqual(value)) count++;
            }
            return count;
        }

        private static void MutateManifestDigest(string path, string digest)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            CustomXmlPart manifestPart = document.WorkbookPart!.CustomXmlParts
                .Single(part => RootLocalName(part) == "reviewManifest");
            XDocument manifest;
            using (Stream input = manifestPart.GetStream(FileMode.Open, FileAccess.Read))
                manifest = XDocument.Load(input);
            manifest.Root!.SetAttributeValue("decisionSha256", digest);
            using Stream output = manifestPart.GetStream(FileMode.Create, FileAccess.Write);
            manifest.Save(output, SaveOptions.DisableFormatting);
        }

        private void AssertNoTemporary(string destination)
        {
            string pattern = "." + Path.GetFileName(destination) + ".*.tmp";
            Assert.Empty(Directory.GetFiles(_directory, pattern));
        }

        private static string RootLocalName(CustomXmlPart part)
        {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream).Root!.Name.LocalName;
        }

        private sealed class PackageSnapshot
        {
            private PackageSnapshot(
                IReadOnlyDictionary<string, byte[]> entries,
                string manifestEntry,
                string decisionEntry)
            {
                Entries = entries;
                ManifestEntry = manifestEntry;
                DecisionEntry = decisionEntry;
            }

            public IReadOnlyDictionary<string, byte[]> Entries { get; }
            public string ManifestEntry { get; }
            public string DecisionEntry { get; }

            public static PackageSnapshot Read(string path)
            {
                string manifestEntry;
                string decisionEntry;
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart!;
                    manifestEntry = workbookPart.CustomXmlParts
                        .Single(part => RootLocalName(part) == "reviewManifest")
                        .Uri.ToString().TrimStart('/');
                    decisionEntry = workbookPart.CustomXmlParts
                        .Single(part => RootLocalName(part) == "decisionState")
                        .Uri.ToString().TrimStart('/');
                }

                var entries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                using (ZipArchive archive = ZipFile.OpenRead(path))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        using Stream input = entry.Open();
                        using var memory = new MemoryStream();
                        input.CopyTo(memory);
                        entries.Add(entry.FullName, memory.ToArray());
                    }
                }
                return new PackageSnapshot(entries, manifestEntry, decisionEntry);
            }
        }
    }
}
