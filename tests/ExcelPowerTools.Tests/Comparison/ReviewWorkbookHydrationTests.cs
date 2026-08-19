using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewWorkbookHydrationTests : IDisposable
    {
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.ReviewHydration.Tests",
            Guid.NewGuid().ToString("N"));

        public ReviewWorkbookHydrationTests()
        {
            Directory.CreateDirectory(_directory);
        }

        [Fact]
        public void Hydrate_retains_exact_package_state_and_privacy_safe_summary_without_decisions()
        {
            const string secret = @"C:\Sensitive\Emailed Copies\input.xlsx";
            string stateJson = "{\"sourcePath\":\"" + secret.Replace("\\", "\\\\") + "\"}";
            string path = Path.Combine(_directory, "Review no decisions.xlsx");
            ReviewWorkbookFile.CreateNew(path, new ReviewWorkbookState("project-hydrate", stateJson));
            File.SetLastWriteTimeUtc(path, new DateTime(2026, 7, 17, 10, 11, 12, DateTimeKind.Utc));

            ReviewWorkbookHydrationSession session = ReviewWorkbookHydrationService.Hydrate(path);

            Assert.Equal(Path.GetFullPath(path), session.FullPath);
            Assert.Equal(Path.GetFileName(path), session.FileName);
            Assert.Equal(new FileInfo(path).Length, session.FileByteLength);
            Assert.Equal(new DateTimeOffset(File.GetLastWriteTimeUtc(path)), session.ModifiedUtc);
            Assert.Equal(FileSha256(path), session.PackageSha256);
            Assert.Equal(64, session.PackageSha256.Length);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Assert.True(session.HasStableOccurrenceToken);
                Assert.Equal(25, session.FileOccurrenceToken.Length);
            }
            Assert.Equal("project-hydrate", session.ReadResult.ProjectId);
            Assert.Equal(stateJson, session.ReadResult.StateJson);
            Assert.False(session.ReadResult.HasDecisionState);
            Assert.Equal(session.FullPath, session.Summary.FullPath);
            Assert.False(session.Summary.HasDecisionSidecar);

            Assert.Null(typeof(ReviewWorkbookValidationSummary).GetProperty("StateJson"));
            string[] summaryStrings = typeof(ReviewWorkbookValidationSummary)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.PropertyType == typeof(string))
                .Select(property => (string)property.GetValue(session.Summary)!)
                .ToArray();
            Assert.DoesNotContain(summaryStrings, value => value.Contains(secret, StringComparison.Ordinal));

            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        }

        [Fact]
        public void Hydrate_retains_the_replayed_advanced_decision_generation()
        {
            const string secretSource = "emailed-secret-copy.xlsx";
            ReviewDecisionGeneration expected = AdvancedDecisionGeneration(secretSource);
            string path = Path.Combine(_directory, "Review decisions.xlsx");
            ReviewWorkbookFile.CreateNew(
                path,
                new ReviewWorkbookState(expected.ProjectId, "{\"safe\":true}")
                {
                    DecisionGeneration = expected,
                    DecisionEvidenceGenerationId = expected.Catalog.EvidenceGenerationId,
                    DecisionMappingGenerationId = expected.MappingGenerationId,
                });

            ReviewWorkbookHydrationSession session = ReviewWorkbookHydrationService.Hydrate(path);

            Assert.True(session.ReadResult.HasDecisionState);
            Assert.NotNull(session.ReadResult.DecisionGeneration);
            Assert.Equal(expected.CatalogId, session.ReadResult.DecisionGeneration.CatalogId);
            Assert.Equal(expected.GenerationId, session.ReadResult.DecisionGeneration.GenerationId);
            Assert.Equal(expected.PriorGenerationId, session.ReadResult.DecisionGeneration.PriorGenerationId);
            Assert.Single(session.ReadResult.DecisionGeneration.Events);
            Assert.Single(session.ReadResult.DecisionGeneration.ActiveDecisions);
            Assert.True(session.Summary.HasDecisionSidecar);
            Assert.Equal(1, session.Summary.ActiveDecisionCount);
            string[] summaryStrings = typeof(ReviewWorkbookValidationSummary)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.PropertyType == typeof(string))
                .Select(property => (string)property.GetValue(session.Summary)!)
                .ToArray();
            Assert.DoesNotContain(summaryStrings, value => value.Contains(secretSource, StringComparison.Ordinal));
        }

        [Fact]
        public void RevalidateCurrentOccurrence_returns_a_fresh_matching_session_and_releases_its_handle()
        {
            string path = CreateReview("unchanged.xlsx", "project-unchanged", "{\"revision\":1}");
            ReviewWorkbookHydrationSession original = ReviewWorkbookHydrationService.Hydrate(path);

            ReviewWorkbookHydrationSession current = original.RevalidateCurrentOccurrence();

            Assert.NotSame(original, current);
            Assert.Equal(original.PackageSha256, current.PackageSha256);
            Assert.Equal(original.ReadResult.StateSha256, current.ReadResult.StateSha256);
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.CanWrite);
        }

        [Fact]
        public void RevalidateCurrentOccurrence_accepts_a_timestamp_only_change()
        {
            string path = CreateReview("timestamp-only.xlsx", "project-timestamp", "{\"revision\":1}");
            ReviewWorkbookHydrationSession original = ReviewWorkbookHydrationService.Hydrate(path);
            DateTime changedTimestamp = original.ModifiedUtc.UtcDateTime.AddMinutes(2);
            File.SetLastWriteTimeUtc(path, changedTimestamp);

            ReviewWorkbookHydrationSession current = original.RevalidateCurrentOccurrence();

            Assert.Equal(original.PackageSha256, current.PackageSha256);
            Assert.Equal(original.FileOccurrenceToken, current.FileOccurrenceToken);
            Assert.NotEqual(original.ModifiedUtc, current.ModifiedUtc);
            Assert.Equal(new DateTimeOffset(File.GetLastWriteTimeUtc(path)), current.ModifiedUtc);
        }

        [Fact]
        public void RevalidateCurrentOccurrence_detects_a_byte_identical_replacement_by_handle_identity_on_windows()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

            string path = CreateReview("identical-selected.xlsx", "project-identical", "{\"revision\":1}");
            ReviewWorkbookHydrationSession original = ReviewWorkbookHydrationService.Hydrate(path);
            string replacement = Path.Combine(_directory, "identical-copy.xlsx");
            File.Copy(path, replacement);
            File.SetLastWriteTimeUtc(replacement, original.ModifiedUtc.UtcDateTime);

            File.Delete(path);
            File.Move(replacement, path);
            ReviewWorkbookHydrationSession current = ReviewWorkbookHydrationService.Hydrate(path);

            Assert.Equal(original.PackageSha256, current.PackageSha256);
            Assert.Equal(original.ReadResult.ProjectId, current.ReadResult.ProjectId);
            Assert.Equal(original.ReadResult.StateSha256, current.ReadResult.StateSha256);
            Assert.Equal(original.ModifiedUtc, current.ModifiedUtc);
            Assert.True(original.HasStableOccurrenceToken);
            Assert.True(current.HasStableOccurrenceToken);
            Assert.NotEqual(original.FileOccurrenceToken, current.FileOccurrenceToken);
            Assert.Throws<InvalidDataException>(() => original.RevalidateCurrentOccurrence());
        }

        [Fact]
        public void RevalidateCurrentOccurrence_detects_in_place_mutation_with_a_restored_timestamp()
        {
            string path = CreateReview("mutated.xlsx", "project-mutated", "{\"revision\":1}");
            ReviewWorkbookHydrationSession original = ReviewWorkbookHydrationService.Hydrate(path);
            DateTime timestamp = original.ModifiedUtc.UtcDateTime;
            string changed = CreateReview("mutated-source.xlsx", "project-mutated", "{\"revision\":2}");
            using (var source = new FileStream(changed, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var target = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                source.CopyTo(target);
                target.Flush(true);
            }
            File.SetLastWriteTimeUtc(path, timestamp);
            ReviewWorkbookHydrationSession current = ReviewWorkbookHydrationService.Hydrate(path);

            Assert.Equal(original.ModifiedUtc, new DateTimeOffset(File.GetLastWriteTimeUtc(path)));
            Assert.NotEqual(original.PackageSha256, current.PackageSha256);
            Assert.NotEqual(original.ReadResult.StateSha256, current.ReadResult.StateSha256);
            Assert.Throws<InvalidDataException>(() => original.RevalidateCurrentOccurrence());
        }

        [Fact]
        public void RevalidateCurrentOccurrence_detects_a_valid_path_replacement_even_when_timestamp_matches()
        {
            string path = CreateReview("selected.xlsx", "project-original", "{\"revision\":1}");
            ReviewWorkbookHydrationSession original = ReviewWorkbookHydrationService.Hydrate(path);
            string replacement = CreateReview("replacement.xlsx", "project-replacement", "{\"revision\":2}");

            File.Delete(path);
            File.Move(replacement, path);
            File.SetLastWriteTimeUtc(path, original.ModifiedUtc.UtcDateTime);
            ReviewWorkbookHydrationSession current = ReviewWorkbookHydrationService.Hydrate(path);

            Assert.Equal(original.ModifiedUtc, current.ModifiedUtc);
            Assert.NotEqual(original.PackageSha256, current.PackageSha256);
            Assert.NotEqual(original.ReadResult.ProjectId, current.ReadResult.ProjectId);
            Assert.Throws<InvalidDataException>(() => original.RevalidateCurrentOccurrence());
        }

        [Fact]
        public void Hydrate_rejects_a_package_above_the_hash_bound_without_retaining_a_handle()
        {
            string path = Path.Combine(_directory, "oversized.xlsx");
            using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                output.SetLength(ReviewWorkbookFile.MaxReviewPackageBytes + 1);

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookHydrationService.Hydrate(path));
            File.Delete(path);
            Assert.False(File.Exists(path));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Hydrate_rejects_a_missing_path(string? path)
        {
            Assert.Throws<ArgumentException>(() => ReviewWorkbookHydrationService.Hydrate(path!));
        }

        private string CreateReview(string fileName, string projectId, string stateJson)
        {
            string path = Path.Combine(_directory, fileName);
            ReviewWorkbookFile.CreateNew(path, new ReviewWorkbookState(projectId, stateJson));
            return path;
        }

        private static ReviewDecisionGeneration AdvancedDecisionGeneration(string sourceLabel)
        {
            var evidence = new ReviewEvidenceGeneration(
                "project-hydrate-decisions",
                "mapping-hydrate-decisions",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("candidate", 0, new string('A', 64)),
                    new ReviewEvidenceRevisionIdentity("master", 1, new string('B', 64)),
                });
            var subject = new ReviewDecisionSubject(
                "subject-value",
                "logical-value",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            var observations = new[]
            {
                ObservationFor(subject.SubjectId, "candidate", "candidate", new string('C', 64), sourceLabel),
                ObservationFor(subject.SubjectId, "master", "master", new string('D', 64), "master.xlsx"),
            };
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(new ReviewDecisionCatalog(
                evidence,
                "master",
                new[] { subject },
                observations));
            return ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    subject.SubjectId,
                    ReviewDecisionStatus.Equivalent,
                    new ReviewerSnapshot("Hydration reviewer"),
                    new DateTimeOffset(2026, 7, 17, 11, 12, 13, TimeSpan.Zero),
                    note: "reviewed"));
        }

        private static Observation ObservationFor(
            string subjectId,
            string revisionId,
            string value,
            string contentHash,
            string sourceLabel) =>
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
                provenance: new ObservationProvenance(sourceLabel));

        private static string FileSha256(string path) =>
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

        public void Dispose()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }
    }
}
