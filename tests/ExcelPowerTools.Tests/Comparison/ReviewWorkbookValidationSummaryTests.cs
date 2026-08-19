using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewWorkbookValidationSummaryTests : IDisposable
    {
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.ReviewSummary.Tests",
            Guid.NewGuid().ToString("N"));

        public ReviewWorkbookValidationSummaryTests()
        {
            Directory.CreateDirectory(_directory);
        }

        [Fact]
        public void Validate_returns_only_bounded_metadata_from_the_integrity_checked_package()
        {
            const string secretSourcePath = @"C:\Sensitive\Input\person-copy.xlsx";
            string stateJson = "{\"schema\":\"summary-test\",\"sourcePath\":\"" +
                               secretSourcePath.Replace("\\", "\\\\") +
                               "\",\"note\":\"π-grid\"}";
            string path = Path.Combine(_directory, "Validated Review.xlsx");
            ReviewWorkbookFile.CreateNew(
                path,
                new ReviewWorkbookState("project-summary", stateJson),
                new ReviewWorkbookWriteOptions { StateChunkSizeBytes = 17 });
            File.SetLastWriteTimeUtc(path, new DateTime(2026, 7, 17, 9, 8, 6, DateTimeKind.Utc));

            ReviewWorkbookReadResult validated = ReviewWorkbookFile.ReadAndValidate(path);
            ReviewWorkbookValidationSummary summary = ReviewWorkbookValidationSummaryService.Validate(path);

            Assert.Equal(Path.GetFullPath(path), summary.FullPath);
            Assert.Equal("Validated Review.xlsx", summary.FileName);
            Assert.Equal(new FileInfo(path).Length, summary.FileByteLength);
            Assert.Equal(new DateTimeOffset(File.GetLastWriteTimeUtc(path)), summary.ModifiedUtc);
            Assert.Equal("project-summary", validated.ProjectId);
            Assert.StartsWith("opaque:", summary.ProjectId);
            Assert.Equal(71, summary.ProjectId.Length);
            Assert.Equal(summary.ProjectId, ReviewWorkbookValidationSummaryService.Validate(path).ProjectId);
            Assert.Equal(
                validated.StateSha256.Substring(0, ReviewWorkbookValidationSummaryService.HashPrefixLength),
                summary.StateSha256Prefix);
            Assert.Equal(
                validated.ProjectionSha256.Substring(0, ReviewWorkbookValidationSummaryService.HashPrefixLength),
                summary.ProjectionSha256Prefix);
            Assert.Equal(validated.ChunkCount, summary.StateChunkCount);
            Assert.True(summary.StateChunkCount > 1);
            Assert.Equal(new UTF8Encoding(false, true).GetByteCount(stateJson), summary.StateByteCount);
            Assert.False(summary.HasDecisionSidecar);
            Assert.Equal(0, summary.DecisionEvidenceRevisionCount);
            Assert.Equal(0, summary.DecisionSubjectCount);
            Assert.Equal(0, summary.DecisionObservationCount);
            Assert.Equal(0, summary.ActiveDecisionCount);
            Assert.Empty(summary.DecisionCatalogIdPrefix);
            Assert.Empty(summary.DecisionGenerationIdPrefix);

            Assert.Null(typeof(ReviewWorkbookValidationSummary).GetProperty("StateJson"));
            string[] exposedStrings = typeof(ReviewWorkbookValidationSummary)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.PropertyType == typeof(string))
                .Select(property => (string)property.GetValue(summary)!)
                .ToArray();
            Assert.DoesNotContain(exposedStrings, value => value.Contains(secretSourcePath, StringComparison.Ordinal));
        }

        [Fact]
        public void Validate_reports_only_safe_counts_and_prefixes_for_a_validated_decision_sidecar()
        {
            const string secretStatePath = @"C:\Sensitive\State\input.xlsx";
            const string secretDecisionSource = "person-secret-emailed-copy.xlsx";
            ReviewDecisionGeneration decision = DecisionRoot(secretDecisionSource);
            string path = Path.Combine(_directory, "Review with decisions.xlsx");
            ReviewWorkbookFile.CreateNew(
                path,
                new ReviewWorkbookState(
                    decision.ProjectId,
                    "{\"sourcePath\":\"" + secretStatePath.Replace("\\", "\\\\") + "\"}")
                {
                    DecisionGeneration = decision,
                    DecisionEvidenceGenerationId = decision.Catalog.EvidenceGenerationId,
                    DecisionMappingGenerationId = decision.MappingGenerationId,
                });

            ReviewWorkbookValidationSummary summary = ReviewWorkbookValidationSummaryService.Validate(path);

            Assert.True(summary.HasDecisionSidecar);
            Assert.Equal(2, summary.DecisionEvidenceRevisionCount);
            Assert.Equal(1, summary.DecisionSubjectCount);
            Assert.Equal(2, summary.DecisionObservationCount);
            Assert.Equal(0, summary.ActiveDecisionCount);
            Assert.Equal(
                decision.CatalogId.Substring(
                    0,
                    ReviewWorkbookValidationSummaryService.HashPrefixLength),
                summary.DecisionCatalogIdPrefix);
            Assert.Equal(
                decision.GenerationId.Substring(
                    0,
                    ReviewWorkbookValidationSummaryService.HashPrefixLength),
                summary.DecisionGenerationIdPrefix);

            Assert.Null(typeof(ReviewWorkbookValidationSummary).GetProperty("DecisionGeneration"));
            Assert.Null(typeof(ReviewWorkbookValidationSummary).GetProperty("DecisionCatalogId"));
            Assert.Null(typeof(ReviewWorkbookValidationSummary).GetProperty("DecisionGenerationId"));
            string[] exposedStrings = typeof(ReviewWorkbookValidationSummary)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(property => property.PropertyType == typeof(string))
                .Select(property => (string)property.GetValue(summary)!)
                .ToArray();
            Assert.DoesNotContain(exposedStrings, value =>
                value.Contains(secretStatePath, StringComparison.Ordinal) ||
                value.Contains(secretDecisionSource, StringComparison.Ordinal));
        }

        [Fact]
        public void Validate_preserves_an_opaque_sha256_project_identifier()
        {
            string projectId = StableId.Compute("review-summary-test-project-v1", "safe-project");
            string path = Path.Combine(_directory, "Review stable project.xlsx");
            ReviewWorkbookFile.CreateNew(path, new ReviewWorkbookState(projectId, "{}"));

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            ReviewWorkbookValidationSummary summary = ReviewWorkbookValidationSummaryService.Validate(path);

            Assert.Equal(projectId, read.ProjectId);
            Assert.Equal(projectId, summary.ProjectId);
        }

        [Fact]
        public void Validate_redacts_path_like_and_very_long_project_identifiers_but_read_retains_them()
        {
            string pathLikeProjectId = @"C:\Sensitive\Customer A\emailed-copy.xlsx";
            string veryLongProjectId = pathLikeProjectId + "|" + new string('X', 3_000);

            string pathLikeSafe = AssertProjectIdIsRedacted(
                "Review path project.xlsx",
                pathLikeProjectId);
            string veryLongSafe = AssertProjectIdIsRedacted(
                "Review long project.xlsx",
                veryLongProjectId);

            Assert.NotEqual(pathLikeSafe, veryLongSafe);
        }

        [Fact]
        public void Validate_fails_closed_for_a_non_review_package()
        {
            string path = Path.Combine(_directory, "not-a-review.xlsx");
            File.WriteAllBytes(path, new byte[] { 0x50, 0x4B, 0x03, 0x04, 1, 2, 3, 4 });

            Assert.Throws<InvalidDataException>(() =>
                ReviewWorkbookValidationSummaryService.Validate(path));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void Validate_rejects_a_missing_path(string? path)
        {
            Assert.Throws<ArgumentException>(() =>
                ReviewWorkbookValidationSummaryService.Validate(path!));
        }

        private static ReviewDecisionGeneration DecisionRoot(string sourceLabel)
        {
            var evidence = new ReviewEvidenceGeneration(
                "project-summary-decisions",
                "mapping-summary-decisions",
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
            string candidateHash = new string('C', 64);
            string masterHash = new string('D', 64);
            var observations = new[]
            {
                new Observation(
                    ComparisonIds.Observation(
                        subject.SubjectId,
                        "candidate",
                        ObservationExactType.String,
                        candidateHash),
                    subject.SubjectId,
                    "candidate",
                    ObservationExactType.String,
                    candidateHash,
                    value: new CanonicalScalar(ObservationExactType.String, "candidate"),
                    provenance: new ObservationProvenance(sourceLabel)),
                new Observation(
                    ComparisonIds.Observation(
                        subject.SubjectId,
                        "master",
                        ObservationExactType.String,
                        masterHash),
                    subject.SubjectId,
                    "master",
                    ObservationExactType.String,
                    masterHash,
                    value: new CanonicalScalar(ObservationExactType.String, "master"),
                    provenance: new ObservationProvenance("master.xlsx")),
            };
            return ReviewDecisionEngine.Create(new ReviewDecisionCatalog(
                evidence,
                "master",
                new[] { subject },
                observations));
        }

        private string AssertProjectIdIsRedacted(string fileName, string rawProjectId)
        {
            string path = Path.Combine(_directory, fileName);
            ReviewWorkbookFile.CreateNew(path, new ReviewWorkbookState(rawProjectId, "{}"));

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            ReviewWorkbookValidationSummary first = ReviewWorkbookValidationSummaryService.Validate(path);
            ReviewWorkbookValidationSummary second = ReviewWorkbookValidationSummaryService.Validate(path);

            Assert.Equal(rawProjectId, read.ProjectId);
            Assert.NotEqual(rawProjectId, first.ProjectId);
            Assert.StartsWith("opaque:", first.ProjectId);
            Assert.Equal(71, first.ProjectId.Length);
            Assert.Equal(first.ProjectId, second.ProjectId);
            Assert.DoesNotContain("Sensitive", first.ProjectId, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("emailed-copy.xlsx", first.ProjectId, StringComparison.OrdinalIgnoreCase);
            return first.ProjectId;
        }

        public void Dispose()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }
    }
}
