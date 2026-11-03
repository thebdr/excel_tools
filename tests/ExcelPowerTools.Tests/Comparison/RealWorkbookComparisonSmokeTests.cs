using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;
using Xunit;
using Xunit.Abstractions;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class RealWorkbookComparisonSmokeTests
    {
        internal const string BeforePathVariable = "EXCEL_POWER_TOOLS_REAL_COMPARISON_BEFORE";
        internal const string AfterPathVariable = "EXCEL_POWER_TOOLS_REAL_COMPARISON_AFTER";

        private readonly ITestOutputHelper _output;

        public RealWorkbookComparisonSmokeTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [RealWorkbookComparisonFact]
        [Trait("Category", "ManualRealWorkbook")]
        public void Selected_real_revision_pair_completes_advanced_alignment_without_source_writes()
        {
            string beforePath = Environment.GetEnvironmentVariable(BeforePathVariable) ?? string.Empty;
            string afterPath = Environment.GetEnvironmentVariable(AfterPathVariable) ?? string.Empty;
            Assert.True(File.Exists(beforePath), "Before workbook not found: " + beforePath);
            Assert.True(File.Exists(afterPath), "After workbook not found: " + afterPath);

            DateTime beforeWriteTime = File.GetLastWriteTimeUtc(beforePath);
            DateTime afterWriteTime = File.GetLastWriteTimeUtc(afterPath);
            byte[] beforeBytes = File.ReadAllBytes(beforePath);
            byte[] afterBytes = File.ReadAllBytes(afterPath);
            string beforeHash = Hash(beforeBytes);
            string afterHash = Hash(afterBytes);

            OpenXmlWorkbookCompactionResult before = OpenXmlWorkbookSnapshotCompactor.CompactToUsefulRanges(
                OpenXmlWorkbookSnapshotAdapter.FromBytes("before", beforeBytes));
            OpenXmlWorkbookCompactionResult after = OpenXmlWorkbookSnapshotCompactor.CompactToUsefulRanges(
                OpenXmlWorkbookSnapshotAdapter.FromBytes("after", afterBytes));
            ReportBounds("Before", before);
            ReportBounds("After", after);

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-real-workbook-smoke",
                "mapping-real-workbook-smoke",
                new[]
                {
                    new OpenXmlProjectRevision(
                        "before",
                        "",
                        0,
                        before.Snapshot,
                        isTrustedBaseline: true),
                    new OpenXmlProjectRevision(
                        "after",
                        "before",
                        1,
                        after.Snapshot,
                        isMaster: true),
                });

            string findings = string.Join(
                Environment.NewLine,
                result.Findings.Select(item =>
                    item.Severity + " | " + item.Code + " | " + item.SheetName + " | " + item.Message));
            _output.WriteLine("Complete: " + result.IsComplete);
            foreach (OpenXmlProjectSheetAlignment alignment in
                     result.Edges.SelectMany(edge => edge.SheetAlignments))
            {
                _output.WriteLine(
                    "Sheet " + alignment.BeforeSheetName + " -> " + alignment.AfterSheetName +
                    ": complete=" + alignment.IsComplete +
                    "; before header=" + DescribeHeader(alignment.BeforeHeader) +
                    "; after header=" + DescribeHeader(alignment.AfterHeader) +
                    "; row ambiguity=" + (alignment.RowAlignment?.Ambiguities.Count.ToString() ?? "n/a") +
                    "; column ambiguity=" + (alignment.ColumnAlignment?.Ambiguities.Count.ToString() ?? "n/a"));
            }
            _output.WriteLine("Findings:" + Environment.NewLine + findings);
            _output.WriteLine("Transitions: " +
                              (result.History?.AtomicTransitionCount.ToString() ?? "none"));

            Assert.True(result.IsComplete, findings);
            Assert.NotNull(result.History);
            Assert.All(result.Edges, edge => Assert.True(edge.IsComplete));

            ReviewWorkbookState review = OpenXmlProjectReviewProjector.Project(
                result,
                "{\"smoke\":\"real-workbook-projection\"}",
                new[]
                {
                    new ReviewRevisionDisplayMetadata(
                        "before", "R1.0", 0, "", beforePath, beforeHash, beforeWriteTime),
                    new ReviewRevisionDisplayMetadata(
                        "after", "R1.1", 1, "before", afterPath, afterHash, afterWriteTime),
                });
            _output.WriteLine("Projected Review rows: " + review.ChangeRows.Count);
            Assert.Equal(2, review.RevisionRows.Count);
            Assert.NotEmpty(review.ChangeRows);
            Assert.True(review.ChangeRows.Count <= ReviewWorkbookFile.MaxProjectedChangeRows);

            Assert.Equal(beforeWriteTime, File.GetLastWriteTimeUtc(beforePath));
            Assert.Equal(afterWriteTime, File.GetLastWriteTimeUtc(afterPath));
            Assert.Equal(beforeHash, Hash(File.ReadAllBytes(beforePath)));
            Assert.Equal(afterHash, Hash(File.ReadAllBytes(afterPath)));
        }

        private void ReportBounds(string label, OpenXmlWorkbookCompactionResult workbook)
        {
            _output.WriteLine(
                label + ": stored=" +
                workbook.Worksheets.Sum(item => item.Range.TotalStoredCellCount) +
                ", retained=" +
                workbook.Worksheets.Sum(item => item.Range.RetainedStoredCellCount) +
                ", grid=" +
                workbook.Worksheets.Sum(item => item.Range.GridCellCount));
        }

        private static string DescribeHeader(
            ExcelPowerTools.Core.Comparison.Alignment.HeaderDetectionResult header) =>
            header == null
                ? "n/a"
                : header.Decision + " accepted=" +
                  (header.AcceptedHeaderRowIndex?.ToString() ?? "none") +
                  " data-start=" + header.DataStartRowIndex;

        private static string Hash(byte[] bytes) =>
            Convert.ToHexString(SHA256.HashData(bytes));
    }

    public sealed class RealWorkbookComparisonFactAttribute : FactAttribute
    {
        public RealWorkbookComparisonFactAttribute()
        {
            if (string.IsNullOrWhiteSpace(
                    Environment.GetEnvironmentVariable(RealWorkbookComparisonSmokeTests.BeforePathVariable)) ||
                string.IsNullOrWhiteSpace(
                    Environment.GetEnvironmentVariable(RealWorkbookComparisonSmokeTests.AfterPathVariable)))
            {
                Skip = "Set both real-comparison workbook path variables to run this read-only smoke test.";
            }
        }
    }
}
