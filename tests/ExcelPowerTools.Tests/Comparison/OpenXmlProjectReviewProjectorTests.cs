using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Core.Comparison.Project;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class OpenXmlProjectReviewProjectorTests
    {
        [Fact]
        public void Complete_project_uses_master_physical_location_and_minimum_axis_confidence()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Original data", "StableDataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20")));
            OpenXmlWorkbookSnapshot inserted = Workbook("inserted", Sheet("Original data", "StableDataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "20")));
            OpenXmlWorkbookSnapshot renamed = Workbook("renamed", Sheet("Reviewed data", "StableDataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "20")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Reviewed data", "StableDataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "25")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-review-complete",
                "mapping-generation-7",
                new[]
                {
                    new OpenXmlProjectRevision(
                        "before",
                        "",
                        0,
                        before,
                        isTrustedBaseline: true),
                    new OpenXmlProjectRevision("inserted", "before", 1, inserted),
                    new OpenXmlProjectRevision("renamed", "inserted", 2, renamed),
                    new OpenXmlProjectRevision("master", "renamed", 3, master, isMaster: true),
                });
            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);

            ReviewWorkbookState state = OpenXmlProjectReviewProjector.Project(
                result,
                "{\"schema\":\"project-review-v1\"}",
                new[]
                {
                    Revision("master", "Reviewed", 3, "renamed", @"C:\\private\\Reviewed.xlsx"),
                    Revision("renamed", "Renamed", 2, "inserted", @"C:\\private\\Renamed.xlsx"),
                    Revision("inserted", "Expanded", 1, "before", @"C:\\private\\Expanded.xlsx"),
                    Revision("before", "Received", 0, "", "/srv/private/Received.xlsx"),
                });

            ReviewChangeProjectionRow amount = Assert.Single(state.ChangeRows, item =>
                item.Aspect == "Value" && item.BeforeText == "20" && item.AfterText == "25");
            Assert.Equal("Reviewed data", amount.SheetName);
            Assert.Equal("B4", amount.Address);

            OpenXmlAlignedSheetRevision masterSheet = result.Revisions
                .Single(item => item.Source.RevisionId == "master")
                .Sheets.Single();
            AlignmentConfidence expectedConfidence = Minimum(
                masterSheet.Rows.AssignmentsByPhysicalIndex[3].Confidence,
                masterSheet.Columns.AssignmentsByPhysicalIndex[1].Confidence);
            Assert.Equal(expectedConfidence.ToString(), amount.AlignmentConfidence);
            ReviewChangeProjectionRow insertedRow = Assert.Single(state.ChangeRows, item =>
                item.Aspect == "Row structure" && item.Flags == "Structural; Insertion");
            Assert.Equal("Row 3", insertedRow.Address);
            Assert.Equal("Inserted Row 3", insertedRow.AfterText);
            ReviewChangeProjectionRow renamedSheet = Assert.Single(state.ChangeRows, item =>
                item.Aspect == "Sheet name" && item.Flags == "Structural; Rename");
            Assert.Equal("Original data", renamedSheet.BeforeText);
            Assert.Equal("Reviewed data", renamedSheet.AfterText);

            Assert.Equal(new[] { "Received.xlsx", "Expanded.xlsx", "Renamed.xlsx", "Reviewed.xlsx" },
                state.RevisionRows.Select(item => item.SourceBaseName));
            Assert.Equal(new[] { "Trusted baseline", "Revision", "Revision", "Master" },
                state.RevisionRows.Select(item => item.Role));
            Assert.Equal(3, state.DiagnosticRows.Count(item =>
                item.Code == OpenXmlProjectFindingCodes.NonFormulaValueContentUnmodeled &&
                item.Severity == "Warning"));
            Assert.Equal("{\"schema\":\"project-review-v1\"}", state.StateJson);
            Assert.Contains(
                @"C:\\private\\Reviewed.xlsx",
                PortableProjectionText(state),
                StringComparison.OrdinalIgnoreCase);
            Assert.Contains("/srv/private/Received.xlsx", PortableProjectionText(state), StringComparison.OrdinalIgnoreCase);
            Assert.StartsWith("Complete comparison", state.ComparisonStatus, StringComparison.Ordinal);
        }

        [Fact]
        public void Blocked_project_still_projects_revision_roles_and_every_verbose_finding()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Original", "StableCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "one")));
            OpenXmlWorkbookSnapshot after = Workbook("after",
                Sheet("Candidate A", "StableCode",
                    Text("A1", "ID"), Text("B1", "Value"),
                    Text("A2", "A"), Text("B2", "one")),
                Sheet("Candidate B", "StableCode",
                    Text("A1", "ID"), Text("B1", "Value"),
                    Text("A2", "A"), Text("B2", "one")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-review-blocked",
                "mapping-generation-8",
                new[]
                {
                    new OpenXmlProjectRevision(
                        "before",
                        "",
                        0,
                        before,
                        isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });
            Assert.False(result.IsComplete);
            Assert.Null(result.History);

            ReviewWorkbookState state = OpenXmlProjectReviewProjector.Project(
                result,
                "{\"status\":\"blocked\"}",
                new[]
                {
                    Revision("after", "Latest", 1, "before", @"C:\\people\\Latest.xlsx"),
                    Revision("before", "Earlier", 0, "", @"D:\\mail\\Earlier.xlsx"),
                });

            Assert.Equal(2, state.ChangeRows.Count(item => item.Flags == "Structural; Insertion"));
            Assert.Equal(new[] { "Earlier.xlsx", "Latest.xlsx" },
                state.RevisionRows.Select(item => item.SourceBaseName));
            Assert.Equal(new[] { "Trusted baseline", "Master" },
                state.RevisionRows.Select(item => item.Role));
            Assert.Equal(result.Findings.Count, state.DiagnosticRows.Count);

            OpenXmlProjectFinding finding = Assert.Single(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.SheetMappingAmbiguous);
            ReviewDiagnosticProjectionRow diagnostic = Assert.Single(state.DiagnosticRows, item =>
                item.FindingUid == finding.FindingId);
            Assert.Equal(finding.Code, diagnostic.Code);
            Assert.Equal(finding.Severity.ToString(), diagnostic.Severity);
            Assert.Equal("before -> after", diagnostic.RevisionId);
            Assert.Equal(finding.Message, diagnostic.Summary);
            Assert.Equal(finding.SuggestedAction, diagnostic.SuggestedAction);
            Assert.Contains("Finding ID: " + finding.FindingId, diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Project ID: project-review-blocked", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Mapping generation ID: mapping-generation-8", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Before revision: before", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("After revision: after", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Logical sheet ID:", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Physical sheet:", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Reason codes:", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Candidate sheets: Candidate A, Candidate B", diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains("Suggested action: " + finding.SuggestedAction, diagnostic.Detail, StringComparison.Ordinal);
            Assert.Contains(
                "Engine source: " + OpenXmlProjectReviewProjector.EngineSourceHint,
                diagnostic.Detail,
                StringComparison.Ordinal);

            string projection = PortableProjectionText(state);
            Assert.Contains(@"C:\\people", projection, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(@"D:\\mail", projection, StringComparison.OrdinalIgnoreCase);
            Assert.StartsWith("Incomplete diagnostic comparison", state.ComparisonStatus, StringComparison.Ordinal);
        }

        [Fact]
        public void Explicit_deleted_new_resolution_projects_both_structural_changes_without_a_rename()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Retired", "RetiredCode",
                Text("A1", "Old ID"), Text("B1", "Old value"), Text("A2", "A"), Text("B2", "one")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Created", "CreatedCode",
                Text("D5", "New ID"), Text("E5", "New value"), Text("D6", "Z"), Text("E6", "nine")));
            string deletedKey = OpenXmlSheetFingerprintBuilder.Build(before).Sheets.Single().Fingerprint.SheetKey;
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-generation-deleted-new"),
                AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                    "before",
                    "after",
                    deletedKey,
                    new ReviewerSnapshot("Reviewer", "", new[] { "projector-test" }),
                    new DateTimeOffset(2026, 7, 28, 8, 0, 0, TimeSpan.Zero),
                    "Retired was deleted; Created is unrelated."));
            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-review-deleted-new",
                "mapping-generation-deleted-new",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                },
                new OpenXmlComparisonProjectOptions { AlignmentCorrections = corrections });
            Assert.True(result.IsComplete);

            ReviewWorkbookState state = OpenXmlProjectReviewProjector.Project(
                result,
                "{\"status\":\"complete\"}",
                new[]
                {
                    Revision("before", "Before", 0, "", "before.xlsx"),
                    Revision("after", "After", 1, "before", "after.xlsx"),
                });

            ReviewChangeProjectionRow deletion = Assert.Single(state.ChangeRows, item =>
                item.Flags == "Structural; Deletion");
            Assert.Equal("Retired", deletion.SheetName);
            Assert.Equal("Deleted sheet Retired", deletion.BeforeText);
            Assert.Empty(deletion.AfterText);
            ReviewChangeProjectionRow insertion = Assert.Single(state.ChangeRows, item =>
                item.Flags == "Structural; Insertion");
            Assert.Equal("Created", insertion.SheetName);
            Assert.Equal("Added sheet Created", insertion.AfterText);
            Assert.DoesNotContain(state.ChangeRows, item => item.Flags == "Structural; Rename");
        }

        private static ReviewRevisionDisplayMetadata Revision(
            string revisionId,
            string label,
            int order,
            string parentRevisionId,
            string sourceFileName) =>
            new ReviewRevisionDisplayMetadata(
                revisionId,
                label,
                order,
                parentRevisionId,
                sourceFileName,
                "sha256:" + revisionId,
                new DateTimeOffset(2026, 7, 17, 10 + order, 0, 0, TimeSpan.Zero),
                "User-confirmed revision tree.",
                "Confirmed");

        private static string PortableProjectionText(ReviewWorkbookState state) =>
            string.Join("\n", state.RevisionRows.Select(item =>
                    item.RevisionId + "|" + item.Label + "|" + item.ParentRevisionId + "|" + item.Role + "|" +
                    item.SourceBaseName + "|" + item.SourcePath + "|" + item.ByteContentId + "|" + item.ModifiedUtc + "|" +
                    item.Inference + "|" + item.Confidence)
                .Concat(state.ChangeRows.Select(item =>
                    item.ChangeId + "|" + item.SheetName + "|" + item.Address + "|" + item.Aspect + "|" +
                    item.BeforeText + "|" + item.AfterText + "|" + item.GroupId + "|" + item.Flags))
                .Concat(state.DiagnosticRows.Select(item =>
                    item.FindingUid + "|" + item.Code + "|" + item.RevisionId + "|" + item.Summary + "|" +
                    item.Detail + "|" + item.SuggestedAction)));

        private static AlignmentConfidence Minimum(
            AlignmentConfidence left,
            AlignmentConfidence right) =>
            left <= right ? left : right;

        private static OpenXmlWorkbookSnapshot Workbook(
            string revisionId,
            params OpenXmlWorkbookFixture.SheetDefinition[] sheets) =>
            OpenXmlWorkbookSnapshotAdapter.FromBytes(revisionId, OpenXmlWorkbookFixture.Create(sheets));

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            string codeName,
            params Cell[] cells) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                CodeName = codeName,
                Cells = cells,
            };

        private static Cell Text(string reference, string value) =>
            new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new SpreadsheetText(value)),
            };
    }
}
