using System;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Diff;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ExactWorkbookComparisonEngineTests
    {
        [Fact]
        public void Produces_deterministic_atomic_changes_for_exact_sheet_name_matches()
        {
            OpenXmlWorkbookSnapshot before = Snapshot(
                "before",
                Sheet("Zeta", 2, Number("B2", "10")),
                Sheet("Alpha", 1, Number("A1", "1")));
            OpenXmlWorkbookSnapshot after = Snapshot(
                "after",
                Sheet("Alpha", 7, Number("A1", "2")),
                Sheet("Zeta", 8, Number("B2", "10")));

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(before, after);

            Assert.Equal(ExactComparisonStatus.Different, result.Status);
            Assert.Equal(new[] { "Alpha", "Zeta" }, result.Sheets.Select(sheet => sheet.SheetName));
            ExactSheetComparison changed = result.Sheets[0];
            AtomicCellChange change = Assert.Single(changed.Changes);
            Assert.Equal("A1", change.SubjectKey);
            Assert.Equal(CellChangeAspect.Value, change.Aspect);
            Assert.Equal("1", change.BeforeValue!.ExactText);
            Assert.Equal("2", change.AfterValue!.ExactText);
            Assert.Equal(1, result.AtomicChangeCount);
            Assert.False(result.CanAssertExactCellEquality);
        }

        [Fact]
        public void Renamed_sheets_are_reported_as_unmatched_and_never_guessed_from_code_name()
        {
            var beforeSheet = Sheet("Old Name", 1, Number("A1", "7"));
            beforeSheet.CodeName = "StableCode";
            var afterSheet = Sheet("New Name", 2, Number("A1", "7"));
            afterSheet.CodeName = "StableCode";

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(
                Snapshot("before", beforeSheet),
                Snapshot("after", afterSheet));

            Assert.Equal(ExactComparisonStatus.Different, result.Status);
            Assert.Collection(
                result.Sheets,
                sheet => Assert.Equal(ExactSheetMatchKind.OnlyAfter, sheet.MatchKind),
                sheet => Assert.Equal(ExactSheetMatchKind.OnlyBefore, sheet.MatchKind));
            Assert.Equal(0, result.AtomicChangeCount);
        }

        [Fact]
        public void Exact_byte_identity_proves_cell_equality_even_when_an_aspect_is_not_inspected()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, Number("A1", "4")) },
                macroEnabled: true);
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes("before", bytes);
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes("after", bytes);
            Assert.False(before.CanAssertExactEquality);

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(before, after);

            Assert.True(result.ByteIdentical);
            Assert.Equal(ExactComparisonStatus.Equal, result.Status);
            Assert.Equal(ExactComparisonStatus.Equal, Assert.Single(result.Sheets).Status);
            Assert.True(result.CanAssertExactCellEquality);
        }

        [Fact]
        public void Coverage_without_an_observed_difference_is_indeterminate_for_different_packages()
        {
            OpenXmlWorkbookFixture.SheetDefinition beforeSheet = Sheet("Data", 1);
            beforeSheet.OmitSheetData = true;
            OpenXmlWorkbookFixture.SheetDefinition afterSheet = Sheet("Data", 9);
            afterSheet.OmitSheetData = true;

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(
                Snapshot("before", beforeSheet),
                Snapshot("after", afterSheet));

            Assert.False(result.ByteIdentical);
            Assert.Equal(ExactComparisonStatus.Indeterminate, result.Status);
            Assert.Equal(ExactComparisonStatus.Indeterminate, Assert.Single(result.Sheets).Status);
            Assert.Equal(0, result.AtomicChangeCount);
        }

        [Fact]
        public void Visibility_is_retained_as_secondary_evidence_without_changing_cell_equality()
        {
            OpenXmlWorkbookFixture.SheetDefinition beforeSheet = Sheet("Data", 1, Number("A1", "4"));
            OpenXmlWorkbookFixture.SheetDefinition afterSheet = Sheet("Data", 1, Number("A1", "4"));
            afterSheet.State = SheetStateValues.Hidden;

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(
                Snapshot("before", beforeSheet),
                Snapshot("after", afterSheet));

            Assert.Equal(ExactComparisonStatus.Equal, result.Status);
            ExactSheetComparison sheet = Assert.Single(result.Sheets);
            Assert.True(sheet.VisibilityChanged);
            Assert.Empty(sheet.Changes);
        }

        [Fact]
        public void Honors_cancellation_before_and_during_sheet_iteration()
        {
            OpenXmlWorkbookSnapshot before = Snapshot("before", Sheet("Data", 1, Number("A1", "1")));
            OpenXmlWorkbookSnapshot after = Snapshot("after", Sheet("Data", 1, Number("A1", "2")));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() => ExactWorkbookComparisonEngine.Compare(
                before,
                after,
                cancellationToken: cancellation.Token));
        }

        private static OpenXmlWorkbookSnapshot Snapshot(
            string revisionId,
            params OpenXmlWorkbookFixture.SheetDefinition[] sheets) =>
            OpenXmlWorkbookSnapshotAdapter.FromBytes(
                revisionId,
                OpenXmlWorkbookFixture.Create(sheets));

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            uint sheetId,
            params Cell[] cells) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                SheetId = sheetId,
                Cells = cells,
            };

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };
    }
}
