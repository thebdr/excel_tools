using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Project;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class AlignmentCorrectionReconcilerTests
    {
        private static readonly ReviewerSnapshot Reviewer = new ReviewerSnapshot(
            "Reviewer",
            "",
            new[] { "reconciliation-test" });
        private static readonly DateTimeOffset When = new DateTimeOffset(
            2026, 7, 17, 10, 11, 12, TimeSpan.Zero);

        [Fact]
        public void Remap_clears_displaced_key_but_preserves_physical_sheet_headers_before_engine_build()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                Sheet("Alpha", "AlphaCode", "A", "one"),
                Sheet("Beta", "BetaCode", "B", "two"));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                Sheet("Alpha", "AlphaCode", "A", "one changed"),
                Sheet("Beta", "BetaCode", "B", "two changed"));
            string beforeAlpha = SheetKey(before, "Alpha");
            string beforeBeta = SheetKey(before, "Beta");
            string afterAlpha = SheetKey(after, "Alpha");
            string afterBeta = SheetKey(after, "Beta");

            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmHeaderRow(
                "before", beforeAlpha, 0, 1, Reviewer, When, "before Alpha header"));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmHeaderRow(
                "after", afterAlpha, 0, 1, Reviewer, When, "after Alpha header"));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmHeaderRow(
                "after", afterBeta, 0, 1, Reviewer, When, "after Beta header"));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmKeyColumns(
                "before",
                "after",
                beforeAlpha,
                afterAlpha,
                new[] { new AlignmentGridIndexReference(0, 1) },
                new[] { new AlignmentGridIndexReference(0, 1) },
                Reviewer,
                When,
                "Alpha key"));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmSheetPair(
                "before",
                "after",
                beforeBeta,
                afterAlpha,
                Reviewer,
                When,
                "Beta is the actual source of after Alpha"));

            Assert.Throws<InvalidOperationException>(() => OpenXmlComparisonProjectEngine.Build(
                "project-remap-unreconciled",
                "mapping-1",
                Revisions(before, after),
                new OpenXmlComparisonProjectOptions
                {
                    AlignmentCorrections = corrections,
                }));

            AlignmentCorrectionReconciliationResult reconciled = AlignmentCorrectionReconciler.Reconcile(
                Revisions(before, after),
                corrections,
                WorkbookSheetMappingOptions.Default,
                Reviewer,
                When);

            AlignmentCorrectionReconciliationAction action = Assert.Single(reconciled.Actions);
            Assert.Equal(AlignmentCorrectionReconciliationCodes.KeyPairNotAccepted, action.ReasonCode);
            Assert.Equal(AlignmentCorrectionKind.ClearKeyColumns, action.Kind);
            Assert.Empty(reconciled.Generation.KeyColumns);
            Assert.Equal(3, reconciled.Generation.HeaderRows.Count);
            Assert.Contains(reconciled.Generation.HeaderRows, item =>
                item.RevisionId == "before" && item.PhysicalSheetKey == beforeAlpha);
            Assert.Contains(reconciled.Generation.HeaderRows, item =>
                item.RevisionId == "after" && item.PhysicalSheetKey == afterAlpha);
            Assert.Contains(reconciled.Generation.HeaderRows, item =>
                item.RevisionId == "after" && item.PhysicalSheetKey == afterBeta);
            Assert.All(reconciled.Actions, item =>
            {
                Assert.NotEmpty(item.EventId);
                Assert.NotEmpty(item.RequestId);
                Assert.NotEmpty(item.SupersededConfirmationId);
                Assert.Contains(item.ReasonCode, item.CorrectionEvent.Note, StringComparison.Ordinal);
            });

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-remap-reconciled",
                "mapping-1",
                Revisions(before, after),
                new OpenXmlComparisonProjectOptions
                {
                    AlignmentCorrections = reconciled.Generation,
                });
            Assert.Equal(reconciled.Generation.GenerationId, result.AlignmentCorrectionGenerationId);
        }

        [Fact]
        public void Complete_swap_retains_headers_on_still_consumed_sheets_but_clears_old_pair_key()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                Sheet("Alpha", "AlphaCode", "A", "one"),
                Sheet("Beta", "BetaCode", "B", "two"));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                Sheet("Alpha", "AlphaCode", "A", "one changed"),
                Sheet("Beta", "BetaCode", "B", "two changed"));
            string beforeAlpha = SheetKey(before, "Alpha");
            string beforeBeta = SheetKey(before, "Beta");
            string afterAlpha = SheetKey(after, "Alpha");
            string afterBeta = SheetKey(after, "Beta");

            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmHeaderRow(
                "before", beforeAlpha, 0, 1, Reviewer, When));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmHeaderRow(
                "after", afterAlpha, 0, 1, Reviewer, When));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmKeyColumns(
                "before",
                "after",
                beforeAlpha,
                afterAlpha,
                new[] { new AlignmentGridIndexReference(0, 1) },
                new[] { new AlignmentGridIndexReference(0, 1) },
                Reviewer,
                When));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", beforeAlpha, afterBeta, Reviewer, When));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", beforeBeta, afterAlpha, Reviewer, When));

            AlignmentCorrectionReconciliationResult result = AlignmentCorrectionReconciler.Reconcile(
                Revisions(before, after),
                corrections,
                WorkbookSheetMappingOptions.Default,
                Reviewer,
                When);

            AlignmentCorrectionReconciliationAction action = Assert.Single(result.Actions);
            Assert.Equal(AlignmentCorrectionReconciliationCodes.KeyPairNotAccepted, action.ReasonCode);
            Assert.Equal(AlignmentCorrectionKind.ClearKeyColumns, action.Kind);
            Assert.Empty(result.Generation.KeyColumns);
            Assert.Equal(2, result.Generation.HeaderRows.Count);
            Assert.Contains(result.Generation.HeaderRows, item =>
                item.RevisionId == "before" && item.PhysicalSheetKey == beforeAlpha);
            Assert.Contains(result.Generation.HeaderRows, item =>
                item.RevisionId == "after" && item.PhysicalSheetKey == afterAlpha);
        }

        [Fact]
        public void Generation_without_active_sheet_remap_is_an_exact_no_op()
        {
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            AlignmentCorrectionReconciliationResult result = AlignmentCorrectionReconciler.Reconcile(
                Array.Empty<OpenXmlProjectRevision>(),
                corrections,
                WorkbookSheetMappingOptions.Default,
                Reviewer,
                When);

            Assert.Same(corrections, result.Generation);
            Assert.Empty(result.Actions);
        }

        [Fact]
        public void Deleted_new_resolution_clears_a_key_bound_to_the_rejected_pair()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                Sheet("Data", "DataCode", "A", "one"));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                Sheet("Data", "DataCode", "A", "two"));
            string beforeKey = SheetKey(before, "Data");
            string afterKey = SheetKey(after, "Data");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmKeyColumns(
                "before",
                "after",
                beforeKey,
                afterKey,
                new[] { new AlignmentGridIndexReference(0, 1) },
                new[] { new AlignmentGridIndexReference(0, 1) },
                Reviewer,
                When));
            corrections = Apply(corrections, AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                "before", "after", beforeKey, Reviewer, When, "The same name was reused for a new sheet."));

            AlignmentCorrectionReconciliationResult result = AlignmentCorrectionReconciler.Reconcile(
                Revisions(before, after),
                corrections,
                WorkbookSheetMappingOptions.Default,
                Reviewer,
                When);

            Assert.Single(result.Generation.UnmatchedSheets);
            Assert.Empty(result.Generation.KeyColumns);
            AlignmentCorrectionReconciliationAction action = Assert.Single(result.Actions);
            Assert.Equal(AlignmentCorrectionKind.ClearKeyColumns, action.Kind);
            Assert.Equal(AlignmentCorrectionReconciliationCodes.KeyPairNotAccepted, action.ReasonCode);
        }

        private static AlignmentCorrectionGeneration Apply(
            AlignmentCorrectionGeneration generation,
            AlignmentCorrectionRequest request) =>
            AlignmentCorrectionEngine.Apply(generation, request);

        private static OpenXmlProjectRevision[] Revisions(
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after) =>
            new[]
            {
                new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
            };

        private static OpenXmlWorkbookSnapshot Workbook(
            string revisionId,
            params OpenXmlWorkbookFixture.SheetDefinition[] sheets) =>
            OpenXmlWorkbookSnapshotAdapter.FromBytes(revisionId, OpenXmlWorkbookFixture.Create(sheets));

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            string codeName,
            string id,
            string value) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                CodeName = codeName,
                Cells = new[]
                {
                    Text("A1", "ID"),
                    Text("B1", "Value"),
                    Text("A2", id),
                    Text("B2", value),
                },
            };

        private static string SheetKey(OpenXmlWorkbookSnapshot workbook, string name) =>
            OpenXmlSheetFingerprintBuilder.Build(workbook).Sheets.Single(item =>
                item.Fingerprint.Name == name).Fingerprint.SheetKey;

        private static Cell Text(string reference, string value) =>
            new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new SpreadsheetText(value)),
            };
    }
}
