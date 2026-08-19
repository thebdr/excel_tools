using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Projection;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class OpenXmlWorkbookSheetMappingInferenceTests
    {
        [Fact]
        public void Builder_derives_deterministic_bounds_counts_and_conservative_evidence()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet("Data", 7, "DataCode",
                    InlineString("A1", "Id"),
                    InlineString("B1", "Amount"),
                    InlineString("A2", "A-1"),
                    Number("B2", "12"),
                    Formula("D5", "B2*2", "24")),
            });
            OpenXmlWorkbookSnapshot first = OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-a", bytes);
            OpenXmlWorkbookSnapshot second = OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-b", bytes);

            OpenXmlSheetFingerprintEvidence firstEvidence =
                Assert.Single(OpenXmlSheetFingerprintBuilder.Build(first).Sheets);
            OpenXmlSheetFingerprintEvidence secondEvidence =
                Assert.Single(OpenXmlSheetFingerprintBuilder.Build(second).Sheets);
            WorkbookSheetFingerprint fingerprint = firstEvidence.Fingerprint;

            Assert.True(firstEvidence.CanUseCanonicalCellEvidence);
            Assert.Empty(firstEvidence.CoverageCodes);
            Assert.Equal("Data", fingerprint.Name);
            Assert.Equal("DataCode", fingerprint.CodeName);
            Assert.Equal(WorkbookSheetKind.Worksheet, fingerprint.Kind);
            Assert.Equal(5, fingerprint.UsedRowCount);
            Assert.Equal(4, fingerprint.UsedColumnCount);
            Assert.Equal(5, fingerprint.StoredCellCount);
            Assert.True(StableId.IsSha256(fingerprint.StructureFingerprint));
            Assert.True(StableId.IsSha256(fingerprint.ContentFingerprint));
            Assert.True(StableId.IsSha256(fingerprint.HeaderFingerprint));
            Assert.True(StableId.IsSha256(fingerprint.FormulaFingerprint));

            Assert.NotEqual(fingerprint.SheetKey, secondEvidence.Fingerprint.SheetKey);
            Assert.Equal(fingerprint.StructureFingerprint, secondEvidence.Fingerprint.StructureFingerprint);
            Assert.Equal(fingerprint.ContentFingerprint, secondEvidence.Fingerprint.ContentFingerprint);
            Assert.Equal(fingerprint.HeaderFingerprint, secondEvidence.Fingerprint.HeaderFingerprint);
            Assert.Equal(fingerprint.FormulaFingerprint, secondEvidence.Fingerprint.FormulaFingerprint);
        }

        [Fact]
        public void Builder_ignores_a_physically_stored_blank_tail_for_bounds_and_identity()
        {
            OpenXmlWorkbookSnapshot original = Snapshot(
                "original",
                Sheet("Data", 1, null,
                    InlineString("A1", "Id"),
                    InlineString("B1", "Amount"),
                    InlineString("A2", "A-1"),
                    Number("B2", "12")));
            OpenXmlWorkbookSnapshot withBlankTail = Snapshot(
                "with-tail",
                Sheet("Data", 1, null,
                    InlineString("A1", "Id"),
                    InlineString("B1", "Amount"),
                    InlineString("A2", "A-1"),
                    Number("B2", "12"),
                    Blank("XFD1048576")));

            OpenXmlSheetFingerprintEvidence originalEvidence = Assert.Single(
                OpenXmlSheetFingerprintBuilder.Build(original).Sheets);
            OpenXmlSheetFingerprintEvidence tailEvidence = Assert.Single(
                OpenXmlSheetFingerprintBuilder.Build(withBlankTail).Sheets);
            CanonicalUsefulRange tailRange = CanonicalUsefulRangeAnalyzer.Analyze(
                Assert.Single(withBlankTail.Sheets).Snapshot);

            Assert.Equal(2, tailEvidence.Fingerprint.UsedRowCount);
            Assert.Equal(2, tailEvidence.Fingerprint.UsedColumnCount);
            Assert.Equal(4, tailEvidence.Fingerprint.StoredCellCount);
            Assert.Equal(originalEvidence.Fingerprint.StructureFingerprint,
                tailEvidence.Fingerprint.StructureFingerprint);
            Assert.Equal(originalEvidence.Fingerprint.ContentFingerprint,
                tailEvidence.Fingerprint.ContentFingerprint);
            Assert.Equal(originalEvidence.Fingerprint.HeaderFingerprint,
                tailEvidence.Fingerprint.HeaderFingerprint);
            Assert.Equal(5, tailRange.TotalStoredCellCount);
            Assert.Equal(1, tailRange.IgnoredOutsideRangeCellCount);
            Assert.Equal("A1:B2", tailRange.DisplayAddress);
        }

        [Fact]
        public void Cached_result_is_not_primary_identity_but_formula_text_is()
        {
            OpenXmlSheetFingerprintEvidence original = Evidence(
                "original",
                Sheet("Data", 1, null, Formula("A1", "1+1", "2")));
            OpenXmlSheetFingerprintEvidence cacheChanged = Evidence(
                "cache-changed",
                Sheet("Data", 1, null, Formula("A1", "1+1", "999")));
            OpenXmlSheetFingerprintEvidence formulaChanged = Evidence(
                "formula-changed",
                Sheet("Data", 1, null, Formula("A1", "1+2", "3")));

            Assert.Equal(original.Fingerprint.StructureFingerprint, cacheChanged.Fingerprint.StructureFingerprint);
            Assert.Equal(original.Fingerprint.ContentFingerprint, cacheChanged.Fingerprint.ContentFingerprint);
            Assert.Equal(original.Fingerprint.FormulaFingerprint, cacheChanged.Fingerprint.FormulaFingerprint);

            Assert.Equal(original.Fingerprint.StructureFingerprint, formulaChanged.Fingerprint.StructureFingerprint);
            Assert.NotEqual(original.Fingerprint.ContentFingerprint, formulaChanged.Fingerprint.ContentFingerprint);
            Assert.NotEqual(original.Fingerprint.FormulaFingerprint, formulaChanged.Fingerprint.FormulaFingerprint);
        }

        [Fact]
        public void Header_evidence_requires_a_strong_text_header_and_a_type_transition()
        {
            OpenXmlSheetFingerprintEvidence strong = Evidence(
                "strong",
                Sheet("Data", 1, null,
                    InlineString("A1", "Id"),
                    InlineString("B1", "Amount"),
                    InlineString("A2", "A-1"),
                    Number("B2", "12")));
            OpenXmlSheetFingerprintEvidence allText = Evidence(
                "all-text",
                Sheet("Data", 1, null,
                    InlineString("A1", "Id"),
                    InlineString("B1", "Amount"),
                    InlineString("A2", "A-1"),
                    InlineString("B2", "twelve")));

            Assert.NotNull(strong.Fingerprint.HeaderFingerprint);
            Assert.Null(allText.Fingerprint.HeaderFingerprint);
            Assert.NotNull(allText.Fingerprint.StructureFingerprint);
            Assert.NotNull(allText.Fingerprint.ContentFingerprint);
        }

        [Fact]
        public void Equality_affecting_coverage_suppresses_cell_fingerprints_not_observed_bounds()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, "DataCode", SharedString("C4", 99)) },
                sharedStrings: new[] { "only item" });
            OpenXmlWorkbookSnapshot snapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes("revision", bytes);

            OpenXmlSheetFingerprintEvidence evidence =
                Assert.Single(OpenXmlSheetFingerprintBuilder.Build(snapshot).Sheets);

            Assert.False(evidence.CanUseCanonicalCellEvidence);
            Assert.Contains(OpenXmlCoverageCodes.InvalidSharedStringIndex, evidence.CoverageCodes);
            Assert.Equal(4, evidence.Fingerprint.UsedRowCount);
            Assert.Equal(3, evidence.Fingerprint.UsedColumnCount);
            Assert.Equal(1, evidence.Fingerprint.StoredCellCount);
            Assert.Equal("DataCode", evidence.Fingerprint.CodeName);
            Assert.Null(evidence.Fingerprint.StructureFingerprint);
            Assert.Null(evidence.Fingerprint.ContentFingerprint);
            Assert.Null(evidence.Fingerprint.HeaderFingerprint);
            Assert.Null(evidence.Fingerprint.FormulaFingerprint);
        }

        [Fact]
        public void Non_cell_coverage_is_retained_without_discarding_valid_cell_evidence()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, Number("A1", "1")) },
                macroEnabled: true);
            OpenXmlWorkbookSnapshot snapshot = OpenXmlWorkbookSnapshotAdapter.FromBytes("revision", bytes);

            OpenXmlWorkbookSheetFingerprintSet set = OpenXmlSheetFingerprintBuilder.Build(snapshot);
            OpenXmlSheetFingerprintEvidence evidence = Assert.Single(set.Sheets);

            Assert.Contains(OpenXmlCoverageCodes.VbaProjectNotInspected, set.WorkbookCoverageCodes);
            Assert.Contains(OpenXmlCoverageCodes.VbaProjectNotInspected, evidence.CoverageCodes);
            Assert.True(evidence.CanUseCanonicalCellEvidence);
            Assert.NotNull(evidence.Fingerprint.StructureFingerprint);
            Assert.NotNull(evidence.Fingerprint.ContentFingerprint);
        }

        [Fact]
        public void Snapshot_helper_infers_a_supported_rename_while_exact_engine_remains_exact_name_only()
        {
            OpenXmlWorkbookSnapshot before = Snapshot(
                "before",
                Sheet("Old Name", 1, "StableCode", Number("A1", "7")));
            OpenXmlWorkbookSnapshot after = Snapshot(
                "after",
                Sheet("New Name", 9, "StableCode", Number("A1", "7")));

            OpenXmlWorkbookSheetMappingResult inferred =
                OpenXmlWorkbookSheetMappingInference.Infer(before, after);
            ResolvedSheetMapping mapping = Assert.Single(inferred.Mapping.Mappings);
            Assert.Equal(SheetMappingKind.InferredRename, mapping.Kind);
            Assert.Equal("Old Name", mapping.Before.Name);
            Assert.Equal("New Name", mapping.After.Name);
            Assert.Contains(SheetMappingReasonCodes.StructureFingerprint, mapping.ReasonCodes);
            Assert.Contains(SheetMappingReasonCodes.ContentFingerprint, mapping.ReasonCodes);
            Assert.Contains(SheetMappingReasonCodes.CodeName, mapping.ReasonCodes);

            ExactWorkbookComparisonResult exact = ExactWorkbookComparisonEngine.Compare(before, after);
            Assert.DoesNotContain(exact.Sheets, sheet => sheet.MatchKind == ExactSheetMatchKind.ExactName);
            Assert.Contains(exact.Sheets, sheet => sheet.MatchKind == ExactSheetMatchKind.OnlyBefore);
            Assert.Contains(exact.Sheets, sheet => sheet.MatchKind == ExactSheetMatchKind.OnlyAfter);
        }

        [Fact]
        public void Snapshot_helper_keeps_equal_rename_candidates_unresolved()
        {
            OpenXmlWorkbookSnapshot before = Snapshot(
                "before",
                Sheet("Original", 1, null, Number("A1", "7")));
            OpenXmlWorkbookSnapshot after = Snapshot(
                "after",
                Sheet("Candidate B", 2, null, Number("A1", "7")),
                Sheet("Candidate A", 3, null, Number("A1", "7")));

            WorkbookSheetMappingResult mapping =
                OpenXmlWorkbookSheetMappingInference.InferMappings(before, after);

            Assert.Empty(mapping.Mappings);
            UnresolvedSheetMapping unresolved = Assert.Single(mapping.UnresolvedBeforeSheets);
            Assert.Equal(2, unresolved.Candidates.Count);
            Assert.Contains(SheetMappingReasonCodes.CompetingTargets, unresolved.ReasonCodes);
        }

        [Fact]
        public void Coverage_placeholders_never_become_rename_evidence()
        {
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(
                    new[] { Sheet("Old", 1, "StableCode", SharedString("A1", 99)) },
                    sharedStrings: new[] { "only item" }));
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(
                    new[] { Sheet("New", 2, "StableCode", SharedString("A1", 99)) },
                    sharedStrings: new[] { "only item" }));

            OpenXmlWorkbookSheetMappingResult result =
                OpenXmlWorkbookSheetMappingInference.Infer(before, after);

            Assert.Empty(result.Mapping.Mappings);
            UnresolvedSheetMapping unresolved = Assert.Single(result.Mapping.UnresolvedBeforeSheets);
            Assert.Empty(unresolved.Candidates);
            Assert.Contains(SheetMappingReasonCodes.NoSupportedRenameEvidence, unresolved.ReasonCodes);
            Assert.False(Assert.Single(result.Before.Sheets).CanUseCanonicalCellEvidence);
            Assert.False(Assert.Single(result.After.Sheets).CanUseCanonicalCellEvidence);
        }

        private static OpenXmlSheetFingerprintEvidence Evidence(
            string revision,
            OpenXmlWorkbookFixture.SheetDefinition sheet) =>
            Assert.Single(OpenXmlSheetFingerprintBuilder.Build(Snapshot(revision, sheet)).Sheets);

        private static OpenXmlWorkbookSnapshot Snapshot(
            string revision,
            params OpenXmlWorkbookFixture.SheetDefinition[] sheets) =>
            OpenXmlWorkbookSnapshotAdapter.FromBytes(
                revision,
                OpenXmlWorkbookFixture.Create(sheets));

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            uint sheetId,
            string? codeName,
            params Cell[] cells) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                SheetId = sheetId,
                CodeName = codeName,
                Cells = cells,
            };

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell SharedString(string reference, int index) =>
            new Cell(new CellValue(index.ToString()))
            {
                CellReference = reference,
                DataType = CellValues.SharedString,
            };

        private static Cell InlineString(string reference, string value) =>
            new Cell(new InlineString(new SpreadsheetText(value)))
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
            };

        private static Cell Formula(string reference, string formula, string cached) =>
            new Cell(new CellFormula(formula), new CellValue(cached))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell Blank(string reference) =>
            new Cell
            {
                CellReference = reference,
                StyleIndex = 0U,
            };
    }
}
