using System.Collections.Generic;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Projection;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ComparisonAlignmentEnvelopePolicyTests
    {
        [Fact]
        public void Large_physical_blank_tail_is_accepted_when_the_effective_grid_is_small()
        {
            ComparisonAlignmentEnvelopeAssessment screenshotScale =
                ComparisonAlignmentEnvelopePolicy.Evaluate(
                    workbookStoredCellCount: 1152115L,
                    aggregateStoredCellCount: 2310363L,
                    aggregateRetainedStoredCellCount: 101273L,
                    aggregateGridCellCount: 109971L);
            Assert.True(screenshotScale.CanAnalyze);
            Assert.Equal(ComparisonAlignmentEnvelopeLimit.None, screenshotScale.ExceededLimit);
        }

        [Fact]
        public void Effective_grid_above_ten_million_cells_is_rejected()
        {
            ComparisonAlignmentEnvelopeAssessment excessiveGrid =
                ComparisonAlignmentEnvelopePolicy.Evaluate(
                    workbookStoredCellCount: 1152115L,
                    aggregateStoredCellCount: 2310363L,
                    aggregateRetainedStoredCellCount: 101273L,
                    aggregateGridCellCount: ComparisonAlignmentEnvelopePolicy.MaximumAggregateGridCells + 1L);
            Assert.False(excessiveGrid.CanAnalyze);
            Assert.Equal(
                ComparisonAlignmentEnvelopeLimit.AggregateGridCells,
                excessiveGrid.ExceededLimit);
        }

        [Fact]
        public void Compaction_drops_only_outside_blank_tails_and_preserves_audited_evidence()
        {
            CanonicalCellState firstContent = CanonicalCellState.Constant(
                "A1",
                CanonicalValue.Text("ID"));
            CanonicalCellState interiorBlank = CanonicalCellState.Blank("B2");
            CanonicalCellState lastContent = CanonicalCellState.Formula(
                "C3",
                "A1",
                CanonicalValue.Number(1D),
                "1",
                "formula-metadata");
            CanonicalCellState outsideBlank = CanonicalCellState.Blank("Z100");
            CanonicalCellState farOutsideBlank = CanonicalCellState.Blank("XFD1048576");
            CanonicalCellState invalidCoordinate = CanonicalCellState.Blank("not-a-cell");
            var formulaObservation = new OpenXmlFormulaObservation(
                "C3",
                OpenXmlFormulaForm.Normal,
                "A1",
                "A1",
                null,
                null,
                false,
                "formula-metadata",
                CanonicalValue.Number(1D));
            var tableObservation = new OpenXmlTableFormulaObservation(
                "Table1",
                "A1:C3",
                3U,
                "Amount",
                OpenXmlTableFormulaKind.CalculatedColumn,
                "[@ID]",
                "table-metadata",
                "column-metadata",
                "combined-metadata",
                "/xl/tables/table1.xml");
            var sheetCoverage = new OpenXmlCoverageMessage(
                OpenXmlCoverageCodes.InvalidCellReference,
                "An invalid cell coordinate was retained for diagnosis.",
                true,
                true,
                "Data",
                "not-a-cell",
                "/xl/worksheets/sheet1.xml");
            var workbookCoverage = new OpenXmlCoverageMessage(
                OpenXmlCoverageCodes.UnmodeledWorkbookMetadata,
                "Workbook metadata remains covered by the byte identity.",
                true,
                false,
                null,
                null,
                "/xl/workbook.xml");
            var sourceSheet = new OpenXmlWorksheetSnapshot(
                "Data",
                7U,
                "rId7",
                "/xl/worksheets/sheet1.xml",
                "DataCode",
                OpenXmlSheetVisibility.VeryHidden,
                OpenXmlPhysicalSheetKind.Worksheet,
                new RevisionSnapshot("revision-a", new[]
                {
                    firstContent,
                    interiorBlank,
                    lastContent,
                    outsideBlank,
                    farOutsideBlank,
                    invalidCoordinate,
                }),
                new List<OpenXmlFormulaObservation> { formulaObservation },
                new List<OpenXmlTableFormulaObservation> { tableObservation },
                new List<OpenXmlCoverageMessage> { sheetCoverage });
            var source = new OpenXmlWorkbookSnapshot(
                "revision-a",
                "sha256-original-package",
                ExcelDateSystem.Date1904,
                true,
                new List<OpenXmlWorksheetSnapshot> { sourceSheet },
                new List<OpenXmlCoverageMessage> { workbookCoverage });

            OpenXmlWorkbookCompactionResult result =
                OpenXmlWorkbookSnapshotCompactor.CompactToUsefulRanges(source);

            OpenXmlWorksheetCompaction worksheet = Assert.Single(result.Worksheets);
            Assert.Equal("A1:C3", worksheet.Range.DisplayAddress);
            Assert.Equal(6, worksheet.Range.TotalStoredCellCount);
            Assert.Equal(2, worksheet.Range.ContentCellCount);
            Assert.Equal(3, worksheet.Range.RetainedStoredCellCount);
            Assert.Equal(2, worksheet.Range.IgnoredOutsideRangeCellCount);
            Assert.Equal(1, worksheet.Range.InvalidCoordinateCellCount);

            OpenXmlWorkbookSnapshot compacted = result.Snapshot;
            OpenXmlWorksheetSnapshot compactedSheet = Assert.Single(compacted.Sheets);
            Assert.Equal(4, compactedSheet.Snapshot.Cells.Count);
            Assert.Same(firstContent, compactedSheet.Snapshot.Cells["A1"]);
            Assert.Same(interiorBlank, compactedSheet.Snapshot.Cells["B2"]);
            Assert.Same(lastContent, compactedSheet.Snapshot.Cells["C3"]);
            Assert.Same(invalidCoordinate, compactedSheet.Snapshot.Cells["not-a-cell"]);
            Assert.DoesNotContain("Z100", compactedSheet.Snapshot.Cells.Keys);
            Assert.DoesNotContain("XFD1048576", compactedSheet.Snapshot.Cells.Keys);

            Assert.Equal(source.RevisionId, compacted.RevisionId);
            Assert.Equal(source.ByteContentId, compacted.ByteContentId);
            Assert.Equal(source.DateSystem, compacted.DateSystem);
            Assert.Equal(source.HasVbaProject, compacted.HasVbaProject);
            Assert.Equal(sourceSheet.Name, compactedSheet.Name);
            Assert.Equal(sourceSheet.SheetId, compactedSheet.SheetId);
            Assert.Equal(sourceSheet.RelationshipId, compactedSheet.RelationshipId);
            Assert.Equal(sourceSheet.PartUri, compactedSheet.PartUri);
            Assert.Equal(sourceSheet.CodeName, compactedSheet.CodeName);
            Assert.Equal(sourceSheet.Visibility, compactedSheet.Visibility);
            Assert.Equal(sourceSheet.Kind, compactedSheet.Kind);
            Assert.Same(formulaObservation, Assert.Single(compactedSheet.FormulaObservations));
            Assert.Same(tableObservation, Assert.Single(compactedSheet.TableFormulaObservations));
            Assert.Same(sheetCoverage, Assert.Single(compactedSheet.CoverageMessages));
            Assert.Same(workbookCoverage, Assert.Single(compacted.CoverageMessages));
            Assert.False(compacted.CanAssertExactEquality);
            Assert.False(compacted.CanAssertExactCellEquality);
        }
    }
}
