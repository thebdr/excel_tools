using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Projection;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class CanonicalGridProjectionTests
    {
        [Theory]
        [InlineData("A1", 1, 1)]
        [InlineData("AA27", 27, 27)]
        [InlineData("R12C4", 12, 4)]
        [InlineData("r2c3", 2, 3)]
        public void Parses_supported_physical_subject_keys(string key, int row, int column)
        {
            Assert.True(CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(key, out int actualRow, out int actualColumn));
            Assert.Equal(row, actualRow);
            Assert.Equal(column, actualColumn);
        }

        [Theory]
        [InlineData("")]
        [InlineData("A0")]
        [InlineData("R1C")]
        [InlineData("subject-hash")]
        [InlineData("XFE1")]
        public void Rejects_non_physical_subject_keys(string key)
        {
            Assert.False(CanonicalGridProjectionBuilder.TryParseSubjectCoordinate(key, out _, out _));
        }

        [Fact]
        public void Builds_a_bounded_view_with_offsets_and_missing_cells_preserved()
        {
            var snapshot = new RevisionSnapshot("r1", new[]
            {
                CanonicalCellState.Constant("B3", CanonicalValue.Text("ID")),
                CanonicalCellState.Blank("C3"),
                CanonicalCellState.Constant("B5", CanonicalValue.Text("K-2")),
                CanonicalCellState.Constant("C5", CanonicalValue.Text("tail")),
            });

            CanonicalGridProjection projection = CanonicalGridProjectionBuilder.Build(snapshot);

            Assert.Equal(3, projection.FirstPhysicalRow);
            Assert.Equal(2, projection.FirstPhysicalColumn);
            Assert.Equal(3, projection.Table.RowCount);
            Assert.Equal(2, projection.Table.ColumnCount);
            Assert.Equal("ID", projection.Table.GetCellOrNull(0, 0));
            Assert.Equal("", projection.Table.GetCellOrNull(0, 1));
            Assert.Null(projection.Table.GetCellOrNull(1, 0));
            Assert.Equal(5, projection.PhysicalRowForTableIndex(2));
            Assert.Equal(3, projection.PhysicalColumnForTableIndex(1));
        }

        [Fact]
        public void Useful_range_uses_formula_and_value_cells_but_retains_interior_blank_evidence()
        {
            var snapshot = new RevisionSnapshot("r1", new[]
            {
                CanonicalCellState.Blank("A1"),
                CanonicalCellState.Constant("B2", CanonicalValue.Text("ID")),
                CanonicalCellState.Blank("C3"),
                CanonicalCellState.Formula("D4", "1+1", CanonicalValue.Number(2)),
                CanonicalCellState.Blank("XFD1048576"),
            });

            CanonicalUsefulRange range = CanonicalUsefulRangeAnalyzer.Analyze(snapshot);
            CanonicalGridProjection projection = CanonicalGridProjectionBuilder.Build(snapshot);

            Assert.False(range.IsEmpty);
            Assert.Equal("B2:D4", range.DisplayAddress);
            Assert.Equal(3, range.RowCount);
            Assert.Equal(3, range.ColumnCount);
            Assert.Equal(9, range.GridCellCount);
            Assert.Equal(5, range.TotalStoredCellCount);
            Assert.Equal(2, range.ContentCellCount);
            Assert.Equal(3, range.RetainedStoredCellCount);
            Assert.Equal(2, range.IgnoredOutsideRangeCellCount);
            Assert.Equal(2, projection.FirstPhysicalRow);
            Assert.Equal(2, projection.FirstPhysicalColumn);
            Assert.Equal(3, projection.Table.RowCount);
            Assert.Equal(3, projection.Table.ColumnCount);
            Assert.True(projection.TryGetPhysicalCell(3, 3, out CanonicalCellState interiorBlank));
            Assert.Equal(CanonicalValueKind.Blank, interiorBlank.ConstantValue.Kind);
            Assert.False(projection.TryGetPhysicalCell(1, 1, out _));
            Assert.False(projection.TryGetPhysicalCell(1048576, 16384, out _));
        }

        [Fact]
        public void Empty_typed_text_is_content_but_a_blank_only_snapshot_has_no_useful_range()
        {
            CanonicalUsefulRange textRange = CanonicalUsefulRangeAnalyzer.Analyze(
                new RevisionSnapshot("text", new[]
                {
                    CanonicalCellState.Constant("C7", CanonicalValue.Text("")),
                }));
            CanonicalUsefulRange blankRange = CanonicalUsefulRangeAnalyzer.Analyze(
                new RevisionSnapshot("blank", new[]
                {
                    CanonicalCellState.Blank("A1"),
                    CanonicalCellState.Blank("XFD1048576"),
                }));

            Assert.Equal("C7:C7", textRange.DisplayAddress);
            Assert.Equal(1, textRange.ContentCellCount);
            Assert.True(blankRange.IsEmpty);
            Assert.Equal("(no formula/value cells)", blankRange.DisplayAddress);
            Assert.Equal(0, blankRange.GridCellCount);
            Assert.Equal(2, blankRange.IgnoredOutsideRangeCellCount);
            Assert.Equal(0, CanonicalGridProjectionBuilder.Build(
                new RevisionSnapshot("blank", new[]
                {
                    CanonicalCellState.Blank("A1"),
                    CanonicalCellState.Blank("XFD1048576"),
                })).Table.RowCount);
        }

        [Fact]
        public void Re_keys_exact_formula_and_value_evidence_through_confirmed_identities()
        {
            var snapshot = new RevisionSnapshot("r2", new[]
            {
                CanonicalCellState.Constant("A1", CanonicalValue.Text("ID")),
                CanonicalCellState.Formula("B2", "A2*2", CanonicalValue.Number(4), exactFormulaMetadata: "normal"),
            });
            CanonicalGridProjection projection = CanonicalGridProjectionBuilder.Build(snapshot);
            var rows = new Dictionary<int, string> { [1] = "header", [2] = "row:item-2" };
            var columns = new Dictionary<int, string> { [1] = "column:id", [2] = "column:result" };

            RevisionSnapshot aligned = projection.CreateAlignedSnapshot("sheet:input", rows, columns);

            Assert.Equal(2, aligned.Cells.Count);
            Assert.Contains(aligned.Cells.Values, cell => cell.ConstantValue.Equals(CanonicalValue.Text("ID")));
            Assert.Contains(aligned.Cells.Values, cell =>
                cell.FormulaText == "A2*2" &&
                cell.FormulaMetadata == "normal" &&
                cell.CachedResult.Equals(CanonicalValue.Number(4)));
        }

        [Fact]
        public void Blocks_a_sparse_extent_that_exceeds_the_materialization_budget()
        {
            var snapshot = new RevisionSnapshot("r1", new[]
            {
                CanonicalCellState.Constant("A1", CanonicalValue.Text("first")),
                CanonicalCellState.Constant("Z100", CanonicalValue.Text("last")),
            });

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                CanonicalGridProjectionBuilder.Build(snapshot, new CanonicalGridProjectionOptions
                {
                    MaxMaterializedRows = 100,
                    MaxMaterializedColumns = 26,
                    MaxMaterializedCells = 1000,
                }));
            Assert.Contains("bounded alignment view", error.Message, StringComparison.OrdinalIgnoreCase);
        }
    }
}
