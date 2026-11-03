using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Project;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    /// <summary>
    /// Synthetic regression shapes inspired by read-only inspection of complex operational workbooks.
    /// The fixtures contain no business names, values, or copied implementation code.
    /// </summary>
    public sealed class Pipeline4ShapeRegressionTests
    {
        [Fact]
        public void Non_leading_header_with_a_later_data_region_blocks_instead_of_guessing()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "Operational register"),
                    Text("A3", "ID"), Text("B3", "Status"), Text("C3", "Amount"),
                    Text("A5", "A-01"), Text("B5", "Open"), Number("C5", "10"),
                    Text("A6", "A-02"), Text("B6", "Closed"), Number("C6", "20")));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "Operational register"),
                    Text("A3", "ID"), Text("B3", "Status"), Text("C3", "Amount"),
                    Text("A5", "A-01"), Text("B5", "Open"), Number("C5", "10"),
                    Text("A6", "A-02"), Text("B6", "Reviewed"), Number("C6", "20")));

            OpenXmlComparisonProjectResult result = Build("non-leading-header", before, after);

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            OpenXmlProjectSheetAlignment alignment = Assert.Single(Assert.Single(result.Edges).SheetAlignments);
            Assert.Equal(2, alignment.BeforeHeader.AcceptedHeaderRowIndex);
            Assert.Equal(2, alignment.AfterHeader.AcceptedHeaderRowIndex);
            Assert.Null(alignment.RowAlignment);
            Assert.Null(alignment.ColumnAlignment);
            OpenXmlProjectFinding finding = Assert.Single(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.NonLeadingHeader);
            Assert.Equal(DiagnosticSeverity.Correctable, finding.Severity);
        }

        [Fact]
        public void Matching_automatic_strong_non_leading_headers_preserve_predata_identity()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                Sheet("Inventory", 1, "InventoryCode",
                    Text("A1", "Warehouse inventory"),
                    Text("A2", "Prepared for weekly review"),
                    Text("A3", "ID"), Text("B3", "Status"), Text("C3", "Quantity"),
                    Text("A4", "ITM-001"), Text("B4", "Available"), Number("C4", "12"),
                    Text("A5", "ITM-002"), Text("B5", "Allocated"), Number("C5", "8"),
                    Text("A6", "ITM-003"), Text("B6", "Available"), Number("C6", "4")));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                Sheet("Inventory", 1, "InventoryCode",
                    Text("A1", "Warehouse inventory"),
                    Text("A2", "Prepared for weekly review"),
                    Text("A3", "ID"), Text("B3", "Status"), Text("C3", "Quantity"),
                    Text("A4", "ITM-001"), Text("B4", "Available"), Number("C4", "12"),
                    Text("A5", "ITM-002"), Text("B5", "Allocated"), Number("C5", "9"),
                    Text("A6", "ITM-003"), Text("B6", "Available"), Number("C6", "4")));

            OpenXmlComparisonProjectResult result = Build(
                "matching-strong-non-leading-header",
                before,
                after);

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            Assert.True(result.History!.AtomicTransitionCount > 0);
            Assert.DoesNotContain(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.NonLeadingHeader);
            Assert.DoesNotContain(result.Findings, item =>
                item.Severity >= DiagnosticSeverity.Correctable);

            OpenXmlProjectSheetAlignment alignment = Assert.Single(Assert.Single(result.Edges).SheetAlignments);
            Assert.True(alignment.IsComplete);
            Assert.Equal(HeaderDecision.StrongHeader, alignment.BeforeHeader.Decision);
            Assert.Equal(HeaderDecision.StrongHeader, alignment.AfterHeader.Decision);
            Assert.Equal(2, alignment.BeforeHeader.AcceptedHeaderRowIndex);
            Assert.Equal(2, alignment.AfterHeader.AcceptedHeaderRowIndex);
            Assert.Equal(3, alignment.BeforeHeader.DataStartRowIndex);
            Assert.Equal(3, alignment.AfterHeader.DataStartRowIndex);

            OpenXmlAlignedSheetRevision beforeSheet = Assert.Single(Revision(result, "before").Sheets);
            OpenXmlAlignedSheetRevision afterSheet = Assert.Single(Revision(result, "after").Sheets);
            for (int predataRow = 0; predataRow <= 2; predataRow++)
            {
                Assert.True(alignment.RowAlignment!.TryMapBefore(predataRow, out int mappedRow));
                Assert.Equal(predataRow, mappedRow);
                Assert.Equal(
                    beforeSheet.Rows.IdentityByPhysicalIndex[predataRow],
                    afterSheet.Rows.IdentityByPhysicalIndex[predataRow]);
            }
            Assert.Empty(afterSheet.Rows.StructuralEvents);
        }

        [Fact]
        public void Repeated_blocks_and_spacers_with_duplicate_confirmed_keys_block_row_identity()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "ID"), Text("B1", "Status"),
                    Text("A2", "A-01"), Text("B2", "Open"),
                    Text("A3", "A-02"), Text("B3", "Closed"),
                    Text("A5", "ID"), Text("B5", "Status"),
                    Text("A6", "A-01"), Text("B6", "Open"),
                    Text("A7", "A-02"), Text("B7", "Closed")));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "ID"), Text("B1", "Status"),
                    Text("A2", "A-01"), Text("B2", "Open"),
                    Text("A3", "A-02"), Text("B3", "Closed"),
                    Text("A5", "ID"), Text("B5", "Status"),
                    Text("A6", "A-01"), Text("B6", "Review"),
                    Text("A7", "A-02"), Text("B7", "Closed")));
            var options = new OpenXmlComparisonProjectOptions
            {
                HeaderDetection = new HeaderDetectionOptions { ConfirmedHeaderRowIndex = 0 },
                KeyInference = new KeyInferenceOptions
                {
                    ConfirmedBeforeColumns = new[] { 0 },
                    ConfirmedAfterColumns = new[] { 0 },
                },
            };

            OpenXmlComparisonProjectResult result = Build("repeated-blocks", before, after, options);

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            OpenXmlProjectSheetAlignment alignment = Assert.Single(Assert.Single(result.Edges).SheetAlignments);
            Assert.Equal(KeyInferenceDecision.ConfirmedKey, alignment.KeyInference.Decision);
            Assert.True(alignment.KeyInference.SelectedCandidate!.HasDuplicates);
            Assert.True(alignment.RowAlignment!.HasAmbiguity);
            OpenXmlProjectFinding finding = Assert.Single(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.RowAlignmentAmbiguous);
            Assert.Equal(DiagnosticSeverity.Correctable, finding.Severity);
            Assert.Contains(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.AxisIdentityIncomplete &&
                item.Severity == DiagnosticSeverity.Correctable);
        }

        [Fact]
        public void Numbered_sheet_family_renames_only_when_the_target_is_unambiguous()
        {
            OpenXmlWorkbookSnapshot safeBefore = Workbook(
                "safe-before",
                DataSheet("Work package 01", 1, "PackageCode", ("A-01", "Open"), ("A-02", "Closed")));
            OpenXmlWorkbookSnapshot safeAfter = Workbook(
                "safe-after",
                DataSheet("Work package 02", 9, "PackageCode", ("A-01", "Open"), ("A-02", "Closed")));

            OpenXmlComparisonProjectResult safe = Build("numbered-rename-safe", safeBefore, safeAfter);

            Assert.True(safe.IsComplete);
            ResolvedSheetMapping rename = Assert.Single(Assert.Single(safe.Edges).SheetMapping.Mapping.Mappings);
            Assert.Equal(SheetMappingKind.InferredRename, rename.Kind);
            Assert.Equal("Work package 01", rename.Before.Name);
            Assert.Equal("Work package 02", rename.After.Name);
            Assert.Equal(
                Assert.Single(Revision(safe, "safe-before").Sheets).LogicalSheetId,
                Assert.Single(Revision(safe, "safe-after").Sheets).LogicalSheetId);

            OpenXmlWorkbookSnapshot ambiguousBefore = Workbook(
                "ambiguous-before",
                DataSheet("Work package 01", 1, "PackageCode", ("A-01", "Open"), ("A-02", "Closed")));
            OpenXmlWorkbookSnapshot ambiguousAfter = Workbook(
                "ambiguous-after",
                DataSheet("Work package 02", 2, "PackageCode", ("A-01", "Open"), ("A-02", "Closed")),
                DataSheet("Work package 03", 3, "PackageCode", ("A-01", "Open"), ("A-02", "Closed")));

            OpenXmlComparisonProjectResult ambiguous =
                Build("numbered-rename-ambiguous", ambiguousBefore, ambiguousAfter);

            Assert.False(ambiguous.IsComplete);
            Assert.Null(ambiguous.History);
            OpenXmlProjectFinding finding = Assert.Single(ambiguous.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.SheetMappingAmbiguous);
            Assert.Equal(DiagnosticSeverity.Correctable, finding.Severity);
            Assert.Equal(new[] { "Work package 02", "Work package 03" }, finding.CandidateSheetNames);
            Assert.Empty(Assert.Single(ambiguous.Edges).SheetAlignments);
        }

        [Fact]
        public void Cross_sheet_membership_move_is_add_remove_never_a_within_sheet_value_edit()
        {
            OpenXmlWorkbookSnapshot before = Workbook(
                "before",
                DataSheet("North", 1, "NorthCode",
                    ("N-01", "alpha"), ("N-02", "beta"), ("N-03", "gamma"),
                    ("MOVE-07", "moved payload")),
                DataSheet("South", 2, "SouthCode",
                    ("S-01", "delta"), ("S-02", "epsilon"), ("S-03", "zeta")));
            OpenXmlWorkbookSnapshot after = Workbook(
                "after",
                DataSheet("North", 1, "NorthCode",
                    ("N-01", "alpha"), ("N-02", "beta"), ("N-03", "gamma")),
                DataSheet("South", 2, "SouthCode",
                    ("S-01", "delta"), ("S-02", "epsilon"), ("S-03", "zeta"),
                    ("MOVE-07", "moved payload")));

            OpenXmlComparisonProjectResult result = Build("cross-sheet-move", before, after);

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            LogicalSubjectHistory[] payloadSubjects = result.History!.Subjects
                .Where(subject => subject.Observations.Any(observation =>
                    string.Equals(
                        observation.ExactValue?.ExactText,
                        "moved payload",
                        StringComparison.Ordinal)))
                .ToArray();
            Assert.Equal(2, payloadSubjects.Length);
            Assert.Equal(2, payloadSubjects.Select(subject => subject.LogicalSubjectKey).Distinct().Count());
            Assert.All(payloadSubjects, subject => Assert.Single(subject.Transitions));
            Assert.Equal(
                new[] { CellChangeKind.Added, CellChangeKind.Removed },
                payloadSubjects.Select(subject => subject.Transitions[0].ExactKind).OrderBy(kind => kind).ToArray());
            Assert.DoesNotContain(payloadSubjects.SelectMany(subject => subject.Transitions), transition =>
                transition.ExactKind == CellChangeKind.Modified);
        }

        [Fact]
        public void Formula_backed_duplicate_keys_use_sequence_without_inventing_key_uniqueness()
        {
            OpenXmlWorkbookSnapshot duplicateBefore = Workbook(
                "duplicate-before",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "ID"),
                    Formula("A2", "ROW()", "2"),
                    Formula("A3", "ROW()", "3")));
            OpenXmlWorkbookSnapshot duplicateAfter = Workbook(
                "duplicate-after",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "ID"),
                    Formula("A2", "ROW()", "200"),
                    Formula("A3", "ROW()", "300")));

            OpenXmlComparisonProjectResult duplicate =
                Build("formula-key-duplicate", duplicateBefore, duplicateAfter);

            Assert.True(duplicate.IsComplete);
            Assert.NotNull(duplicate.History);
            OpenXmlProjectSheetAlignment duplicateAlignment =
                Assert.Single(Assert.Single(duplicate.Edges).SheetAlignments);
            Assert.Equal(KeyInferenceDecision.Ambiguous, duplicateAlignment.KeyInference.Decision);
            KeyCandidate duplicateCandidate = Assert.Single(duplicateAlignment.KeyInference.Candidates);
            Assert.Equal(1, duplicateCandidate.BeforeDuplicateCount);
            Assert.Equal(1, duplicateCandidate.AfterDuplicateCount);
            Assert.False(duplicateAlignment.RowAlignment!.HasAmbiguity);
            Assert.Contains(duplicateAlignment.RowAlignment.Reasons,
                item => item.Code == "row.repeated-exact-sequence");
            Assert.DoesNotContain(duplicate.Findings,
                item => item.Severity >= DiagnosticSeverity.Correctable);

            OpenXmlWorkbookSnapshot exactBefore = Workbook(
                "exact-before",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "ID"),
                    Formula("A2", "1+1", "2"),
                    Formula("A3", "1+2", "3")));
            OpenXmlWorkbookSnapshot exactAfter = Workbook(
                "exact-after",
                Sheet("Register", 1, "RegisterCode",
                    Text("A1", "ID"),
                    Formula("A2", "1+1", "999"),
                    Formula("A3", "1+2", "888")));

            OpenXmlComparisonProjectResult exact = Build("formula-key-exact", exactBefore, exactAfter);

            Assert.True(exact.IsComplete);
            Assert.NotNull(exact.History);
            OpenXmlProjectSheetAlignment exactAlignment = Assert.Single(Assert.Single(exact.Edges).SheetAlignments);
            Assert.Equal(KeyInferenceDecision.StrongKey, exactAlignment.KeyInference.Decision);
            Assert.Equal(new[] { 0 }, exactAlignment.KeyInference.SelectedCandidate!.BeforeColumns);
            OpenXmlAlignedSheetRevision beforeSheet = Assert.Single(Revision(exact, "exact-before").Sheets);
            OpenXmlAlignedSheetRevision afterSheet = Assert.Single(Revision(exact, "exact-after").Sheets);
            Assert.Equal(beforeSheet.Rows.IdentityByPhysicalIndex[1], afterSheet.Rows.IdentityByPhysicalIndex[1]);
            Assert.Equal(beforeSheet.Rows.IdentityByPhysicalIndex[2], afterSheet.Rows.IdentityByPhysicalIndex[2]);
            Assert.Equal("=1+1", beforeSheet.Grid.Table.GetCellOrNull(1, 0));
            Assert.Equal("=1+1", afterSheet.Grid.Table.GetCellOrNull(1, 0));
        }

        private static OpenXmlComparisonProjectResult Build(
            string projectId,
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after,
            OpenXmlComparisonProjectOptions? options = null) =>
            OpenXmlComparisonProjectEngine.Build(
                projectId,
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision(before.RevisionId, "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision(after.RevisionId, before.RevisionId, 1, after, isMaster: true),
                },
                options);

        private static OpenXmlAlignedProjectRevision Revision(
            OpenXmlComparisonProjectResult result,
            string revisionId) =>
            result.Revisions.Single(item =>
                string.Equals(item.Source.RevisionId, revisionId, StringComparison.Ordinal));

        private static OpenXmlWorkbookSnapshot Workbook(
            string revisionId,
            params OpenXmlWorkbookFixture.SheetDefinition[] sheets) =>
            OpenXmlWorkbookSnapshotAdapter.FromBytes(revisionId, OpenXmlWorkbookFixture.Create(sheets));

        private static OpenXmlWorkbookFixture.SheetDefinition DataSheet(
            string name,
            uint sheetId,
            string codeName,
            params (string Id, string Value)[] rows)
        {
            var cells = new List<Cell>
            {
                Text("A1", "ID"),
                Text("B1", "Value"),
            };
            for (int index = 0; index < rows.Length; index++)
            {
                int row = index + 2;
                cells.Add(Text("A" + row, rows[index].Id));
                cells.Add(Text("B" + row, rows[index].Value));
            }
            return Sheet(name, sheetId, codeName, cells.ToArray());
        }

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            uint sheetId,
            string codeName,
            params Cell[] cells) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                SheetId = sheetId,
                CodeName = codeName,
                Cells = cells,
            };

        private static Cell Text(string reference, string value) =>
            new Cell(new InlineString(new SpreadsheetText(value)))
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
            };

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell Formula(string reference, string formula, string cachedValue) =>
            new Cell(new CellFormula(formula), new CellValue(cachedValue))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };
    }
}
