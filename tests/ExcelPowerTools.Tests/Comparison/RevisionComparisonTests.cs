using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.RevisionBridge;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class RevisionComparisonTests
    {
        [Fact]
        public void Compare_tracks_adjacent_values_formulas_added_removed_and_duplicate_keys()
        {
            using var files = new TemporaryFiles();
            string r1 = files.File("r1.xlsx");
            string r2 = files.File("r2.xlsx");
            string r3 = files.File("r3.xlsx");
            WriteWorkbook(r1,
                Row("K1", 10, "B2*2", 20, "alpha"),
                Row("K2", 5, "B3*2", 10, "stable"),
                Row("REMOVED", 1, "B4*2", 2, "gone"));
            WriteWorkbook(r2,
                Row("K1", 12, "B2*3", 36, "alpha"),
                Row("K_NEW", 7, "B3*2", 14, "inserted"),
                Row("K2", 5, "B4*2", 10, "stable"));
            WriteWorkbook(r3,
                Row("K1", 12, "B2*3", 36, "alpha"),
                Row("K_NEW", 8, "B3*2", 16, "inserted"),
                Row("K2", 5, "B4*2", 10, "beta"));

            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { r1, r2, r3 });
            RevisionComparisonResult result = RevisionComparisonEngine.Compare(inspection, Plan(0, 1, 2));

            Assert.Equal(2, result.KeyPreviews.Count);
            Assert.Contains(result.Actions, action =>
                action.Kind == RevisionComparisonActionKind.AddedRow && action.KeyDisplay == "K_NEW");
            Assert.Contains(result.RemovedRows, row => row.KeyDisplay == "REMOVED");
            Assert.Contains(result.Actions, action =>
                action.Kind == RevisionComparisonActionKind.CellChange &&
                action.KeyDisplay == "K1" && action.AfterColumnHeader == "Quantity");
            Assert.Contains(result.Actions, action =>
                action.Kind == RevisionComparisonActionKind.CellChange &&
                action.KeyDisplay == "K1" && action.AfterColumnHeader == "Calculated" &&
                action.AfterState.FormulaText == "B2*3");
            Assert.Contains(result.Actions, action =>
                action.Kind == RevisionComparisonActionKind.CellChange &&
                action.KeyDisplay == "K2" && action.AfterColumnHeader == "Note" &&
                action.AfterState.ConstantValue.ExactText == "beta");
        }

        [Fact]
        public void Duplicate_key_on_one_side_is_diagnostic_not_a_false_add_or_remove()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            WriteWorkbook(before,
                Row("DUP", 1, "B2*2", 2, "first"),
                Row("DUP", 2, "B3*2", 4, "second"));
            WriteWorkbook(after, Row("DUP", 3, "B2*2", 6, "single"));

            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            RevisionComparisonResult result = RevisionComparisonEngine.Compare(inspection, Plan(0, 1));

            Assert.NotEmpty(result.DuplicateKeys);
            Assert.Empty(result.Actions);
            Assert.Empty(result.RemovedRows);
        }

        [Fact]
        public void Review_round_trips_and_apply_builds_a_verified_new_workbook_with_changed_cells_and_added_rows()
        {
            using var files = new TemporaryFiles();
            string r1 = files.File("r1.xlsx");
            string r2 = files.File("r2.xlsx");
            string r3 = files.File("r3.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Applied.xlsx");
            WriteWorkbook(r1,
                Row("K1", 10, "B2*2", 20, "alpha"),
                Row("K2", 5, "B3*2", 10, "stable"));
            WriteWorkbook(r2,
                Row("K1", 12, "B2*3", 36, "alpha"),
                Row("K_NEW", 7, "B3*2", 14, "inserted"),
                Row("K2", 5, "B4*2", 10, "stable"));
            WriteWorkbook(r3,
                Row("K1", 12, "B2*3", 36, "alpha"),
                Row("K_NEW", 8, "B3*2", 16, "inserted"),
                Row("K2", 5, "B4*2", 10, "beta"));
            byte[] originalR1 = File.ReadAllBytes(r1);
            byte[] originalR2 = File.ReadAllBytes(r2);
            byte[] originalR3 = File.ReadAllBytes(r3);

            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { r1, r2, r3 });
            RevisionComparisonResult comparison = RevisionComparisonEngine.Compare(inspection, Plan(0, 1, 2));
            RevisionComparisonReviewFile.Write(review, comparison);
            Assert.Empty(RevisionComparisonReviewFile.ReadSelection(review).SelectedActionIds);
            MarkAllActionsYes(review);

            RevisionComparisonApplyResult applied = RevisionComparisonApplyEngine.ApplyReview(review, output);

            Assert.True(File.Exists(output));
            Assert.True(applied.ChangedCellCount > 0);
            Assert.Equal(1, applied.AddedRowCount);
            Assert.Equal(originalR1, File.ReadAllBytes(r1));
            Assert.Equal(originalR2, File.ReadAllBytes(r2));
            Assert.Equal(originalR3, File.ReadAllBytes(r3));

            OpenXmlWorksheetSnapshot data = ReadData(output);
            int k1 = FindRow(data, "K1");
            int k2 = FindRow(data, "K2");
            int inserted = FindRow(data, "K_NEW");
            Assert.Equal("12", Cell(data, "B", k1).ConstantValue.ExactText);
            Assert.Equal("B" + k1 + "*3", Cell(data, "C", k1).FormulaText);
            Assert.Equal("beta", Cell(data, "D", k2).ConstantValue.ExactText);
            Assert.Equal("8", Cell(data, "B", inserted).ConstantValue.ExactText);
            Assert.Equal("B" + inserted + "*2", Cell(data, "C", inserted).FormulaText);
        }

        [Fact]
        public void Apply_refuses_a_source_changed_after_review_creation_and_publishes_nothing()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Output.xlsx");
            WriteWorkbook(before, Row("K1", 1, "B2*2", 2, "before"));
            WriteWorkbook(after, Row("K1", 2, "B2*2", 4, "after"));
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            RevisionComparisonReviewFile.Write(
                review,
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1)));
            MarkAllActionsYes(review);
            SetText(after, "D2", "changed externally");

            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                RevisionComparisonApplyEngine.ApplyReview(review, output));

            Assert.Contains("changed", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(output));
        }

        [Fact]
        public void Apply_blocks_editing_a_shared_formula_group()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Output.xlsx");
            WriteSharedFormulaWorkbook(before, changed: false);
            WriteSharedFormulaWorkbook(after, changed: true);
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            RevisionComparisonReviewFile.Write(
                review,
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1)));
            MarkAllActionsYes(review);

            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                RevisionComparisonApplyEngine.ApplyReview(review, output));

            Assert.Contains("ordinary single-cell formulas", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(output));
        }

        [Fact]
        public void Formula_cells_cannot_be_used_as_silent_row_keys()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            WriteFormulaKeyWorkbook(before, "A2");
            WriteFormulaKeyWorkbook(after, "A2");
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            var plan = new RevisionComparisonPlan(
                new[] { 0, 1 },
                new[] { new RevisionComparisonPairSpec(0, 1, "Data", "Data", 1, 1, new[] { "ID" }, new[] { "ID" }) },
                0);

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                RevisionComparisonEngine.PreviewKeys(inspection, plan));
            Assert.Contains("Formula cells cannot be used", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Constant_only_apply_forces_full_recalculation_on_next_excel_open()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Output.xlsx");
            WriteWorkbook(before, Row("K1", 10, "B2*2", 20, "same"));
            WriteWorkbook(after, Row("K1", 12, "B2*2", 24, "same"));
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            RevisionComparisonReviewFile.Write(
                review,
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1)));
            MarkAllActionsYes(review);

            RevisionComparisonApplyEngine.ApplyReview(review, output);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(output, false);
            CalculationProperties? calculation = document.WorkbookPart!.Workbook.CalculationProperties;
            Assert.Equal(true, calculation?.FullCalculationOnLoad?.Value);
            Assert.Equal(true, calculation?.ForceFullCalculation?.Value);
            Assert.Equal(0U, calculation?.CalculationId?.Value);
        }

        [Fact]
        public void Apply_refuses_to_replace_an_existing_output_workbook()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Already here.xlsx");
            WriteWorkbook(before, Row("K1", 1, "B2*2", 2, "before"));
            WriteWorkbook(after, Row("K1", 2, "B2*2", 4, "after"));
            WriteWorkbook(output, Row("KEEP", 99, "B2", 99, "unrelated"));
            byte[] originalOutput = File.ReadAllBytes(output);
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            RevisionComparisonReviewFile.Write(
                review,
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1)));
            MarkAllActionsYes(review);

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                RevisionComparisonApplyEngine.ApplyReview(review, output));

            Assert.Contains("new workbook", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(originalOutput, File.ReadAllBytes(output));
        }

        [Fact]
        public void Readded_row_with_blank_content_does_not_accept_stale_existing_values()
        {
            using var files = new TemporaryFiles();
            string r1 = files.File("r1.xlsx");
            string r2 = files.File("r2.xlsx");
            string r3 = files.File("r3.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Output.xlsx");
            WriteWorkbook(r1, Row("K1", 1, "B2*2", 2, "old"));
            WriteWorkbook(r2);
            WriteWorkbook(r3, Row("K1", 1, "B2*2", 2, ""));
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { r1, r2, r3 });
            RevisionComparisonReviewFile.Write(
                review,
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1, 2)));
            MarkAllActionsYes(review);

            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                RevisionComparisonApplyEngine.ApplyReview(review, output));

            Assert.Contains("different row content", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(output));
        }

        [Fact]
        public void Compare_blocks_added_removed_or_renamed_columns_instead_of_hiding_them()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            WriteWorkbook(before, Row("K1", 1, "B2*2", 2, "before"));
            WriteWorkbook(after, Row("K1", 1, "B2*2", 2, "after"));
            SetText(after, "D1", "Renamed note");
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1)));

            Assert.Contains("same unique column headers", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Note", error.Message, StringComparison.Ordinal);
            Assert.Contains("Renamed note", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Duplicate_headers_on_an_unused_sheet_do_not_block_a_valid_table_pair()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            WriteWorkbook(before, Row("K1", 1, "B2*2", 2, "before"));
            WriteWorkbook(after, Row("K1", 2, "B2*2", 4, "after"));
            AddDuplicateHeaderSheet(before);

            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            RevisionComparisonResult result = RevisionComparisonEngine.Compare(inspection, Plan(0, 1));

            Assert.Contains(result.Actions, action => action.AfterColumnHeader == "Quantity");
        }

        [Fact]
        public void Same_named_files_have_unambiguous_revision_ordinals_in_report_transitions()
        {
            using var files = new TemporaryFiles();
            string r1 = files.File(Path.Combine("R1", "Book.xlsx"));
            string r2 = files.File(Path.Combine("R2", "Book.xlsx"));
            string r3 = files.File(Path.Combine("R3", "Book.xlsx"));
            WriteWorkbook(r1, Row("K1", 1, "B2*2", 2, "one"));
            WriteWorkbook(r2, Row("K1", 2, "B2*2", 4, "two"));
            WriteWorkbook(r3, Row("K1", 3, "B2*2", 6, "three"));

            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { r1, r2, r3 });
            RevisionComparisonResult result = RevisionComparisonEngine.Compare(inspection, Plan(0, 1, 2));

            Assert.StartsWith("R1 (Book.xlsx) -> R2 (Book.xlsx)", result.KeyPreviews[0].Transition, StringComparison.Ordinal);
            Assert.StartsWith("R2 (Book.xlsx) -> R3 (Book.xlsx)", result.KeyPreviews[1].Transition, StringComparison.Ordinal);
        }

        [Fact]
        public void Populated_column_with_blank_header_is_blocked_instead_of_ignored()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            WriteWorkbook(before, Row("K1", 1, "B2*2", 2, "before"));
            WriteWorkbook(after, Row("K1", 1, "B2*2", 2, "changed but unnamed"));
            SetText(after, "D1", "");
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                RevisionComparisonEngine.Compare(inspection, Plan(0, 1)));

            Assert.Contains("header cell", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("blank", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Explicit_key_header_rename_supports_added_row_apply()
        {
            using var files = new TemporaryFiles();
            string before = files.File("before.xlsx");
            string after = files.File("after.xlsx");
            string review = files.File("Review.xlsx");
            string output = files.File("Output.xlsx");
            WriteWorkbook(before);
            WriteWorkbook(after, Row("K1", 3, "B2*2", 6, "new"));
            SetText(before, "A1", "Legacy ID");
            SetText(after, "A1", "Current ID");
            RevisionComparisonInspection inspection = RevisionComparisonEngine.Inspect(new[] { before, after });
            var plan = new RevisionComparisonPlan(
                new[] { 0, 1 },
                new[]
                {
                    new RevisionComparisonPairSpec(
                        0, 1, "Data", "Data", 1, 1,
                        new[] { "Legacy ID" }, new[] { "Current ID" }),
                },
                0);
            RevisionComparisonResult comparison = RevisionComparisonEngine.Compare(inspection, plan);
            Assert.Single(comparison.Actions, action => action.Kind == RevisionComparisonActionKind.AddedRow);
            RevisionComparisonReviewFile.Write(review, comparison);
            MarkAllActionsYes(review);

            RevisionComparisonApplyEngine.ApplyReview(review, output);

            OpenXmlWorksheetSnapshot data = ReadData(output);
            int row = FindRow(data, "K1");
            Assert.Equal("K1", Cell(data, "A", row).ConstantValue.ExactText);
            Assert.Equal("3", Cell(data, "B", row).ConstantValue.ExactText);
        }

        private static RevisionComparisonPlan Plan(params int[] order)
        {
            var pairs = new List<RevisionComparisonPairSpec>();
            for (int index = 0; index < order.Length - 1; index++)
            {
                pairs.Add(new RevisionComparisonPairSpec(
                    order[index],
                    order[index + 1],
                    "Data",
                    "Data",
                    1,
                    1,
                    new[] { "ID" },
                    new[] { "ID" }));
            }
            return new RevisionComparisonPlan(order, pairs, order[0]);
        }

        private static TestRow Row(string id, double quantity, string formula, double cached, string note) =>
            new TestRow(id, quantity, formula, cached, note);

        private static void WriteWorkbook(string path, params TestRow[] rows)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
            var data = new SheetData();
            data.Append(new Row(
                TextCell("A1", "ID"),
                TextCell("B1", "Quantity"),
                TextCell("C1", "Calculated"),
                TextCell("D1", "Note")) { RowIndex = 1U });
            for (int index = 0; index < rows.Length; index++)
            {
                uint row = checked((uint)index + 2U);
                TestRow value = rows[index];
                data.Append(new Row(
                    TextCell("A" + row, value.Id),
                    NumberCell("B" + row, value.Quantity),
                    FormulaCell("C" + row, value.Formula, value.Cached),
                    TextCell("D" + row, value.Note)) { RowIndex = row });
            }
            worksheet.Worksheet = new Worksheet(data);
            workbook.Workbook.Append(new Sheets(new Sheet
            {
                Name = "Data",
                SheetId = 1U,
                Id = workbook.GetIdOfPart(worksheet),
            }));
            workbook.Workbook.Save();
        }

        private static void WriteSharedFormulaWorkbook(string path, bool changed)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
            var masterFormula = new CellFormula(changed ? "B2*3" : "B2*2")
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0U,
                Reference = "C2:C3",
            };
            var followerFormula = new CellFormula
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0U,
            };
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(TextCell("A1", "ID"), TextCell("B1", "Quantity"), TextCell("C1", "Calculated"), TextCell("D1", "Note")) { RowIndex = 1U },
                new Row(TextCell("A2", "K1"), NumberCell("B2", 2), new Cell(masterFormula, new CellValue(changed ? "6" : "4")) { CellReference = "C2" }, TextCell("D2", "one")) { RowIndex = 2U },
                new Row(TextCell("A3", "K2"), NumberCell("B3", 3), new Cell(followerFormula, new CellValue("6")) { CellReference = "C3" }, TextCell("D3", "two")) { RowIndex = 3U }));
            workbook.Workbook.Append(new Sheets(new Sheet { Name = "Data", SheetId = 1U, Id = workbook.GetIdOfPart(worksheet) }));
            workbook.Workbook.Save();
        }

        private static void WriteFormulaKeyWorkbook(string path, string keyFormula)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(TextCell("A1", "ID"), TextCell("B1", "Value")) { RowIndex = 1U },
                new Row(new Cell(new CellFormula(keyFormula), new CellValue("K1")) { CellReference = "A2", DataType = CellValues.String }, NumberCell("B2", 1)) { RowIndex = 2U }));
            workbook.Workbook.Append(new Sheets(new Sheet { Name = "Data", SheetId = 1U, Id = workbook.GetIdOfPart(worksheet) }));
            workbook.Workbook.Save();
        }

        private static void AddDuplicateHeaderSheet(string path)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbook = document.WorkbookPart!;
            WorksheetPart worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet(new SheetData(
                new Row(TextCell("A1", "Repeated"), TextCell("B1", "Repeated")) { RowIndex = 1U },
                new Row(TextCell("A2", "left"), TextCell("B2", "right")) { RowIndex = 2U }));
            Sheets sheets = workbook.Workbook.Sheets!;
            uint sheetId = sheets.Elements<Sheet>().Max(item => item.SheetId!.Value) + 1U;
            sheets.Append(new Sheet
            {
                Name = "Notes",
                SheetId = sheetId,
                Id = workbook.GetIdOfPart(worksheet),
            });
            workbook.Workbook.Save();
        }

        private static void MarkAllActionsYes(string reviewPath)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(reviewPath, true);
            foreach (string sheetName in new[] { RevisionComparisonReviewFile.AddedSheetName, RevisionComparisonReviewFile.ChangedSheetName })
            {
                Sheet sheet = document.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                    .Single(item => item.Name!.Value == sheetName);
                WorksheetPart part = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id!.Value!);
                foreach (Row row in part.Worksheet.Descendants<Row>().Where(item => item.RowIndex!.Value > 1U))
                {
                    Cell? action = row.Elements<Cell>().FirstOrDefault(item => item.CellReference!.Value!.StartsWith("A", StringComparison.Ordinal));
                    if (action == null || string.IsNullOrEmpty(action.InnerText)) continue;
                    string reference = "B" + row.RowIndex!.Value;
                    Cell apply = row.Elements<Cell>().FirstOrDefault(item => item.CellReference!.Value == reference) ??
                                 new Cell { CellReference = reference };
                    if (apply.Parent == null) row.Append(apply);
                    apply.RemoveAllChildren();
                    apply.DataType = CellValues.InlineString;
                    apply.InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Yes"));
                }
                part.Worksheet.Save();
            }
        }

        private static void SetText(string path, string reference, string value)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            WorksheetPart part = document.WorkbookPart!.WorksheetParts.Single();
            Cell cell = part.Worksheet.Descendants<Cell>().Single(item => item.CellReference!.Value == reference);
            cell.RemoveAllChildren();
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(value));
            part.Worksheet.Save();
        }

        private static OpenXmlWorksheetSnapshot ReadData(string path)
        {
            OpenXmlWorkbookSnapshot workbook = OpenXmlWorkbookSnapshotAdapter.FromBytes("test", File.ReadAllBytes(path));
            return workbook.Sheets.Single(item => item.Name == "Data");
        }

        private static int FindRow(OpenXmlWorksheetSnapshot sheet, string id)
        {
            for (int row = 2; row <= 50; row++)
            {
                if (sheet.Snapshot.TryGetCell("A" + row, out CanonicalCellState? state) &&
                    !state.HasFormula && state.ConstantValue.ExactText == id)
                    return row;
            }
            throw new Xunit.Sdk.XunitException("Key not found: " + id);
        }

        private static CanonicalCellState Cell(OpenXmlWorksheetSnapshot sheet, string column, int row)
        {
            Assert.True(sheet.Snapshot.TryGetCell(column + row, out CanonicalCellState? state));
            return state!;
        }

        private static Cell TextCell(string reference, string value) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(value)),
        };

        private static Cell NumberCell(string reference, double value) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.Number,
            CellValue = new CellValue(value),
        };

        private static Cell FormulaCell(string reference, string formula, double cached) => new Cell(
            new CellFormula(formula),
            new CellValue(cached))
        {
            CellReference = reference,
        };

        private sealed record TestRow(string Id, double Quantity, string Formula, double Cached, string Note);

        private sealed class TemporaryFiles : IDisposable
        {
            private readonly string _directory = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "ExcelPowerTools-RevisionComparison-" + Guid.NewGuid().ToString("N"));

            internal TemporaryFiles() => Directory.CreateDirectory(_directory);
            internal string File(string name)
            {
                string path = System.IO.Path.Combine(_directory, name);
                string? directory = System.IO.Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
                return path;
            }
            public void Dispose()
            {
                if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
            }
        }
    }
}
