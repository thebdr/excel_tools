using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public class OpenXmlWorkbookSnapshotAdapterTests
    {
        [Fact]
        public void Reads_typed_constants_shared_and_inline_strings_and_formula_caches()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[]
                {
                    Sheet("Data", 7, "DataCode",
                        Number("A1", "42.5"),
                        Boolean("B1", true),
                        Error("C1", "#DIV/0!"),
                        SharedString("D1", 0),
                        InlineString("E1", "inline value"),
                        Formula("F1", "A1*2", "85", CellValues.Number),
                        Formula("G1", "\"hello\"", "hello", CellValues.String))
                },
                sharedStrings: new[] { "shared value" });
            byte[] original = (byte[])bytes.Clone();

            OpenXmlWorkbookSnapshot result = OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-a", bytes);

            Assert.Equal(original, bytes); // caller bytes were never exposed to the package API
            Assert.Equal(StableId.ComputeByteContentId(bytes), result.ByteContentId);
            Assert.Equal(ExcelDateSystem.Date1900, result.DateSystem);
            Assert.True(result.CanAssertExactEquality);

            OpenXmlWorksheetSnapshot sheet = Assert.Single(result.Sheets);
            Assert.Equal(CanonicalValue.Number(42.5), sheet.Snapshot.Cells["A1"].ConstantValue);
            Assert.Equal(CanonicalValue.Boolean(true), sheet.Snapshot.Cells["B1"].ConstantValue);
            Assert.Equal(CanonicalValue.Error("#DIV/0!"), sheet.Snapshot.Cells["C1"].ConstantValue);
            Assert.Equal(CanonicalValue.Text("shared value"), sheet.Snapshot.Cells["D1"].ConstantValue);
            Assert.Equal(CanonicalValue.Text("inline value"), sheet.Snapshot.Cells["E1"].ConstantValue);

            CanonicalCellState numericFormula = sheet.Snapshot.Cells["F1"];
            Assert.Equal("A1*2", numericFormula.FormulaText); // exact text stored in the OOXML <f> element
            Assert.Equal(CanonicalValue.Number(85), numericFormula.CachedResult);
            Assert.Equal(CanonicalValue.Blank, numericFormula.ConstantValue);

            CanonicalCellState textFormula = sheet.Snapshot.Cells["G1"];
            Assert.Equal("\"hello\"", textFormula.FormulaText);
            Assert.Equal(CanonicalValue.Text("hello"), textFormula.CachedResult);
        }

        [Fact]
        public void Enumerates_visible_hidden_and_very_hidden_sheet_metadata()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet("Visible", 11, "VisibleCode", InlineString("A1", "v")),
                Sheet("Hidden", 12, "HiddenCode", SheetStateValues.Hidden, InlineString("A1", "h")),
                Sheet("VeryHidden", 13, null, SheetStateValues.VeryHidden, InlineString("A1", "vh")),
            });

            OpenXmlWorkbookSnapshot result = OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-a", bytes);

            Assert.Collection(
                result.Sheets,
                visible =>
                {
                    Assert.Equal("Visible", visible.Name);
                    Assert.Equal(11U, visible.SheetId);
                    Assert.Equal("VisibleCode", visible.CodeName);
                    Assert.Equal(OpenXmlSheetVisibility.Visible, visible.Visibility);
                    Assert.Equal(OpenXmlPhysicalSheetKind.Worksheet, visible.Kind);
                    Assert.False(string.IsNullOrWhiteSpace(visible.RelationshipId));
                    Assert.StartsWith("/xl/worksheets/", visible.PartUri, StringComparison.Ordinal);
                },
                hidden =>
                {
                    Assert.Equal("Hidden", hidden.Name);
                    Assert.Equal(12U, hidden.SheetId);
                    Assert.Equal("HiddenCode", hidden.CodeName);
                    Assert.Equal(OpenXmlSheetVisibility.Hidden, hidden.Visibility);
                },
                veryHidden =>
                {
                    Assert.Equal("VeryHidden", veryHidden.Name);
                    Assert.Equal(13U, veryHidden.SheetId);
                    Assert.Null(veryHidden.CodeName);
                    Assert.Equal(OpenXmlSheetVisibility.VeryHidden, veryHidden.Visibility);
                });
        }

        [Fact]
        public void Date_serials_retain_the_workbook_1900_or_1904_context_conservatively()
        {
            byte[] date1900Bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, StyledNumber("A1", "45000", 1), Number("B1", "45000")) },
                includeDateStyle: true);
            byte[] date1904Bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, StyledNumber("A1", "45000", 1), Number("B1", "45000")) },
                date1904: true,
                includeDateStyle: true);

            OpenXmlWorkbookSnapshot date1900 = OpenXmlWorkbookSnapshotAdapter.FromBytes("r1900", date1900Bytes);
            OpenXmlWorkbookSnapshot date1904 = OpenXmlWorkbookSnapshotAdapter.FromBytes("r1904", date1904Bytes);

            CanonicalValue value1900 = date1900.Sheets[0].Snapshot.Cells["A1"].ConstantValue;
            CanonicalValue value1904 = date1904.Sheets[0].Snapshot.Cells["A1"].ConstantValue;
            Assert.Equal(CanonicalValueKind.DateSerial, value1900.Kind);
            Assert.Equal(ExcelDateSystem.Date1900, value1900.DateSystem);
            Assert.Equal(CanonicalValueKind.DateSerial, value1904.Kind);
            Assert.Equal(ExcelDateSystem.Date1904, value1904.DateSystem);
            Assert.NotEqual(value1900, value1904);

            Assert.Equal(CanonicalValueKind.Number, date1900.Sheets[0].Snapshot.Cells["B1"].ConstantValue.Kind);
            Assert.Equal(CanonicalValueKind.Number, date1904.Sheets[0].Snapshot.Cells["B1"].ConstantValue.Kind);
        }

        [Fact]
        public void Iso_dates_are_exact_values_independent_of_the_workbook_serial_date_system()
        {
            var isoCell1900 = new Cell(new CellValue("2026-07-16T14:15:16.123"))
            {
                CellReference = "A1",
                DataType = CellValues.Date,
            };
            var isoCell1904 = (Cell)isoCell1900.CloneNode(true);

            OpenXmlWorkbookSnapshot date1900 = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "r1900",
                OpenXmlWorkbookFixture.Create(new[] { Sheet("Data", 1, null, isoCell1900) }));
            OpenXmlWorkbookSnapshot date1904 = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "r1904",
                OpenXmlWorkbookFixture.Create(
                    new[] { Sheet("Data", 1, null, isoCell1904) },
                    date1904: true));

            CanonicalValue value1900 = date1900.Sheets[0].Snapshot.Cells["A1"].ConstantValue;
            CanonicalValue value1904 = date1904.Sheets[0].Snapshot.Cells["A1"].ConstantValue;
            Assert.Equal(CanonicalValueKind.IsoDateTime, value1900.Kind);
            Assert.Equal(ExcelDateSystem.NotApplicable, value1900.DateSystem);
            Assert.Equal(value1900, value1904);
            Assert.Empty(ExactComparisonEngine.Compare(
                date1900.Sheets[0].Snapshot,
                date1904.Sheets[0].Snapshot));
        }

        [Fact]
        public void Malformed_shared_string_index_is_coverage_not_silent_blank_or_equality()
        {
            byte[] firstBytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, SharedString("A1", 99)) },
                sharedStrings: new[] { "only item" });
            byte[] secondBytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, SharedString("A1", 98)) },
                sharedStrings: new[] { "only item" });

            OpenXmlWorkbookSnapshot first = OpenXmlWorkbookSnapshotAdapter.FromBytes("first", firstBytes);
            OpenXmlWorkbookSnapshot second = OpenXmlWorkbookSnapshotAdapter.FromBytes("second", secondBytes);

            OpenXmlCoverageMessage message = Assert.Single(first.Sheets[0].CoverageMessages);
            Assert.Equal(OpenXmlCoverageCodes.InvalidSharedStringIndex, message.Code);
            Assert.Equal("A1", message.CellReference);
            Assert.True(message.AffectsExactComparison);
            Assert.False(first.Sheets[0].CanAssertExactCellEquality);

            CanonicalValue retained = first.Sheets[0].Snapshot.Cells["A1"].ConstantValue;
            Assert.Equal(CanonicalValueKind.Text, retained.Kind);
            Assert.StartsWith("#OPENXML-COVERAGE[", retained.ExactText, StringComparison.Ordinal);
            Assert.Single(ExactComparisonEngine.Compare(first.Sheets[0].Snapshot, second.Sheets[0].Snapshot));
        }

        [Fact]
        public void Unsupported_cell_placeholder_is_local_and_does_not_change_with_unrelated_package_content()
        {
            byte[] firstBytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, SharedString("A1", 99), Number("B1", "1")) },
                sharedStrings: new[] { "only item" });
            byte[] secondBytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, SharedString("A1", 99), Number("B1", "2")) },
                sharedStrings: new[] { "only item" });

            OpenXmlWorksheetSnapshot first =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("first", firstBytes).Sheets[0];
            OpenXmlWorksheetSnapshot second =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("second", secondBytes).Sheets[0];

            Assert.Equal(
                first.Snapshot.Cells["A1"].ConstantValue,
                second.Snapshot.Cells["A1"].ConstantValue);
            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(first.Snapshot, second.Snapshot));
            Assert.Equal("B1", change.SubjectKey);
        }

        [Fact]
        public void Shared_array_and_data_table_formulas_are_retained_but_explicitly_incomplete()
        {
            var shared = new Cell(
                new CellFormula("A1+1") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U },
                new CellValue("2"))
            {
                CellReference = "A1",
                DataType = CellValues.Number,
            };
            var array = new Cell(
                new CellFormula("SUM(B1:B2)") { FormulaType = CellFormulaValues.Array, Reference = "B1:B2" },
                new CellValue("3"))
            {
                CellReference = "B1",
                DataType = CellValues.Number,
            };
            var dataTable = new Cell(
                new CellFormula("ROW()") { FormulaType = CellFormulaValues.DataTable, Reference = "C1:C2" },
                new CellValue("1"))
            {
                CellReference = "C1",
                DataType = CellValues.Number,
            };

            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, shared, array, dataTable) });

            OpenXmlWorksheetSnapshot result =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-a", bytes).Sheets[0];

            Assert.Equal("A1+1", result.Snapshot.Cells["A1"].FormulaText);
            Assert.Equal("SUM(B1:B2)", result.Snapshot.Cells["B1"].FormulaText);
            Assert.Equal("ROW()", result.Snapshot.Cells["C1"].FormulaText);
            Assert.Contains(result.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded && item.CellReference == "A1");
            Assert.Contains(result.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.ArrayFormulaNotExpanded && item.CellReference == "B1");
            Assert.Contains(result.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.DataTableFormulaUnsupported && item.CellReference == "C1");
            Assert.False(result.CanAssertExactCellEquality);
        }

        [Fact]
        public void Formula_group_attributes_are_exact_evidence_separate_from_formula_text()
        {
            var beforeFormula = new Cell(
                new CellFormula("SUM(A1:A2)")
                {
                    FormulaType = CellFormulaValues.Array,
                    Reference = "B1:B2",
                },
                new CellValue("3"))
            {
                CellReference = "B1",
                DataType = CellValues.Number,
            };
            var afterFormula = new Cell(
                new CellFormula("SUM(A1:A2)")
                {
                    FormulaType = CellFormulaValues.Array,
                    Reference = "B1:B3",
                },
                new CellValue("3"))
            {
                CellReference = "B1",
                DataType = CellValues.Number,
            };

            OpenXmlWorksheetSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[] { Sheet("Data", 1, null, beforeFormula) })).Sheets[0];
            OpenXmlWorksheetSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[] { Sheet("Data", 1, null, afterFormula) })).Sheets[0];

            Assert.Equal(before.Snapshot.Cells["B1"].FormulaText, after.Snapshot.Cells["B1"].FormulaText);
            Assert.NotEqual(before.Snapshot.Cells["B1"].FormulaMetadata, after.Snapshot.Cells["B1"].FormulaMetadata);
            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));
            Assert.Equal(CellChangeAspect.FormulaStructure, change.Aspect);
        }

        [Fact]
        public void Missing_or_duplicate_sheetData_and_cell_payloads_block_empty_equality()
        {
            OpenXmlWorkbookFixture.SheetDefinition missing = Sheet("Missing", 1, null);
            missing.OmitSheetData = true;
            OpenXmlWorkbookFixture.SheetDefinition duplicate = Sheet("Duplicate", 2, null);
            duplicate.AddSecondSheetData = true;
            var malformedCell = new Cell(new CellValue("1"), new CellValue("2"))
            {
                CellReference = "A1",
                DataType = CellValues.Number,
            };
            OpenXmlWorkbookFixture.SheetDefinition malformed = Sheet("Malformed", 3, null, malformedCell);

            OpenXmlWorkbookSnapshot result = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "revision-malformed",
                OpenXmlWorkbookFixture.Create(new[] { missing, duplicate, malformed }));

            Assert.Contains(result.Sheets[0].CoverageMessages,
                finding => finding.Code == OpenXmlCoverageCodes.MissingSheetData);
            Assert.Contains(result.Sheets[1].CoverageMessages,
                finding => finding.Code == OpenXmlCoverageCodes.MultipleSheetData);
            Assert.Contains(result.Sheets[2].CoverageMessages,
                finding => finding.Code == OpenXmlCoverageCodes.MalformedCellPayload &&
                           finding.CellReference == "A1");
            Assert.All(result.Sheets, sheet => Assert.False(sheet.CanAssertExactCellEquality));
        }

        [Fact]
        public void Macro_enabled_package_is_read_without_inspecting_or_executing_vba()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, "DataCode", InlineString("A1", "safe")) },
                macroEnabled: true);

            OpenXmlWorkbookSnapshot result = OpenXmlWorkbookSnapshotAdapter.FromBytes("macro-revision", bytes);

            Assert.True(result.HasVbaProject);
            Assert.Equal("safe", result.Sheets[0].Snapshot.Cells["A1"].ConstantValue.ExactText);
            Assert.Contains(result.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.VbaProjectNotInspected && item.AffectsExactComparison);
            Assert.False(result.CanAssertExactEquality);
            Assert.True(result.CanAssertExactCellEquality);
        }

        [Fact]
        public void Unmodeled_inline_workbook_metadata_blocks_only_broad_artifact_equality()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, Formula("A1", "Rate*2", "10", CellValues.Number)) });
            bytes = AddWorkbookElements(
                bytes,
                new DefinedNames(new DefinedName("5") { Name = "Rate" }),
                new CalculationProperties { CalculationId = 191029U });

            OpenXmlWorkbookSnapshot result =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-metadata", bytes);

            Assert.Contains(result.CoverageMessages,
                message => message.Code == OpenXmlCoverageCodes.UnmodeledWorkbookMetadata &&
                           !message.AffectsExactCellComparison);
            Assert.False(result.CanAssertExactEquality);
            Assert.True(result.CanAssertExactCellEquality);
            Assert.Equal("Rate*2", result.Sheets[0].Snapshot.Cells["A1"].FormulaText);
        }

        [Fact]
        public void Mapped_physical_sheet_snapshots_feed_the_exact_engine_by_cell_reference()
        {
            byte[] beforeBytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Old Name", 4, "StableCode", InlineString("A1", "same"), Number("B2", "10")) });
            byte[] afterBytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("New Name", 9, "StableCode", InlineString("A1", "same"), Number("B2", "11")) });

            OpenXmlWorksheetSnapshot before =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("before", beforeBytes).Sheets[0];
            OpenXmlWorksheetSnapshot after =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("after", afterBytes).Sheets[0];

            AtomicCellChange change = Assert.Single(ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));
            Assert.Equal("B2", change.SubjectKey);
            Assert.Equal(CellChangeAspect.Value, change.Aspect);
            Assert.Equal("10", change.BeforeValue!.ExactText);
            Assert.Equal("11", change.AfterValue!.ExactText);
        }

        [Fact]
        public void External_worksheet_relationship_is_reported_and_never_resolved()
        {
            var definition = Sheet("Data", 1, null, InlineString("A1", "local"));
            definition.AddExternalRelationship = true;
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[] { definition });

            OpenXmlWorksheetSnapshot result =
                OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-a", bytes).Sheets[0];

            Assert.Equal("local", result.Snapshot.Cells["A1"].ConstantValue.ExactText);
            Assert.Contains(result.CoverageMessages,
                message => message.Code == OpenXmlCoverageCodes.ExternalRelationshipsIgnored);
            Assert.True(result.CanAssertExactCellEquality);
            Assert.False(
                OpenXmlWorkbookSnapshotAdapter.FromBytes("revision-b", bytes).CanAssertExactEquality);
        }

        [Fact]
        public void Package_preflight_is_bounded_and_honors_cancellation()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(
                new[] { Sheet("Data", 1, null, Number("A1", "1")) });
            using var cancellation = new System.Threading.CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "cancelled",
                bytes,
                cancellation.Token));

            byte[] excessiveEntryEnvelope = (byte[])bytes.Clone();
            int eocd = FindEndOfCentralDirectory(excessiveEntryEnvelope);
            WriteUInt16(excessiveEntryEnvelope, eocd + 8, 10001);
            WriteUInt16(excessiveEntryEnvelope, eocd + 10, 10001);
            Assert.Throws<InvalidDataException>(() => OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "bounded",
                excessiveEntryEnvelope));
        }

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

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            uint sheetId,
            string? codeName,
            SheetStateValues state,
            params Cell[] cells)
        {
            OpenXmlWorkbookFixture.SheetDefinition definition = Sheet(name, sheetId, codeName, cells);
            definition.State = state;
            return definition;
        }

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell StyledNumber(string reference, string value, uint styleIndex) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
                StyleIndex = styleIndex,
            };

        private static Cell Boolean(string reference, bool value) =>
            new Cell(new CellValue(value ? "1" : "0"))
            {
                CellReference = reference,
                DataType = CellValues.Boolean,
            };

        private static Cell Error(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Error,
            };

        private static Cell SharedString(string reference, int index) =>
            new Cell(new CellValue(index.ToString(System.Globalization.CultureInfo.InvariantCulture)))
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

        private static Cell Formula(
            string reference,
            string formula,
            string cachedValue,
            CellValues cachedType) =>
            new Cell(new CellFormula(formula), new CellValue(cachedValue))
            {
                CellReference = reference,
                DataType = cachedType,
            };

        private static int FindEndOfCentralDirectory(byte[] bytes)
        {
            for (int offset = bytes.Length - 22; offset >= 0; offset--)
            {
                if (bytes[offset] == 0x50 && bytes[offset + 1] == 0x4B &&
                    bytes[offset + 2] == 0x05 && bytes[offset + 3] == 0x06)
                    return offset;
            }
            throw new InvalidOperationException("Fixture ZIP has no EOCD record.");
        }

        private static byte[] AddWorkbookElements(byte[] sourceBytes, params OpenXmlElement[] elements)
        {
            using var stream = new MemoryStream();
            stream.Write(sourceBytes, 0, sourceBytes.Length);
            stream.Position = 0;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                foreach (OpenXmlElement element in elements)
                    document.WorkbookPart!.Workbook!.Append(element);
                document.WorkbookPart!.Workbook!.Save();
            }
            return stream.ToArray();
        }

        private static void WriteUInt16(byte[] bytes, int offset, ushort value)
        {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }
    }
}
