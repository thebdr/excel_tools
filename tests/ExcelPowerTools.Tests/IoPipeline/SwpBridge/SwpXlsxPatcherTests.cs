using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpXlsxPatcherTests
    {
        private const string Stamp = "2026-08-20.18:00:00";

        // ---- synthetic SWP ------------------------------------------------------------------

        private static Cell Inline(string reference, string text, uint? style = null)
        {
            var cell = new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new XlsxText(text)),
            };
            if (style.HasValue) cell.StyleIndex = style.Value;
            return cell;
        }

        private static Cell Number(string reference, double value) => new Cell
        {
            CellReference = reference,
            CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture)),
        };

        private static Cell StrFormula(string reference, string formula, string cached, uint? style = null)
        {
            var cell = new Cell
            {
                CellReference = reference,
                DataType = CellValues.String,
                CellFormula = new CellFormula(formula),
                CellValue = new CellValue(cached),
            };
            if (style.HasValue) cell.StyleIndex = style.Value;
            return cell;
        }

        private static Cell NumFormula(string reference, string formula, double cached, uint? style = null)
        {
            var cell = new Cell
            {
                CellReference = reference,
                CellFormula = new CellFormula(formula),
                CellValue = new CellValue(cached.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            };
            if (style.HasValue) cell.StyleIndex = style.Value;
            return cell;
        }

        private static Cell SharedStrFormula(string reference, string cached) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellFormula = new CellFormula("") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U },
            CellValue = new CellValue(cached),
        };

        private static byte[] BuildWorkbook() => BuildWorkbook("+X1");

        private static byte[] BuildWorkbook(string headerLocation)
        {
            using var stream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbook = document.AddWorkbookPart();
                workbook.Workbook = new Workbook();

                var styles = workbook.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = new Stylesheet(
                    new Fonts(new Font()),
                    new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })),
                    new Borders(new Border()),
                    new CellFormats(
                        new CellFormat(), new CellFormat(), new CellFormat(),
                        new CellFormat(), new CellFormat(), new CellFormat()));

                var cabinet = workbook.AddNewPart<WorksheetPart>();
                cabinet.Worksheet = new Worksheet(
                    new Columns(new Column { Min = 7U, Max = 7U, Width = 12, CustomWidth = true, Style = 4U }),
                    new SheetData(
                        new Row(Inline("B4", "DB_TEST_DIAG")) { RowIndex = 4 },
                        new Row(
                            Inline("B15", "Table entry:"),
                            Number("C15", 0),
                            StrFormula("F15", "H15&I15", "=S1+X"),
                            StrFormula("H15", "X!A1", "=S1"),
                            Inline("I15", headerLocation)) { RowIndex = 15 },
                        new Row(
                            StrFormula("A17", "$B$4&\".Alarm1\"", "DB_TEST_DIAG.Alarm1"),
                            NumFormula("B17", "[1]DB_TEST!$C$23", 0, style: 2U),
                            NumFormula("C17", "C15", 0),
                            Number("D17", 0),
                            StrFormula("Q17", "SUBSTITUTE([1]DB_TEST!$K$57&T17,\"%1\",F15)", "CABINET =S1+X - AL00", style: 3U),
                            StrFormula("R17", "SUBSTITUTE([1]DB_TEST!$L$57&U17,\"%1\",F15)", "ARMOIRE =S1+X - AL00"),
                            Inline("T17", "AL00"),
                            Inline("U17", "AL00"))
                        { RowIndex = 17, Spans = new ListValue<StringValue> { InnerText = "1:24" } },
                        new Row(
                            StrFormula("A18", "$B$4&\".Alarm1\"", "DB_TEST_DIAG.Alarm1"),
                            Inline("B18", "x"),
                            Number("C18", 0),
                            Number("D18", 1),
                            Inline("G18", "80.3.1"),
                            SharedStrFormula("H18", "=S1"),
                            Inline("J18", "-F1"),
                            SharedStrFormula("L18", "TS_SHARED"),
                            SharedStrFormula("Q18", "shared cached"),
                            Inline("T18", "OLD TEXT")) { RowIndex = 18 }),
                    new PageMargins { Left = 0.7, Right = 0.7, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 });

                var other = workbook.AddNewPart<WorksheetPart>();
                other.Worksheet = new Worksheet(new SheetData(
                    new Row(Inline("A1", "untouched")) { RowIndex = 1 }));

                var calcChain = workbook.AddNewPart<CalculationChainPart>();
                calcChain.CalculationChain = new CalculationChain(
                    new CalculationCell { CellReference = "Q17", SheetId = 1 });

                workbook.Workbook.AppendChild(new Sheets(
                    new Sheet { Id = workbook.GetIdOfPart(cabinet), SheetId = 1, Name = "CABINET" },
                    new Sheet { Id = workbook.GetIdOfPart(other), SheetId = 2, Name = "OTHER" }));
                workbook.Workbook.CalculationProperties = new CalculationProperties { CalculationId = 191029U };
            }
            return stream.ToArray();
        }

        private static SwpAction Action(SwpField field, string oldValue, string newValue,
            string coordinate = "DB_TEST_DIAG.Alarm1>0>0", string sheet = "CABINET")
        {
            return new SwpAction
            {
                Sheet = sheet,
                CoordinateKey = coordinate,
                Field = field,
                Mnemonic = "CABINETAL00000",
                OldValue = oldValue,
                NewValue = newValue,
            };
        }

        private static SwpRowSnapshot ReRead(byte[] bytes, string key)
        {
            SwpWorkbookIndex index = SwpSheetReader.Read(bytes, new[] { "CABINET" }, new List<SwpLogEntry>());
            return index.Sheets["CABINET"].RowsByKey[key];
        }

        // ---- tests --------------------------------------------------------------------------

        [Fact]
        public void Applies_a_full_row_and_stamps_it()
        {
            var actions = new[]
            {
                Action(SwpField.ExportB, "0", "x"),
                Action(SwpField.AddressG, "", "80.3.1"),
                Action(SwpField.Desc1T, "AL00", "NEW DESC "),
                Action(SwpField.FormulaQ,
                    "SUBSTITUTE([1]DB_TEST!$K$57&T17,\"%1\",F15)",
                    "SUBSTITUTE([1]DB_TEST!$K$57&T17,\"%1\",F15)&\" \"&T17"),
            };

            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());

            Assert.Equal(4, result.AppliedCount);
            Assert.All(result.Outcomes, o => Assert.Equal(SwpActionStatus.Applied, o.Result));

            SwpRowSnapshot row = ReRead(result.PatchedBytes, "DB_TEST_DIAG.Alarm1>0>0");
            Assert.Equal("x", row.ExportB);
            Assert.False(row.ExportBIsFormula);              // formula → literal conversion
            Assert.Equal("80.3.1", row.AddressG);
            Assert.Equal("NEW DESC ", row.Desc1T);           // trailing space preserved
            Assert.EndsWith("&\" \"&T17", row.FormulaQ);
            Assert.Equal("CABINET =S1+X - AL00 NEW DESC ", row.CachedQ);  // cache uses the NEW T
            Assert.Equal(Stamp, row.StampAA);
        }

        [Fact]
        public void A_rebuilt_formula_caches_the_truncated_text_and_keeps_its_future_function_name()
        {
            // The rebuild cuts the template's trailing text off, so the cache is head + " - " +
            // the NEW T — not the whole displayed text with the description tacked on.
            string rebuilt = SwpDescriptionFormula.Build(
                Parse("SUBSTITUTE([1]DB_TEST!$K$57&T17,\"%1\",F15)"), "T17");
            var actions = new[]
            {
                Action(SwpField.Desc1T, "AL00", "NEW DESC"),
                Action(SwpField.FormulaQ, "SUBSTITUTE([1]DB_TEST!$K$57&T17,\"%1\",F15)", rebuilt),
            };

            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());

            SwpRowSnapshot row = ReRead(result.PatchedBytes, "DB_TEST_DIAG.Alarm1>0>0");
            Assert.Equal("CABINET =S1+X - NEW DESC", row.CachedQ);

            // The _xlfn. prefix must survive into the file verbatim: without it Excel reads the
            // formula as an unknown user-defined function and the cell shows #NAME?.
            Assert.Equal(rebuilt, row.FormulaQ);
            Assert.StartsWith("_xlfn.LET(", row.FormulaQ);
        }

        private static SwpDescriptionFormula.Parsed Parse(string formula)
        {
            Assert.True(SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed parsed));
            return parsed;
        }

        [Fact]
        public void Inserted_cells_land_in_column_order_with_the_column_style()
        {
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(),
                new[] { Action(SwpField.AddressG, "", "80.3.1") }, Stamp, new List<SwpLogEntry>());

            using var stream = new MemoryStream(result.PatchedBytes);
            using var document = SpreadsheetDocument.Open(stream, false);
            WorksheetPart part = FindSheet(document, "CABINET");
            Row row = part.Worksheet.Descendants<Row>().First(r => r.RowIndex.Value == 17);

            var references = row.Elements<Cell>().Select(c => c.CellReference.Value).ToList();
            Assert.Equal(new[] { "A17", "B17", "C17", "D17", "G17", "Q17", "R17", "T17", "U17", "AA17" }, references);

            Cell g = row.Elements<Cell>().First(c => c.CellReference.Value == "G17");
            Assert.Equal(4U, g.StyleIndex.Value);            // inherited from the cols element
            Assert.Null(row.Spans);                          // the stale hint is dropped
        }

        [Fact]
        public void Kept_cells_keep_their_style()
        {
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(),
                new[] { Action(SwpField.ExportB, "0", "x") }, Stamp, new List<SwpLogEntry>());

            using var stream = new MemoryStream(result.PatchedBytes);
            using var document = SpreadsheetDocument.Open(stream, false);
            Row row = FindSheet(document, "CABINET").Worksheet.Descendants<Row>().First(r => r.RowIndex.Value == 17);
            Cell b = row.Elements<Cell>().First(c => c.CellReference.Value == "B17");
            Assert.Equal(2U, b.StyleIndex.Value);
            Assert.Null(b.CellFormula);
        }

        [Fact]
        public void Three_way_check_conflicts_and_observes_external_application()
        {
            var actions = new[]
            {
                Action(SwpField.Desc1T, "SOMETHING ELSE", "NEW"),          // stale base → conflict
                Action(SwpField.ExportB, "0", "x", "DB_TEST_DIAG.Alarm1>0>1"), // already "x" → external
                Action(SwpField.TsRefL, "", "TS_001"),                     // clean → applied
            };

            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());

            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[0].Result);
            Assert.Contains("AL00", result.Outcomes[0].Note);
            Assert.Equal(SwpActionStatus.AppliedExternally, result.Outcomes[1].Result);
            Assert.Equal(SwpActionStatus.Applied, result.Outcomes[2].Result);
            Assert.Equal(1, result.AppliedCount);

            SwpRowSnapshot row = ReRead(result.PatchedBytes, "DB_TEST_DIAG.Alarm1>0>0");
            Assert.Equal("AL00", row.Desc1T);       // the conflicted write never happened
            Assert.Equal("TS_001", row.TsRefL);
        }

        [Fact]
        public void Missing_rows_and_sheets_conflict()
        {
            var actions = new[]
            {
                Action(SwpField.ExportB, "0", "x", coordinate: "DB_TEST_DIAG.Alarm1>9>9"),
                Action(SwpField.ExportB, "0", "x", sheet: "GHOST"),
            };
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());

            Assert.Null(result.PatchedBytes);
            Assert.Equal(0, result.AppliedCount);
            Assert.Contains("row not found", result.Outcomes[0].Note);
            Assert.Contains("sheet not found", result.Outcomes[1].Note);
        }

        [Fact]
        public void Shared_formulas_are_refused()
        {
            var action = Action(SwpField.FormulaQ, "", "&\" \"&T18", "DB_TEST_DIAG.Alarm1>0>1");
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), new[] { action }, Stamp, new List<SwpLogEntry>());

            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[0].Result);
            Assert.Contains("shared", result.Outcomes[0].Note);
        }

        [Fact]
        public void Calc_chain_is_deleted_and_full_recalc_forced()
        {
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(),
                new[] { Action(SwpField.ExportB, "0", "x") }, Stamp, new List<SwpLogEntry>());

            using var stream = new MemoryStream(result.PatchedBytes);
            using var document = SpreadsheetDocument.Open(stream, false);
            Assert.Null(document.WorkbookPart.CalculationChainPart);
            Assert.True(document.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad.Value);
        }

        [Fact]
        public void Untouched_sheets_and_page_setup_survive()
        {
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(),
                new[] { Action(SwpField.ExportB, "0", "x") }, Stamp, new List<SwpLogEntry>());

            using var stream = new MemoryStream(result.PatchedBytes);
            using var document = SpreadsheetDocument.Open(stream, false);

            WorksheetPart other = FindSheet(document, "OTHER");
            Assert.Equal("untouched", other.Worksheet.Descendants<Cell>().First().InnerText);

            WorksheetPart cabinet = FindSheet(document, "CABINET");
            Assert.NotNull(cabinet.Worksheet.GetFirstChild<PageMargins>());
            Assert.NotNull(cabinet.Worksheet.GetFirstChild<Columns>());
        }

        [Fact]
        public void The_patched_package_validates_cleanly()
        {
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), new[]
            {
                Action(SwpField.ExportB, "0", "x"),
                Action(SwpField.AddressG, "", "80.3.1"),
                Action(SwpField.Desc1T, "AL00", "NEW"),
            }, Stamp, new List<SwpLogEntry>());

            using var stream = new MemoryStream(result.PatchedBytes);
            using var document = SpreadsheetDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(document).ToList();
            Assert.Empty(errors);
        }

        [Fact]
        public void A_batch_with_nothing_applicable_produces_no_bytes()
        {
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(),
                new[] { Action(SwpField.Desc1T, "WRONG BASE", "NEW") }, Stamp, new List<SwpLogEntry>());
            Assert.Null(result.PatchedBytes);
        }

        [Fact]
        public void A_formula_repair_writes_the_reference_with_the_header_cache()
        {
            var action = Action(SwpField.FormulaI, "", "I15");
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), new[] { action }, Stamp, new List<SwpLogEntry>());

            Assert.Equal(1, result.AppliedCount);
            SwpRowSnapshot row = ReRead(result.PatchedBytes, "DB_TEST_DIAG.Alarm1>0>0");
            Assert.Equal("I15", row.FormulaI);
            Assert.Equal("+X1", row.CachedI);   // predicted from the block header
            Assert.Equal(Stamp, row.StampAA);
        }

        [Fact]
        public void A_header_location_repair_writes_the_value()
        {
            var action = Action(SwpField.HeaderLocationI, "X1 broken", "+X1", coordinate: "header>15");
            var result = SwpXlsxPatcher.Apply(BuildWorkbook("X1 broken"), new[] { action }, Stamp, new List<SwpLogEntry>());

            Assert.Equal(1, result.AppliedCount);
            SwpWorkbookIndex index = SwpSheetReader.Read(result.PatchedBytes, new[] { "CABINET" }, new List<SwpLogEntry>());
            Assert.Equal("+X1", index.Sheets["CABINET"].BlockHeaders[0].LocationI);
        }

        [Fact]
        public void An_intact_shared_formula_makes_the_repair_moot()
        {
            // H18 is a shared member (empty formula text) — the three-way base matches, but the
            // template is intact, so the repair must conflict instead of overwriting it.
            var action = Action(SwpField.FormulaH, "", "H15", coordinate: "DB_TEST_DIAG.Alarm1>0>1");
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), new[] { action }, Stamp, new List<SwpLogEntry>());

            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[0].Result);
            Assert.Contains("intact", result.Outcomes[0].Note);
        }

        [Fact]
        public void A_missing_block_header_conflicts()
        {
            var action = Action(SwpField.HeaderLocationI, "X", "+X", coordinate: "header>99");
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), new[] { action }, Stamp, new List<SwpLogEntry>());
            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[0].Result);
            Assert.Contains("header not found", result.Outcomes[0].Note);
        }

        // ---- the FLD group, shared cells and duplicate writes ------------------------------------

        private const string Bit1 = "DB_TEST_DIAG.Alarm1>0>1";
        private const string Bit1Group = "CABINET|" + Bit1;

        private static SwpAction Grouped(SwpField field, string oldValue, string newValue)
        {
            SwpAction action = Action(field, oldValue, newValue, Bit1);
            action.FldGroup = Bit1Group;
            return action;
        }

        [Fact]
        public void An_fld_group_applies_together_or_not_at_all()
        {
            // The device blank's base matches, the label write's does not (someone edited T):
            // applied alone, the blank would leave the device existing nowhere.
            var actions = new[]
            {
                Grouped(SwpField.DeviceJ, "-F1", ""),
                Grouped(SwpField.Desc1T, "STALE BASE", "OLD TEXT =S1+X-F1"),
                Action(SwpField.AddressG, "80.3.1", "80.3.2", Bit1),   // the row's independent write
            };

            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());

            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[0].Result);
            Assert.Contains("held back", result.Outcomes[0].Note);
            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[1].Result);
            Assert.Contains("STALE BASE", result.Outcomes[1].Note + result.Outcomes[1].Action.OldValue);
            Assert.Equal(SwpActionStatus.Applied, result.Outcomes[2].Result);
            Assert.Equal(1, result.AppliedCount);

            SwpRowSnapshot row = ReRead(result.PatchedBytes, Bit1);
            Assert.Equal("-F1", row.DeviceJ);        // the device is still there
            Assert.Equal("OLD TEXT", row.Desc1T);
            Assert.Equal("80.3.2", row.AddressG);
        }

        [Fact]
        public void A_whole_group_with_nothing_else_to_apply_yields_no_bytes()
        {
            var actions = new[]
            {
                Grouped(SwpField.DeviceJ, "-F1", ""),
                Grouped(SwpField.Desc1T, "STALE BASE", "X"),
            };
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());
            Assert.Equal(0, result.AppliedCount);
            Assert.Null(result.PatchedBytes);
        }

        [Fact]
        public void An_intact_group_applies_whole()
        {
            var actions = new[]
            {
                Grouped(SwpField.DeviceJ, "-F1", ""),
                Grouped(SwpField.Desc1T, "OLD TEXT", "OLD TEXT =S1+X-F1"),
            };
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());
            Assert.Equal(2, result.AppliedCount);
            SwpRowSnapshot row = ReRead(result.PatchedBytes, Bit1);
            Assert.Equal("", row.DeviceJ);
            Assert.Equal("OLD TEXT =S1+X-F1", row.Desc1T);
        }

        [Fact]
        public void A_shared_formula_cell_refuses_a_literal_in_any_column()
        {
            // SharedMask, wired: L18 is a shared-formula member in a column the old guard never
            // looked at — a literal over it conflicts instead of orphaning the master's dependents.
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(),
                new[] { Action(SwpField.TsRefL, "TS_SHARED", "TS_001", Bit1) }, Stamp, new List<SwpLogEntry>());

            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[0].Result);
            Assert.Contains("shared formula", result.Outcomes[0].Note);
            Assert.Contains("not editable", result.Outcomes[0].Note);
            Assert.Null(result.PatchedBytes);
        }

        [Fact]
        public void Two_actions_on_one_cell_let_only_the_first_through()
        {
            var actions = new[]
            {
                Action(SwpField.Desc1T, "AL00", "FIRST"),
                Action(SwpField.Desc1T, "AL00", "SECOND"),
            };
            var result = SwpXlsxPatcher.Apply(BuildWorkbook(), actions, Stamp, new List<SwpLogEntry>());

            Assert.Equal(SwpActionStatus.Applied, result.Outcomes[0].Result);
            Assert.Equal(SwpActionStatus.Conflict, result.Outcomes[1].Result);
            Assert.Contains("same cell", result.Outcomes[1].Note);
            Assert.Equal("FIRST", ReRead(result.PatchedBytes, "DB_TEST_DIAG.Alarm1>0>0").Desc1T);
        }

        // ---- the real FVZ sorter file, fully in memory (nothing on disk is touched) ---------

        private const string RealSwpPath =
            @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\0 Commesse\FVZ_Intermarche\3 Software" +
            @"\3_Diagnostica\SWP\sorter\06_20-Aug-2026\SWP_01_Sorter_v2.xlsx";

        [Fact]
        public void Real_swp_streamed_patch_round_trips()
        {
            if (!File.Exists(RealSwpPath)) return;

            byte[] original = SwpSheetReaderTests.ReadShared(RealSwpPath);
            var reader0 = SwpSheetReader.Read(original, new[] { "CABINET" }, new List<SwpLogEntry>());
            SwpRowSnapshot before = reader0.Sheets["CABINET"].RowsByKey["DB_CABINETS_DIAG.Alarm1>0>0"];

            var action = new SwpAction
            {
                Sheet = "CABINET",
                CoordinateKey = "DB_CABINETS_DIAG.Alarm1>0>0",
                Field = SwpField.Desc1T,
                OldValue = before.Desc1T,
                NewValue = before.Desc1T + " PATCH-TEST",
            };
            var result = SwpXlsxPatcher.Apply(original, new[] { action }, Stamp, new List<SwpLogEntry>());

            Assert.Equal(1, result.AppliedCount);
            var reader1 = SwpSheetReader.Read(result.PatchedBytes,
                new[] { "CABINET", "CHUTE" }, new List<SwpLogEntry>());

            SwpRowSnapshot after = reader1.Sheets["CABINET"].RowsByKey["DB_CABINETS_DIAG.Alarm1>0>0"];
            Assert.Equal(before.Desc1T + " PATCH-TEST", after.Desc1T);
            Assert.Equal(Stamp, after.StampAA);

            // The untouched 15k-row CHUTE sheet still reads intact.
            Assert.True(reader1.Sheets["CHUTE"].BitRows.Count > 1000);

            using var stream = new MemoryStream(result.PatchedBytes);
            using var document = SpreadsheetDocument.Open(stream, false);
            Assert.Null(document.WorkbookPart.CalculationChainPart);
            Assert.True(document.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad.Value);
        }

        private static WorksheetPart FindSheet(SpreadsheetDocument document, string name)
        {
            Sheet sheet = document.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                .First(s => s.Name.Value == name);
            return (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
        }
    }
}
