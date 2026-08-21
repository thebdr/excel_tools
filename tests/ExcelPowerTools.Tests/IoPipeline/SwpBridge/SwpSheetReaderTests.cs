using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpSheetReaderTests
    {
        // ---- synthetic workbook -------------------------------------------------------------

        private static Cell Inline(string reference, string text) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new XlsxText(text)),
        };

        private static Cell Number(string reference, double value) => new Cell
        {
            CellReference = reference,
            CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture)),
        };

        private static Cell SharedString(string reference, int index) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.SharedString,
            CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture)),
        };

        private static Cell StrFormula(string reference, string formula, string cached) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellFormula = new CellFormula(formula),
            CellValue = new CellValue(cached),
        };

        private static Cell NumFormula(string reference, string formula, double cached) => new Cell
        {
            CellReference = reference,
            CellFormula = new CellFormula(formula),
            CellValue = new CellValue(cached.ToString(CultureInfo.InvariantCulture)),
        };

        private static Cell SharedStrFormula(string reference, string cached) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellFormula = new CellFormula("") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U },
            CellValue = new CellValue(cached),
        };

        private static Cell UncachedFormula(string reference) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellFormula = new CellFormula("$B$4&\".Alarm1\""),
        };

        private static byte[] BuildWorkbook(bool withUncachedIdentity = false)
        {
            using var stream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbook = document.AddWorkbookPart();
                workbook.Workbook = new Workbook();

                var sharedStrings = workbook.AddNewPart<SharedStringTablePart>();
                sharedStrings.SharedStringTable = new SharedStringTable(
                    new SharedStringItem(new XlsxText("Table entry:")),          // 0
                    new SharedStringItem(new XlsxText("DB_TEST_DIAG.Alarm1"))); // 1

                var rows = new List<Row>
                {
                    new Row(Inline("B4", "DB_TEST_DIAG")) { RowIndex = 4 },
                    new Row(Inline("B6", "Cabinet")) { RowIndex = 6 },

                    // Block 0 header: entry number in C, label / FUnit / Location in F / H / I.
                    new Row(
                        SharedString("B15", 0),
                        Number("C15", 0),
                        StrFormula("F15", "H15&I15", "=S1+X"),
                        StrFormula("H15", "X!A1", "=S1"),
                        Inline("I15", "+X1")) { RowIndex = 15 },

                    // Section label — no dot in A, never indexed.
                    new Row(Inline("A16", "Alarm1")) { RowIndex = 16 },

                    // Bit 0: identity via shared formula with cached text, B still a template formula.
                    new Row(
                        SharedStrFormula("A17", "DB_TEST_DIAG.Alarm1"),
                        NumFormula("B17", "[1]DB_TEST!$C$23", 0),
                        NumFormula("C17", "C15", 0),
                        Number("D17", 0),
                        StrFormula("Q17", "SUBSTITUTE([1]DB_TEST!$K$57&T17,\"%1\",F15)", "CABINET =S1+X - AL00"),
                        StrFormula("R17", "SUBSTITUTE([1]DB_TEST!$L$57&U17,\"%1\",F15)", "ARMOIRE =S1+X - AL00"),
                        Inline("T17", "AL00"),
                        Inline("U17", "AL00")) { RowIndex = 17 },

                    // Bit 1: identity via shared STRING, literal data cells, a shared Q formula, AA stamp.
                    new Row(
                        SharedString("A18", 1),
                        Inline("B18", "x"),
                        Number("C18", 0),
                        Number("D18", 1),
                        Inline("G18", "80.3.1"),
                        Inline("J18", "-F1"),
                        Inline("L18", "TS_001"),
                        SharedStrFormula("Q18", "shared cached"),
                        Inline("AA18", "2026-01-01.10:00:00")) { RowIndex = 18 },

                    // A duplicate of bit 0's triple — a broken SWP; the first row must win.
                    new Row(
                        withUncachedIdentity ? UncachedFormula("A19") : StrFormula("A19", "$B$4&\".Alarm1\"", "DB_TEST_DIAG.Alarm1"),
                        Number("C19", 0),
                        Number("D19", 0)) { RowIndex = 19 },

                    // Block 1 header far below, then one of its warning bits.
                    new Row(
                        SharedString("B181", 0),
                        Number("C181", 1),
                        StrFormula("F181", "H181&I181", "=S2+Y"),
                        StrFormula("H181", "X!A2", "=S2"),
                        Inline("I181", "+Y1")) { RowIndex = 181 },
                    new Row(
                        Inline("A183", "DB_TEST_DIAG.Warning1"),
                        Number("C183", 1),
                        Number("D183", 5)) { RowIndex = 183 },
                };

                var worksheetPart = workbook.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                foreach (Row row in rows) sheetData.AppendChild(row);
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var otherPart = workbook.AddNewPart<WorksheetPart>();
                otherPart.Worksheet = new Worksheet(new SheetData(
                    new Row(Inline("A1", "unrelated")) { RowIndex = 1 }));

                workbook.Workbook.AppendChild(new Sheets(
                    new Sheet { Id = workbook.GetIdOfPart(worksheetPart), SheetId = 1, Name = "CABINET" },
                    new Sheet { Id = workbook.GetIdOfPart(otherPart), SheetId = 2, Name = "OTHER" }));
            }
            return stream.ToArray();
        }

        private static SwpWorkbookIndex Read(byte[] bytes, List<SwpLogEntry> log = null,
            params string[] sheets)
        {
            return SwpSheetReader.Read(bytes,
                sheets.Length == 0 ? new[] { "CABINET", "GHOST" } : sheets,
                log ?? new List<SwpLogEntry>());
        }

        // ---- tests --------------------------------------------------------------------------

        [Fact]
        public void Reads_sheet_parameters_and_lists_missing_and_all_sheets()
        {
            SwpWorkbookIndex index = Read(BuildWorkbook());

            SwpSheetIndex sheet = index.Sheets["CABINET"];
            Assert.Equal("DB_TEST_DIAG", sheet.DdeTable);
            Assert.Equal("Cabinet", sheet.Description);
            Assert.Equal(new[] { "GHOST" }, index.MissingSheets);
            Assert.Contains("OTHER", index.AllSheetNames);
            Assert.False(index.Sheets.ContainsKey("OTHER")); // not requested → never indexed
        }

        [Fact]
        public void Indexes_bit_rows_with_cached_formula_identity_and_block_context()
        {
            SwpSheetIndex sheet = Read(BuildWorkbook()).Sheets["CABINET"];

            SwpRowSnapshot bit0 = sheet.RowsByKey["DB_TEST_DIAG.Alarm1>0>0"];
            Assert.Equal(17, bit0.RowNumber);
            Assert.Equal("DB_TEST_DIAG.Alarm1", bit0.TagType);
            Assert.Equal(0, bit0.TableEntry);
            Assert.Equal(0, bit0.Bit);
            Assert.True(bit0.ExportBIsFormula);
            Assert.Equal("0", bit0.ExportB);
            Assert.Equal("AL00", bit0.Desc1T);
            Assert.Equal("AL00", bit0.Desc2U);
            Assert.Contains("SUBSTITUTE", bit0.FormulaQ);
            Assert.Equal("CABINET =S1+X - AL00", bit0.CachedQ);
            Assert.False(bit0.FormulaQShared);
            Assert.Equal("", bit0.AddressG);       // cell absent entirely
            Assert.Equal("", bit0.StampAA);
            Assert.Equal("=S1", bit0.BlockFUnit);
            Assert.Equal("+X1", bit0.BlockLocation);
            Assert.Equal("=S1+X", bit0.BlockLabel);
            Assert.Equal(15, bit0.BlockHeaderRow);
        }

        [Fact]
        public void Reads_literal_cells_shared_strings_and_shared_formula_flags()
        {
            SwpSheetIndex sheet = Read(BuildWorkbook()).Sheets["CABINET"];

            SwpRowSnapshot bit1 = sheet.RowsByKey["DB_TEST_DIAG.Alarm1>0>1"];
            Assert.Equal(18, bit1.RowNumber);
            Assert.False(bit1.ExportBIsFormula);
            Assert.Equal("x", bit1.ExportB);
            Assert.Equal("80.3.1", bit1.AddressG);
            Assert.Equal("-F1", bit1.DeviceJ);
            Assert.Equal("TS_001", bit1.TsRefL);
            Assert.Equal("2026-01-01.10:00:00", bit1.StampAA);
            Assert.True(bit1.FormulaQShared);
            Assert.Equal("shared cached", bit1.CachedQ);
        }

        [Fact]
        public void Duplicate_triples_keep_the_first_row_and_are_surfaced()
        {
            var log = new List<SwpLogEntry>();
            SwpSheetIndex sheet = Read(BuildWorkbook(), log).Sheets["CABINET"];

            Assert.Equal(17, sheet.RowsByKey["DB_TEST_DIAG.Alarm1>0>0"].RowNumber);
            Assert.Contains("DB_TEST_DIAG.Alarm1>0>0", sheet.DuplicateKeys);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("both carry"));
        }

        [Fact]
        public void Second_block_header_switches_the_context()
        {
            SwpSheetIndex sheet = Read(BuildWorkbook()).Sheets["CABINET"];

            Assert.Equal(2, sheet.BlockCount);
            SwpRowSnapshot warning = sheet.RowsByKey["DB_TEST_DIAG.Warning1>1>5"];
            Assert.Equal(183, warning.RowNumber);
            Assert.Equal(181, warning.BlockHeaderRow);
            Assert.Equal("=S2", warning.BlockFUnit);
            Assert.Equal("+Y1", warning.BlockLocation);
        }

        [Fact]
        public void Bit_rows_keep_file_order_and_labels_are_never_indexed()
        {
            SwpSheetIndex sheet = Read(BuildWorkbook()).Sheets["CABINET"];

            Assert.Equal(new[] { 17, 18, 19, 183 }, sheet.BitRows.Select(r => r.RowNumber).ToArray());
            Assert.DoesNotContain(sheet.RowsByKey.Keys, k => k.StartsWith("Alarm1"));
        }

        [Fact]
        public void A_formula_identity_without_cached_value_refuses_with_a_remedy()
        {
            var ex = Assert.Throws<SwpSheetReader.SwpUnreadableException>(
                () => Read(BuildWorkbook(withUncachedIdentity: true)));
            Assert.Contains("Excel", ex.Message);
            Assert.Contains("row 19", ex.Message);
        }

        [Fact]
        public void A_garbage_sheet_elsewhere_never_breaks_the_wanted_ones()
        {
            // The field crash of 2026-08-21: a part that fails mid-read used to leave the
            // package's zip streams half-dead and the DISPOSE blew up ("Cannot access a closed
            // Stream"). Parts are copied out whole now — a broken bystander sheet is skipped
            // and the wanted sheets still index.
            byte[] bytes = BuildWorkbook();
            using (var stream = new MemoryStream())
            {
                stream.Write(bytes, 0, bytes.Length);
                using (var zip = new System.IO.Compression.ZipArchive(
                    stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
                {
                    System.IO.Compression.ZipArchiveEntry entry = zip.GetEntry("xl/worksheets/sheet2.xml")!;
                    entry.Delete();
                    System.IO.Compression.ZipArchiveEntry garbage = zip.CreateEntry("xl/worksheets/sheet2.xml");
                    using StreamWriter writer = new StreamWriter(garbage.Open());
                    writer.Write("this is <<< not xml at all");
                }
                bytes = stream.ToArray();
            }

            SwpWorkbookIndex index = SwpSheetReader.Read(bytes, new[] { "CABINET" },
                new List<SwpLogEntry>(), includeAllDiagnosticSheets: true);
            Assert.True(index.Sheets.ContainsKey("CABINET"));
            Assert.False(index.Sheets.ContainsKey("OTHER"));
        }

        // ---- the real FVZ sorter file, when this machine has it -----------------------------

        private const string RealSwpPath =
            @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\0 Commesse\FVZ_Intermarche\3 Software" +
            @"\3_Diagnostica\SWP\sorter\06_20-Aug-2026\SWP_01_Sorter_v2.xlsx";

        /// <summary>Reads even while the file is open in Excel (field sessions hold it).</summary>
        internal static byte[] ReadShared(string path)
        {
            using var input = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var buffer = new MemoryStream();
            input.CopyTo(buffer);
            return buffer.ToArray();
        }

        [Fact]
        public void Real_sorter_swp_indexes_every_target_sheet()
        {
            if (!File.Exists(RealSwpPath)) return; // other machines: nothing to verify against

            string[] targets = { "CABINET", "CHUTE", "SCANNER", "CONTROL_ZONE", "CHECK_POSITION", "CELL_TEST" };
            var log = new List<SwpLogEntry>();
            SwpWorkbookIndex index = SwpSheetReader.Read(ReadShared(RealSwpPath), targets, log);

            Assert.Empty(index.MissingSheets);
            Assert.Equal(6, index.Sheets.Count);

            SwpSheetIndex cabinet = index.Sheets["CABINET"];
            Assert.Equal("DB_CABINETS_DIAG", cabinet.DdeTable);

            // Facts verified by direct inspection of the file's XML.
            Assert.Equal(17, cabinet.RowsByKey["DB_CABINETS_DIAG.State>0>0"].RowNumber);
            Assert.Equal(50, cabinet.RowsByKey["DB_CABINETS_DIAG.Alarm1>0>0"].RowNumber);
            Assert.True(cabinet.BlockCount >= 12, "cabinet blocks: " + cabinet.BlockCount);
            Assert.True(cabinet.BitRows.Count >= 1000, "cabinet bit rows: " + cabinet.BitRows.Count);
            Assert.True(cabinet.RowsByKey["DB_CABINETS_DIAG.Alarm1>0>0"].FormulaQ.Length > 0);
        }
    }
}
