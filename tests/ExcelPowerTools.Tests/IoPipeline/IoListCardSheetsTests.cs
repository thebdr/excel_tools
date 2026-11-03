using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.IoPipeline;
using Xunit;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.IoPipeline
{
    /// <summary>
    /// Card sheets are found by NAME when the workbook has R.S.C sheets, and by CONTENT — a row-1
    /// header pair of Mnemonic + Bit, the card reader's own requirement — only when it has none.
    /// The safety PLC's IO list ("NET SAFETY 192.168.20" beside COVER, CABINET and CabinetToFill)
    /// is the shape that made the content path necessary; the sorter list with a summary tab is
    /// the shape the names-first order protects.
    /// </summary>
    public class IoListCardSheetsTests
    {
        private static Cell Inline(string reference, string text) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new XlsxText(text)),
        };

        private static Row[] CardHeader(uint rowIndex = 1) => new[]
        {
            new Row(Inline("A" + rowIndex, "Description Module "), Inline("F" + rowIndex, "ID"),
                Inline("G" + rowIndex, "Bit"), Inline("K" + rowIndex, "Description language 1  (1° part) (permanent part)"),
                Inline("O" + rowIndex, "Functional  unit"), Inline("P" + rowIndex, "Location"),
                Inline("Q" + rowIndex, "Device"), Inline("X" + rowIndex, "Mnemonic")) { RowIndex = rowIndex },
        };

        private static byte[] Build(params (string Name, Row[] Rows)[] sheets)
        {
            using var stream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbook = document.AddWorkbookPart();
                workbook.Workbook = new Workbook();
                var sheetList = new Sheets();
                uint id = 1;
                foreach ((string name, Row[] rows) in sheets)
                {
                    var part = workbook.AddNewPart<WorksheetPart>();
                    var data = new SheetData();
                    foreach (Row row in rows) data.AppendChild(row);
                    part.Worksheet = new Worksheet(data);
                    sheetList.AppendChild(new Sheet { Id = workbook.GetIdOfPart(part), SheetId = id++, Name = name });
                }
                workbook.Workbook.AppendChild(sheetList);
            }
            return stream.ToArray();
        }

        [Fact]
        public void The_safety_list_is_recognized_by_its_header_row()
        {
            byte[] workbook = Build(
                ("COVER", new[] { new Row(Inline("A1", "IO LIST"), Inline("A2", "8FVT_IO0100")) { RowIndex = 1 } }),
                ("CABINET", new[] { new Row(Inline("A1", "Location"), Inline("B1", "Description "), Inline("C1", "Table entry (SWP)")) { RowIndex = 1 } }),
                ("NET SAFETY 192.168.20", CardHeader()),
                // Carries a "…Bit…" header but no Mnemonic: the fill helper, not a card.
                ("CabinetToFill", new[] { new Row(Inline("A1", "Description language 1  (1° part) (permanent part)"),
                    Inline("B1", "Cabinet Name"), Inline("M1", "Last Bit used with no profinet")) { RowIndex = 1 } }));

            IoListCardSheets.Result cards = IoListCardSheets.Find(workbook);

            Assert.True(cards.ByContent);
            Assert.Equal(new[] { "NET SAFETY 192.168.20" }, cards.Names);
        }

        [Fact]
        public void Named_card_sheets_win_and_a_summary_tab_headed_like_a_card_stays_out()
        {
            byte[] workbook = Build(
                ("COVER", new Row[0]),
                ("R1.S1.C1 - MC1", CardHeader()),
                ("Mnemonics", CardHeader()),            // a summary tab carrying the same headers
                ("R1.S2.C1 - PC2", new Row[0]));       // named: in, whatever its content

            IoListCardSheets.Result cards = IoListCardSheets.Find(workbook);

            Assert.False(cards.ByContent);
            Assert.Equal(new[] { "R1.S1.C1 - MC1", "R1.S2.C1 - PC2" }, cards.Names);
        }

        [Fact]
        public void A_header_below_row_one_is_not_where_the_reader_looks()
        {
            byte[] workbook = Build(
                ("Sheet1", CardHeader(rowIndex: 2)),
                ("Sheet2", new[] { new Row(Inline("A1", "Bit"), Inline("B1", "Type")) { RowIndex = 1 } }));

            IoListCardSheets.Result cards = IoListCardSheets.Find(workbook);

            Assert.True(cards.ByContent);
            Assert.Empty(cards.Names);
        }

        [Fact]
        public void Shared_string_headers_read_like_inline_ones()
        {
            byte[] workbook;
            using (var stream = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var strings = workbookPart.AddNewPart<SharedStringTablePart>();
                    strings.SharedStringTable = new SharedStringTable(
                        new SharedStringItem(new XlsxText("Bit")),
                        new SharedStringItem(new XlsxText("Mnemonic")));
                    var part = workbookPart.AddNewPart<WorksheetPart>();
                    part.Worksheet = new Worksheet(new SheetData(new Row(
                        new Cell { CellReference = "G1", DataType = CellValues.SharedString, CellValue = new CellValue("0") },
                        new Cell { CellReference = "X1", DataType = CellValues.SharedString, CellValue = new CellValue("1") })
                    { RowIndex = 1 }));
                    workbookPart.Workbook.AppendChild(new Sheets(
                        new Sheet { Id = workbookPart.GetIdOfPart(part), SheetId = 1, Name = "NET SAFETY 50" }));
                }
                workbook = stream.ToArray();
            }

            Assert.Equal(new[] { "NET SAFETY 50" }, IoListCardSheets.Find(workbook).Names);
        }

        [Fact]
        public void Empty_bytes_yield_nothing_and_the_header_test_is_the_readers_pair()
        {
            Assert.Empty(IoListCardSheets.Find(null).Names);
            Assert.Empty(IoListCardSheets.Find(new byte[0]).Names);
            Assert.True(IoListCardSheets.LooksLikeCardHeader(new[] { "ID", "Bit", "Mnemonic" }));
            Assert.False(IoListCardSheets.LooksLikeCardHeader(new[] { "ID", "Bit" }));
            Assert.False(IoListCardSheets.LooksLikeCardHeader(new[] { "Mnemonic" }));
            Assert.False(IoListCardSheets.LooksLikeCardHeader(null));
        }
    }
}
