using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Experimental.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental.IoBridge
{
    public class IoListFillReaderTests
    {
        /// <summary>
        /// A minimal real workbook: style 0 = no fill, style 1 = solid RGB yellow, style 2 = solid
        /// indexed red, style 3 = solid THEME fill (must be ignored), style 4 = gray125 pattern
        /// (must be ignored).
        /// </summary>
        private static byte[] BuildWorkbook()
        {
            using var stream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbook = document.AddWorkbookPart();
                workbook.Workbook = new Workbook();

                var styles = workbook.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = new Stylesheet(
                    new Fonts(new Font()),
                    new Fills(
                        new Fill(new PatternFill { PatternType = PatternValues.None }),                       // 0
                        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),                    // 1
                        new Fill(new PatternFill(
                            new ForegroundColor { Rgb = "FFFFFF00" }) { PatternType = PatternValues.Solid }), // 2 yellow
                        new Fill(new PatternFill(
                            new ForegroundColor { Indexed = 10 }) { PatternType = PatternValues.Solid }),     // 3 indexed red
                        new Fill(new PatternFill(
                            new ForegroundColor { Theme = 4 }) { PatternType = PatternValues.Solid })),       // 4 theme
                    new Borders(new Border()),
                    new CellFormats(
                        new CellFormat { FillId = 0 },     // style 0: none
                        new CellFormat { FillId = 2 },     // style 1: yellow
                        new CellFormat { FillId = 3 },     // style 2: indexed red
                        new CellFormat { FillId = 4 },     // style 3: theme (ignored)
                        new CellFormat { FillId = 1 }));   // style 4: gray125 pattern (ignored)

                var worksheetPart = workbook.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData(
                    new Row(
                        Cell("A1", 0, "plain"),
                        Cell("B1", 1, "yellow"),
                        Cell("C1", 2, "red"))
                    { RowIndex = 1 },
                    new Row(
                        Cell("B3", 3, "theme"),
                        Cell("AA3", 1, "yellow far right"),
                        Cell("AB3", 4, "gray pattern"))
                    { RowIndex = 3 }));

                workbook.Workbook.AppendChild(new Sheets(new Sheet
                {
                    Id = workbook.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "R1.S1.C1",
                }));
            }
            return stream.ToArray();
        }

        private static Cell Cell(string reference, uint style, string text) => new Cell
        {
            CellReference = reference,
            StyleIndex = style,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(text)),
        };

        [Fact]
        public void Reads_rgb_and_indexed_solid_fills_and_ignores_the_rest()
        {
            var fills = IoListFillReader.ReadFills(BuildWorkbook(), new[] { "R1.S1.C1" });

            Dictionary<(int Row, int Col), int> sheet = fills["R1.S1.C1"];
            Assert.Equal(3, sheet.Count);
            Assert.Equal(0xFFFF00, sheet[(1, 2)]);    // B1 — solid RGB
            Assert.Equal(0xFF0000, sheet[(1, 3)]);    // C1 — indexed 10
            Assert.Equal(0xFFFF00, sheet[(3, 27)]);   // AA3 — two-letter column parses
            Assert.False(sheet.ContainsKey((1, 1)));  // no fill
            Assert.False(sheet.ContainsKey((3, 2)));  // theme fill ignored
            Assert.False(sheet.ContainsKey((3, 28))); // gray125 pattern ignored
        }

        [Fact]
        public void Unknown_sheets_come_back_empty_and_garbage_bytes_do_not_throw()
        {
            var fills = IoListFillReader.ReadFills(BuildWorkbook(), new[] { "NOPE" });
            Assert.Empty(fills["NOPE"]);

            Assert.Empty(IoListFillReader.ReadFills(new byte[] { 1, 2, 3 }, new[] { "X" })["X"]);
            Assert.Empty(IoListFillReader.ReadFills(null!, new[] { "X" })["X"]);
        }

        [Fact]
        public void Reads_fills_from_the_real_sorter_io_list_when_present()
        {
            // Opportunistic integration check against the real document this feature exists for.
            const string path = @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\0 Commesse\FVZ_Intermarche\0 Documenti Elettrici\_IOlist\8FVZ_IO0001_R1.3 - sorter.xlsx";
            if (!File.Exists(path)) return;

            byte[] bytes;
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete))
            using (var buffer = new MemoryStream())
            {
                input.CopyTo(buffer);
                bytes = buffer.ToArray();
            }

            var fills = IoListFillReader.ReadFills(bytes, new[] { "R1.S1.C1" });
            Assert.True(fills.ContainsKey("R1.S1.C1"));   // extraction ran without throwing
        }
    }
}
