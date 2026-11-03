using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    internal static class OpenXmlWorkbookFixture
    {
        internal sealed class SheetDefinition
        {
            internal string Name { get; set; } = "Sheet1";
            internal uint SheetId { get; set; } = 1;
            internal SheetStateValues? State { get; set; }
            internal string? CodeName { get; set; }
            internal IReadOnlyList<Cell> Cells { get; set; } = Array.Empty<Cell>();
            internal bool AddExternalRelationship { get; set; }
            internal bool OmitSheetData { get; set; }
            internal bool AddSecondSheetData { get; set; }
            internal bool OmitSheetId { get; set; }
            internal string? RawState { get; set; }
        }

        internal static byte[] Create(
            IReadOnlyList<SheetDefinition> sheets,
            bool date1904 = false,
            IReadOnlyList<string>? sharedStrings = null,
            bool includeDateStyle = false,
            bool macroEnabled = false)
        {
            using var stream = new MemoryStream();
            SpreadsheetDocumentType type = macroEnabled
                ? SpreadsheetDocumentType.MacroEnabledWorkbook
                : SpreadsheetDocumentType.Workbook;

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, type, true))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                if (date1904)
                    workbookPart.Workbook.Append(new WorkbookProperties { Date1904 = true });

                if (sharedStrings != null)
                {
                    SharedStringTablePart sharedPart = workbookPart.AddNewPart<SharedStringTablePart>();
                    sharedPart.SharedStringTable = new SharedStringTable(
                        sharedStrings.Select(value => new SharedStringItem(new SpreadsheetText(value))));
                }

                if (includeDateStyle)
                {
                    WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = new Stylesheet(
                        new CellFormats(
                            new CellFormat { NumberFormatId = 0U },
                            new CellFormat { NumberFormatId = 14U, ApplyNumberFormat = true })
                        { Count = 2U });
                }

                if (macroEnabled)
                {
                    VbaProjectPart vbaPart = workbookPart.AddNewPart<VbaProjectPart>();
                    using Stream vbaStream = vbaPart.GetStream(FileMode.Create, FileAccess.Write);
                    byte[] inertFixtureBytes = { 0x56, 0x42, 0x41, 0x00, 0x01 };
                    vbaStream.Write(inertFixtureBytes, 0, inertFixtureBytes.Length);
                }

                var sheetCollection = new Sheets();
                workbookPart.Workbook.Append(sheetCollection);

                foreach (SheetDefinition definition in sheets)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var worksheet = new Worksheet();
                    if (definition.CodeName != null)
                        worksheet.Append(new SheetProperties { CodeName = definition.CodeName });

                    var sheetData = new SheetData();
                    foreach (IGrouping<uint, Cell> rowCells in definition.Cells.GroupBy(CellRow).OrderBy(group => group.Key))
                    {
                        var row = new Row { RowIndex = rowCells.Key };
                        foreach (Cell cell in rowCells.OrderBy(item => item.CellReference!.Value, StringComparer.Ordinal))
                            row.Append(cell);
                        sheetData.Append(row);
                    }
                    if (!definition.OmitSheetData)
                        worksheet.Append(sheetData);
                    if (definition.AddSecondSheetData)
                        worksheet.Append(new SheetData());
                    worksheetPart.Worksheet = worksheet;

                    if (definition.AddExternalRelationship)
                    {
                        worksheetPart.AddHyperlinkRelationship(
                            new Uri("https://invalid.example.test/never-follow"),
                            true);
                    }

                    var sheet = new Sheet
                    {
                        Name = definition.Name,
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                    };
                    if (!definition.OmitSheetId) sheet.SheetId = definition.SheetId;
                    if (definition.RawState != null)
                        sheet.SetAttribute(new OpenXmlAttribute("state", "", definition.RawState));
                    else if (definition.State.HasValue)
                        sheet.State = definition.State.Value;
                    sheetCollection.Append(sheet);
                }

                workbookPart.Workbook.Save();
            }

            return stream.ToArray();
        }

        private static uint CellRow(Cell cell)
        {
            string reference = cell.CellReference?.Value ?? "A1";
            int index = 0;
            while (index < reference.Length && char.IsLetter(reference[index])) index++;
            return uint.TryParse(reference.Substring(index), out uint row) ? row : 1U;
        }
    }
}
