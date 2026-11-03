using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.IoPipeline
{
    /// <summary>
    /// Which worksheets of an IO-list workbook are card sheets — the sheets a bridge reads its
    /// signal rows from. The NAME decides when it can: "R{rack}.S{slot}.C{card}" is the electrical
    /// department's convention and carries the card identity the CSV joins need. When a workbook
    /// names no such sheet, the CONTENT decides instead: a sheet whose header row resolves both a
    /// Mnemonic and a Bit column is a card sheet, because that pair is exactly what the card
    /// readers require before they yield a row (<see cref="IoListLogic.FindColumn"/> answers both
    /// questions here and there, so a listed sheet is one the reader will read and a sheet it
    /// would skip is never offered). The safety PLC's list is the case: its one card sheet is
    /// named by network ("NET SAFETY 192.168.20") and nothing else in that workbook — cover,
    /// cabinet table, fill helper — carries both headers.
    ///
    /// Names first, content only when the names name nothing — the order <c>CsvColumnsMap.Locate</c>
    /// follows too: a sorter list with real R.S.C sheets never grows an extra ticked sheet because
    /// a summary tab happens to carry a Mnemonic column.
    /// </summary>
    public static class IoListCardSheets
    {
        public sealed class Result
        {
            /// <summary>The card sheets, in workbook order.</summary>
            public List<string> Names = new List<string>();

            /// <summary>True when no R.S.C name matched and the header rows decided instead.</summary>
            public bool ByContent;
        }

        /// <summary>The reader's own criterion: a header row naming a Mnemonic and a Bit column.</summary>
        public static bool LooksLikeCardHeader(IReadOnlyList<string> headers) =>
            headers != null && headers.Count > 0 &&
            IoListLogic.FindColumn(headers, "mnemonic") >= 0 &&
            IoListLogic.FindColumn(headers, "bit") >= 0;

        /// <summary>
        /// The card sheets of a workbook given as bytes (an in-memory copy reads while Excel holds
        /// the file). Throws when the bytes are not a workbook — the caller says so.
        /// </summary>
        public static Result Find(byte[] xlsxBytes)
        {
            var result = new Result();
            if (xlsxBytes == null || xlsxBytes.Length == 0) return result;

            using (var stream = new MemoryStream(xlsxBytes, writable: false))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbook = document.WorkbookPart;
                Sheets sheets = workbook?.Workbook?.Sheets;
                if (workbook == null || sheets == null) return result;

                var named = new List<(Sheet Sheet, string Name)>();
                foreach (Sheet sheet in sheets.Elements<Sheet>())
                {
                    string name = sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : null;
                    if (string.IsNullOrEmpty(name)) continue;
                    named.Add((sheet, name));
                    if (IoListLogic.TryParseCardSheetName(name, out _, out _, out _))
                        result.Names.Add(name);
                }
                if (result.Names.Count > 0) return result;

                result.ByContent = true;
                IReadOnlyList<string> sharedStrings = LoadSharedStrings(workbook);
                foreach ((Sheet sheet, string name) in named)
                {
                    if (!(workbook.GetPartById(sheet.Id.Value) is WorksheetPart part)) continue;
                    if (LooksLikeCardHeader(HeaderRow(part, sharedStrings))) result.Names.Add(name);
                }
            }
            return result;
        }

        /// <summary>
        /// The texts of row 1 — the row the card readers take their headers from. A sheet whose
        /// first stored row sits lower has no header where the reader looks and yields nothing.
        /// </summary>
        private static List<string> HeaderRow(WorksheetPart part, IReadOnlyList<string> sharedStrings)
        {
            var headers = new List<string>();

            // Copied out in one deterministic pass — partially-consumed package streams leave
            // WindowsBase (net48) zip state that detonates at package dispose.
            MemoryStream xml;
            try
            {
                using (Stream source = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    xml = new MemoryStream();
                    source.CopyTo(xml);
                }
            }
            catch
            {
                return headers;   // an unreadable sheet is simply not a card sheet
            }

            using (xml)
            using (var reader = new OpenXmlPartReader(
                new MemoryStream(xml.GetBuffer(), 0, (int)xml.Length, writable: false),
                part.Features, new OpenXmlPartReaderOptions()))
            {
                while (reader.Read())
                {
                    if (reader.ElementType != typeof(Row) || !reader.IsStartElement) continue;
                    var row = (Row)reader.LoadCurrentElement();
                    if (row.RowIndex != null && row.RowIndex.Value != 1) break;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        int column = ColumnOf(cell.CellReference?.Value);
                        if (column <= 0) column = headers.Count + 1;
                        while (headers.Count < column) headers.Add("");
                        headers[column - 1] = CellText(cell, sharedStrings);
                    }
                    break;
                }
            }
            return headers;
        }

        private static IReadOnlyList<string> LoadSharedStrings(WorkbookPart workbook)
        {
            var strings = new List<string>();
            SharedStringTable table = workbook.SharedStringTablePart?.SharedStringTable;
            if (table == null) return strings;
            foreach (OpenXmlElement item in table.Elements())
                strings.Add(item.InnerText ?? "");
            return strings;
        }

        private static string CellText(Cell cell, IReadOnlyList<string> sharedStrings)
        {
            if (cell == null) return "";
            string raw = cell.CellValue?.Text ?? "";
            if (cell.DataType == null) return raw;
            if (cell.DataType.Value == CellValues.SharedString)
            {
                return int.TryParse(raw, out int index) && index >= 0 && index < sharedStrings.Count
                    ? sharedStrings[index]
                    : "";
            }
            if (cell.DataType.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? cell.InnerText ?? "";
            return raw;
        }

        private static int ColumnOf(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            int column = 0;
            foreach (char ch in cellReference)
            {
                if (ch >= 'A' && ch <= 'Z') column = column * 26 + (ch - 'A' + 1);
                else break;
            }
            return column;
        }
    }
}
