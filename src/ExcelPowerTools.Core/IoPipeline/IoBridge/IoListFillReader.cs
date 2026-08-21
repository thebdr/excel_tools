using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.IoPipeline.IoBridge
{
    /// <summary>
    /// Extracts cell INTERIOR colors from an .xlsx headlessly — the electrical department uses fills
    /// for revision tracking, and that is the one piece of source formatting the bridge carries over.
    /// Direct RGB and indexed (classic-palette) solid fills are read; theme-based and patterned fills
    /// are ignored (revision highlights are manual fills, which Excel stores as RGB).
    /// </summary>
    public static class IoListFillReader
    {
        /// <summary>
        /// Per requested sheet: (1-based row, 1-based column) → RGB (0xRRGGBB) for every cell with a
        /// readable solid fill. Sheets not found come back as empty maps.
        /// </summary>
        public static Dictionary<string, Dictionary<(int Row, int Col), int>> ReadFills(
            byte[] xlsxBytes, IEnumerable<string> sheetNames)
        {
            var result = new Dictionary<string, Dictionary<(int, int), int>>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in sheetNames ?? new string[0])
                if (!result.ContainsKey(name)) result[name] = new Dictionary<(int, int), int>();
            if (xlsxBytes == null || xlsxBytes.Length == 0) return result;

            try
            {
                ReadInto(result, xlsxBytes);
            }
            catch
            {
                // Fills are decoration: a damaged or foreign file yields no colors, never a failure.
            }
            return result;
        }

        private static void ReadInto(Dictionary<string, Dictionary<(int, int), int>> result, byte[] xlsxBytes)
        {
            using (var stream = new MemoryStream(xlsxBytes, writable: false))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbook = document.WorkbookPart;
                Sheets sheets = workbook?.Workbook?.Sheets;
                if (workbook == null || sheets == null) return;

                int?[] fillRgbByStyle = BuildStyleFillTable(workbook.WorkbookStylesPart?.Stylesheet);
                if (fillRgbByStyle.Length == 0) return;

                foreach (Sheet sheet in sheets.Elements<Sheet>())
                {
                    string name = sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : null;
                    if (name == null || !result.TryGetValue(name, out Dictionary<(int, int), int> fills)) continue;
                    if (!(workbook.GetPartById(sheet.Id.Value) is WorksheetPart worksheetPart)) continue;

                    foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                    {
                        if (row.RowIndex == null) continue;
                        int rowNumber = (int)row.RowIndex.Value;

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            if (cell.StyleIndex == null) continue;
                            uint style = cell.StyleIndex.Value;
                            if (style >= fillRgbByStyle.Length || fillRgbByStyle[style] == null) continue;

                            int column = ColumnIndexOf(cell.CellReference?.Value);
                            if (column > 0) fills[(rowNumber, column)] = fillRgbByStyle[style].Value;
                        }
                    }
                }
            }
        }

        /// <summary>Per cellXfs index: the solid-fill RGB, or null when the style paints nothing readable.</summary>
        private static int?[] BuildStyleFillTable(Stylesheet stylesheet)
        {
            CellFormats cellFormats = stylesheet?.CellFormats;
            Fills fills = stylesheet?.Fills;
            if (cellFormats == null || fills == null) return new int?[0];

            var fillRgb = new List<int?>();
            foreach (Fill fill in fills.Elements<Fill>())
                fillRgb.Add(SolidRgbOf(fill));

            var result = new List<int?>();
            foreach (CellFormat format in cellFormats.Elements<CellFormat>())
            {
                int fillId = format.FillId != null ? (int)format.FillId.Value : 0;
                result.Add(fillId >= 0 && fillId < fillRgb.Count ? fillRgb[fillId] : null);
            }
            return result.ToArray();
        }

        private static int? SolidRgbOf(Fill fill)
        {
            PatternFill pattern = fill?.PatternFill;
            if (pattern?.PatternType == null || pattern.PatternType.Value != PatternValues.Solid) return null;

            ForegroundColor color = pattern.ForegroundColor;
            if (color == null) return null;

            if (color.Rgb != null && color.Rgb.HasValue && color.Rgb.Value.Length >= 6)
            {
                string hex = color.Rgb.Value;
                hex = hex.Substring(hex.Length - 6);   // ARGB → RGB
                return int.TryParse(hex, System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture, out int rgb) ? rgb : (int?)null;
            }

            if (color.Indexed != null && color.Indexed.HasValue)
            {
                uint index = color.Indexed.Value;
                return index < IndexedPalette.Length ? IndexedPalette[index] : (int?)null;
            }

            return null;   // theme/tint fills are not resolved — see the class summary
        }

        private static int ColumnIndexOf(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            int column = 0;
            foreach (char c in cellReference)
            {
                if (c >= 'A' && c <= 'Z') column = column * 26 + (c - 'A' + 1);
                else if (c >= 'a' && c <= 'z') column = column * 26 + (c - 'a' + 1);
                else break;
            }
            return column;
        }

        /// <summary>The classic indexed palette (0–63) as 0xRRGGBB.</summary>
        private static readonly int[] IndexedPalette =
        {
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
            0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
            0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
            0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
            0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
            0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
            0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
        };
    }
}
