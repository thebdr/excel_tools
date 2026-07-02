using System;
using System.Drawing;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Applies the CSV view styling: monospace font, header row (blue + AutoFilter + freeze), grey rows above
    /// the header, zebra striping and bold numeric cells. Striping and bold-numbers use conditional formatting
    /// so there is no slow per-cell COM loop on large files.
    /// </summary>
    internal static class CsvStyler
    {
        private const int XlExpression = 2;

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);   // Excel OLE colour (BGR)
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        private static readonly int AboveGrey = Rgb(210, 210, 210);
        private static readonly int ZebraGrey = Rgb(240, 240, 240);
        private static readonly int CommentGreen = Rgb(0, 128, 0);

        public static void Apply(dynamic app, dynamic sheet, int headerRow, long lastRow, long lastCol)
        {
            if (lastRow < 1 || lastCol < 1) return;

            dynamic all = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            all.Font.Name = MonospaceFont();
            try { all.FormatConditions.Delete(); } catch { }

            // Comment rows (first cell starts with '#'): green font, row-wide, via conditional formatting.
            dynamic comment = all.FormatConditions.Add(XlExpression, Type.Missing, "=LEFT($A1,1)=\"#\"");
            comment.Font.Color = CommentGreen;

            int dataTop = headerRow >= 1 ? headerRow + 1 : 1;

            // Zebra + bold-numbers over the data rows, via conditional formatting.
            if (dataTop <= lastRow)
            {
                dynamic data = sheet.Range(sheet.Cells(dataTop, 1), sheet.Cells(lastRow, lastCol));
                string anchor = (string)sheet.Cells(dataTop, 1).Address(false, false);   // e.g. "A2"

                dynamic zebra = data.FormatConditions.Add(XlExpression, Type.Missing, "=MOD(ROW(),2)=0");
                zebra.Interior.Color = ZebraGrey;

                dynamic bold = data.FormatConditions.Add(XlExpression, Type.Missing, "=ISNUMBER(VALUE(" + anchor + "))");
                bold.Font.Bold = true;
            }

            // Grey rows above the header.
            if (headerRow > 1)
            {
                dynamic above = sheet.Range(sheet.Cells(1, 1), sheet.Cells(headerRow - 1, lastCol));
                above.Interior.Color = AboveGrey;
            }

            // Header: blue, bold, AutoFilter, freeze.
            if (headerRow >= 1)
            {
                dynamic header = sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(headerRow, lastCol));
                header.Interior.Color = HeaderBlue;
                header.Font.Name = "Arial Narrow";   // narrow, not bold — dense header labels
                header.WrapText = true;   // long header names wrap instead of inflating the column width

                dynamic filterRange = sheet.Range(sheet.Cells(headerRow, 1), sheet.Cells(lastRow, lastCol));
                try
                {
                    if ((bool)sheet.AutoFilterMode) sheet.AutoFilterMode = false;   // drop any stale filter first
                    filterRange.AutoFilter();
                }
                catch { }

                try
                {
                    dynamic win = app.ActiveWindow;
                    win.FreezePanes = false;   // reset first, else an existing freeze mis-lands the split
                    win.SplitRow = headerRow;
                    win.SplitColumn = 0;
                    win.FreezePanes = true;
                }
                catch { }
            }

            // Column widths from the DATA rows only, so the wrapped header doesn't inflate them;
            // then a floor for very narrow columns, and the header row grows to fit its wrapped text.
            try
            {
                if (headerRow >= 1 && dataTop <= lastRow)
                    sheet.Range(sheet.Cells(dataTop, 1), sheet.Cells(lastRow, lastCol)).Columns.AutoFit();
                else
                    all.Columns.AutoFit();

                const double minColumnWidth = 8.5;
                for (int c = 1; c <= (int)lastCol; c++)
                {
                    dynamic col = sheet.Columns(c);
                    if ((double)col.ColumnWidth < minColumnWidth) col.ColumnWidth = minColumnWidth;
                }
            }
            catch { }

            if (headerRow >= 1)
                try { sheet.Rows(headerRow).AutoFit(); } catch { }
        }

        /// <summary>"Monaspace Neon Var Light Wide" if Monaspace is installed, else Consolas.</summary>
        private static string MonospaceFont()
        {
            const string wanted = "Monaspace Neon Var Light Wide";
            try
            {
                // GDI+ may list the named instance itself or only the base variable-font family —
                // either way, if any Monaspace Neon family is installed, Excel resolves the instance name.
                foreach (FontFamily f in FontFamily.Families)
                    if (f.Name.StartsWith("Monaspace Neon", StringComparison.OrdinalIgnoreCase))
                        return wanted;
            }
            catch { }
            return "Consolas";
        }
    }
}
