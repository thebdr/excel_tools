using System;
using System.Collections.Generic;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// Reads the selected card sheets of the IO-list workbook into <see cref="IoRowData"/> rows.
    /// An instance already open in THIS Excel is read live; otherwise the workbook opens hidden
    /// read-only and closes again (the IO Bridge's reading discipline, duplicated on purpose so
    /// either feature stays removable as one folder). Columns are located by header text via
    /// <see cref="IoListLogic.FindColumn"/> — never by fixed letters.
    /// </summary>
    internal static class SwpBridgeCardReader
    {
        public static List<IoRowData> Read(dynamic app, string ioListPath, List<string> sheetNames,
            List<SwpLogEntry> log)
        {
            var wanted = new HashSet<string>(sheetNames, StringComparer.OrdinalIgnoreCase);
            var rows = new List<IoRowData>();

            dynamic wb = FindOpenWorkbook(app, ioListPath);
            bool opened = false;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
            try
            {
                if (wb == null)
                {
                    wb = app.Workbooks.Open(ioListPath, 0, true);   // UpdateLinks:=0, ReadOnly:=true
                    opened = true;
                }

                foreach (dynamic ws in wb.Worksheets)
                {
                    string name;
                    try { name = (string)ws.Name; } catch { continue; }
                    if (!wanted.Contains(name)) continue;
                    ReadSheet(ws, name, rows, log);
                }
                return rows;
            }
            finally
            {
                if (opened) try { wb?.Close(false); } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
            }
        }

        internal static dynamic FindOpenWorkbook(dynamic app, string path)
        {
            try
            {
                foreach (dynamic wb in app.Workbooks)
                {
                    string full = null;
                    try { full = (string)wb.FullName; } catch { }
                    if (CsvWorkspaceLogic.SameSourcePath(full, path)) return wb;
                }
            }
            catch { }
            return null;
        }

        private static bool IsAllDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char c in s)
                if (c < '0' || c > '9') return false;
            return true;
        }

        private static void ReadSheet(dynamic sheet, string name, List<IoRowData> rows, List<SwpLogEntry> log)
        {
            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);
            if (lastRow < 2) return;

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            object raw = range.Value2;
            if (!(raw is object[,] block)) return;

            int rb = block.GetLowerBound(0), cb = block.GetLowerBound(1);
            int rowCount = (int)lastRow, colCount = (int)lastCol;

            string Cell(int r, int c) =>
                c < 0 ? "" : (string)CsvSheet.CellText(block[rb + r, cb + c], sheet, r + 1, c + 1);

            var headers = new string[colCount];
            for (int c = 0; c < colCount; c++) headers[c] = Cell(0, c);

            int mnemonic = IoListLogic.FindColumn(headers, "mnemonic");
            int bit = IoListLogic.FindColumn(headers, "bit");
            int id = IoListLogic.FindColumn(headers, "id");
            int slot = IoListLogic.FindColumn(headers, "slot");
            if (mnemonic < 0 || bit < 0)
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "sheet skipped: no Mnemonic + Bit header", name));
                return;
            }

            int type = IoListLogic.FindColumn(headers, "type");
            int desc1A = IoListLogic.FindColumn(headers, "Description language 1 (1° part", "description english language");
            int desc1B = IoListLogic.FindColumn(headers, "Description language 1 (2° part");
            int desc2A = IoListLogic.FindColumn(headers, "Description language 2 (1° part");
            int desc2B = IoListLogic.FindColumn(headers, "Description language 2 (2° part");
            int funit = IoListLogic.FindColumn(headers, "functional unit");
            int location = IoListLogic.FindColumn(headers, "location");
            int device = IoListLogic.FindColumn(headers, "device");
            int tsRef = IoListLogic.FindColumn(headers, "t.s.");
            int pnIp = IoListLogic.FindColumn(headers, "profinet ip");
            int pnName = IoListLogic.FindColumn(headers, "profinet name");

            // The IO list is hierarchical: a bus module's row (numeric ID, Profinet name or a
            // PA/PW type) is followed by the rows that hang under it — an AS-I gateway's slaves
            // among them. The context travels down so a slave knows its gateway's node.
            string moduleNode = "";
            for (int r = 1; r < rowCount; r++)
            {
                string idText = Cell(r, id).Trim().TrimStart('\'').Trim();
                string typeText = Cell(r, type).Trim();
                if (IsAllDigits(idText) &&
                    (typeText.Equals("PA", StringComparison.OrdinalIgnoreCase) ||
                     typeText.Equals("PW", StringComparison.OrdinalIgnoreCase) ||
                     Cell(r, pnName).Trim().Length > 0))
                    moduleNode = idText;

                string mnemonicText = Cell(r, mnemonic);
                if (mnemonicText.Trim().Length == 0) continue;

                rows.Add(new IoRowData
                {
                    CardSheet = name,
                    SourceRow = r + 1,
                    MnemonicCell = mnemonicText,
                    AddressCell = Cell(r, bit),
                    IdCell = Cell(r, id),
                    SlotCell = Cell(r, slot),
                    ModuleNode = moduleNode,
                    TypeCell = Cell(r, type),
                    Desc1A = Cell(r, desc1A),
                    Desc1B = Cell(r, desc1B),
                    Desc2A = Cell(r, desc2A),
                    Desc2B = Cell(r, desc2B),
                    FUnit = Cell(r, funit),
                    Location = Cell(r, location),
                    Device = Cell(r, device),
                    TsRef = Cell(r, tsRef),
                    ProfinetIp = Cell(r, pnIp),
                    ProfinetName = Cell(r, pnName),
                });
            }
        }
    }
}
