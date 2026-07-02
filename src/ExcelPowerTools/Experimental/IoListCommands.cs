using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.Experimental;
using ExcelPowerTools.Interop;
using Microsoft.Win32;

namespace ExcelPowerTools.Experimental
{
    /// <summary>
    /// EXPERIMENTAL — IO-list ↔ plant-CSV bridge (legacy IO-list format). Physical join key:
    /// node.offset.card.bit, where the CSV supplies Node/Offset/Bit + BOARD NUMBER (= card) and the
    /// IO list supplies "PC Board Rack.Slot.Card" + Bit. Helper columns are marked "//" in row 1 and
    /// are never saved into the CSV. To remove the whole feature: delete this folder,
    /// Core/Experimental, tests/Experimental, and the "IO List [exp]" ribbon group.
    /// </summary>
    public static class IoListCommands
    {
        private const string SettingsKey = @"Software\ExcelPowerTools\Experimental";

        // ---- 1. Build the //Address column on a plant CSV ------------------------------------------

        [ExcelCommand(Name = "PT_XBuildAddress", MenuName = "Power Tools", MenuText = "[exp] Build //Address column")]
        public static void BuildAddressColumn()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            var g = ReadSheet(sheet);
            if (g == null) { Dialogs.Info("The sheet is empty."); return; }

            var csv = LocateCsvColumns(g, requireBoard: false);
            if (csv == null) return;

            char dir = DetectDirection(app, sheet);

            var undo = UndoManager.NewEntry("Build //Address", sheet);
            using (new ExcelStateGuard(manualCalculation: true))
            {
                int col = FindOrCreateHelperColumn(sheet, g, "//Address", csv.HeaderSheetRow, undo);

                var values = new string[g.Rows - csv.HeaderSheetRow];
                for (int r = csv.HeaderSheetRow; r < g.Rows; r++)
                {
                    string node = g.Texts[r, csv.Node], off = g.Texts[r, csv.Offset], bit = g.Texts[r, csv.Bit];
                    values[r - csv.HeaderSheetRow] =
                        node.Length > 0 && off.Length > 0 && bit.Length > 0
                            ? IoListLogic.BuildAddress(dir, node, off, bit)
                            : "";
                }
                WriteColumn(sheet, col, csv.HeaderSheetRow + 1, values, undo);
            }
            UndoManager.Push(undo);

            Dialogs.Info($"//Address ready for {g.Rows - csv.HeaderSheetRow} row(s)" +
                         (dir == '\0' ? "\n(direction unknown — no I/O prefix; name the file/sheet *Inputs*/*Outputs*)" : $" (direction {dir})") + ".");
        }

        // ---- 2. Pull IO-list data into the plant CSV ------------------------------------------------

        private static readonly (string Label, string[] Fragments)[] PullFields =
        {
            ("Mnemonic", new[] { "mnemonic" }),
            ("ID", new[] { "id" }),
            ("Bit", new[] { "bit" }),
            ("Description", new[] { "description english language", "description english", "description" }),
            ("Functional unit", new[] { "functional unit" }),
            ("Location", new[] { "location" }),
            ("Device", new[] { "device" }),
            ("Type", new[] { "type" }),
            ("Drawing", new[] { "drawing" }),
            ("Sheet", new[] { "sheet" }),
        };

        [ExcelCommand(Name = "PT_XPullIoData", MenuName = "Power Tools", MenuText = "[exp] Get data from IO list")]
        public static void PullFromIoList()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            var g = ReadSheet(sheet);
            if (g == null) { Dialogs.Info("The sheet is empty."); return; }

            var csv = LocateCsvColumns(g, requireBoard: true);
            if (csv == null) return;

            string path = AskIoListPath();
            if (path == null) return;

            // Read the IO list (hidden, read-only) into key -> rows of field values.
            Dictionary<string, List<string[]>> ioMap;
            List<string> missingFields;
            using (new ExcelStateGuard(suppressEvents: true))
            {
                ioMap = ReadIoList(app, path, out missingFields);
            }
            if (ioMap == null) return;

            var undo = UndoManager.NewEntry("Get data from IO list", sheet);
            int matched = 0, withKey = 0;
            using (new ExcelStateGuard(manualCalculation: true))
            {
                int dataRows = g.Rows - csv.HeaderSheetRow;
                var keys = new string[dataRows];
                for (int r = csv.HeaderSheetRow; r < g.Rows; r++)
                {
                    string node = g.Texts[r, csv.Node], off = g.Texts[r, csv.Offset],
                           bit = g.Texts[r, csv.Bit], board = g.Texts[r, csv.Board];
                    if (node.Length > 0 && off.Length > 0 && bit.Length > 0 && board.Length > 0)
                    {
                        keys[r - csv.HeaderSheetRow] = IoListLogic.JoinKey(node, off, board, bit);
                        withKey++;
                        if (ioMap.ContainsKey(keys[r - csv.HeaderSheetRow])) matched++;
                    }
                }

                for (int f = 0; f < PullFields.Length; f++)
                {
                    if (missingFields.Contains(PullFields[f].Label)) continue;

                    int col = FindOrCreateHelperColumn(sheet, g, "//" + PullFields[f].Label, csv.HeaderSheetRow, undo);
                    var values = new string[dataRows];
                    for (int i = 0; i < dataRows; i++)
                        values[i] = keys[i] != null && ioMap.TryGetValue(keys[i], out var rows)
                            ? JoinField(rows, f)
                            : "";
                    WriteColumn(sheet, col, csv.HeaderSheetRow + 1, values, undo);
                }
            }
            UndoManager.Push(undo);

            Dialogs.Info($"IO-list data filled.\n\nRows with a complete node.offset.card.bit key: {withKey}\n" +
                         $"Matched in the IO list: {matched}" +
                         (missingFields.Count > 0 ? $"\nIO-list columns not found (skipped): {string.Join(", ", missingFields)}" : ""));
        }

        // ---- 3. Pull CSV mnemonics into the IO list -------------------------------------------------

        [ExcelCommand(Name = "PT_XPullMnemonics", MenuName = "Power Tools", MenuText = "[exp] Get mnemonics from CSVs")]
        public static void PullMnemonicsIntoIoList()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            var g = ReadSheet(sheet);
            if (g == null) { Dialogs.Info("The sheet is empty."); return; }

            int headerIdx = FindIoHeaderRow(g);
            if (headerIdx < 0) { Dialogs.Info("This doesn't look like an IO list (no 'Mnemonic' header found)."); return; }

            var headers = RowOf(g, headerIdx);
            int iBoard = IoListLogic.FindColumn(headers, "rack.slot.card", "pc board");
            int iBit = IoListLogic.FindColumn(headers, "bit");
            if (iBoard < 0 || iBit < 0)
            {
                Dialogs.Info("IO-list columns not found: " +
                             (iBoard < 0 ? "'PC Board Rack.Slot.Card' " : "") + (iBit < 0 ? "'Bit'" : ""));
                return;
            }

            string[] csvPaths = AskCsvPaths();
            if (csvPaths == null || csvPaths.Length == 0) return;

            var mnemonics = new Dictionary<string, List<string>>();
            var directions = new Dictionary<string, char>();
            var fileNotes = new List<string>();
            foreach (string path in csvPaths)
            {
                int n = LoadCsvMnemonics(path, mnemonics, directions);
                fileNotes.Add($"{Path.GetFileName(path)}: {(n < 0 ? "no Name/Node/Offset/Bit/Board columns" : n + " row(s)")}");
            }

            var undo = UndoManager.NewEntry("Get mnemonics from CSVs", sheet);
            int matched = 0;
            using (new ExcelStateGuard(manualCalculation: true))
            {
                int headerSheetRow = headerIdx + 1;
                int colAddr = FindOrCreateHelperColumn(sheet, g, "//Address", headerSheetRow, undo);
                int colMn = FindOrCreateHelperColumn(sheet, g, "//Mnemonic", headerSheetRow, undo);

                int dataRows = g.Rows - headerSheetRow;
                var addrValues = new string[dataRows];
                var mnValues = new string[dataRows];
                for (int r = headerSheetRow; r < g.Rows; r++)
                {
                    int i = r - headerSheetRow;
                    addrValues[i] = ""; mnValues[i] = "";

                    if (!IoListLogic.TryParseRackSlotCard(g.Texts[r, iBoard], out string node, out string off, out string card))
                        continue;
                    string bit = g.Texts[r, iBit];
                    if (bit.Length == 0) continue;

                    string key = IoListLogic.JoinKey(node, off, card, bit);
                    directions.TryGetValue(key, out char dir);
                    addrValues[i] = IoListLogic.BuildAddress(dir, node, off, bit);
                    mnValues[i] = IoListLogic.Joined(mnemonics, key);
                    if (mnValues[i].Length > 0) matched++;
                }
                WriteColumn(sheet, colAddr, headerSheetRow + 1, addrValues, undo);
                WriteColumn(sheet, colMn, headerSheetRow + 1, mnValues, undo);
            }
            UndoManager.Push(undo);

            Dialogs.Info("Mnemonics filled.\n\n" + string.Join("\n", fileNotes) +
                         $"\n\nIO-list rows matched: {matched}");
        }

        // ---- CSV-side helpers -----------------------------------------------------------------------

        private sealed class SheetGrid
        {
            public string[,] Texts;   // [Rows, Cols], A1-anchored
            public int Rows, Cols;
        }

        private sealed class CsvColumns
        {
            public int HeaderSheetRow;   // 1-based sheet row of the header (== header index + 1)
            public int Node, Offset, Bit, Board;
        }

        private static SheetGrid ReadSheet(dynamic sheet)
        {
            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);
            if (lastRow < 1) return null;

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            object raw = range.Value2;

            var g = new SheetGrid { Rows = (int)lastRow, Cols = (int)lastCol };
            g.Texts = new string[g.Rows, g.Cols];
            if (raw is object[,] grid)
            {
                int rb = grid.GetLowerBound(0), cb = grid.GetLowerBound(1);
                for (int r = 0; r < g.Rows; r++)
                    for (int c = 0; c < g.Cols; c++)
                        g.Texts[r, c] = CsvSheet.CellText(grid[rb + r, cb + c], sheet, r + 1, c + 1);
            }
            else
            {
                g.Texts[0, 0] = CsvSheet.CellText(raw, sheet, 1, 1);
            }
            return g;
        }

        private static string[] RowOf(SheetGrid g, int rowIdx)
        {
            var row = new string[g.Cols];
            for (int c = 0; c < g.Cols; c++) row[c] = g.Texts[rowIdx, c];
            return row;
        }

        private static CsvColumns LocateCsvColumns(SheetGrid g, bool requireBoard)
        {
            var rowsList = new List<string[]>(g.Rows);
            for (int r = 0; r < g.Rows; r++) rowsList.Add(RowOf(g, r));

            int headerIdx = CsvDetect.DetectHeaderRow(rowsList, "#");
            if (headerIdx < 0)
            {
                Dialogs.Info("No header row detected on this sheet.");
                return null;
            }

            var headers = rowsList[headerIdx];
            var c = new CsvColumns
            {
                HeaderSheetRow = headerIdx + 1,
                Node = IoListLogic.FindColumn(headers, "node"),
                Offset = IoListLogic.FindColumn(headers, "offset"),
                Bit = IoListLogic.FindColumn(headers, "bit"),
                Board = IoListLogic.FindColumn(headers, "board number", "board"),
            };

            var missing = new List<string>();
            if (c.Node < 0) missing.Add("Node");
            if (c.Offset < 0) missing.Add("Offset");
            if (c.Bit < 0) missing.Add("Bit");
            if (requireBoard && c.Board < 0) missing.Add("BOARD NUMBER");
            if (missing.Count > 0)
            {
                Dialogs.Info("CSV columns not found in the header: " + string.Join(", ", missing));
                return null;
            }
            return c;
        }

        private static char DetectDirection(dynamic app, dynamic sheet)
        {
            try
            {
                var session = CsvSession.Get((string)app.ActiveWorkbook.Name);
                if (session?.SourcePath != null)
                {
                    char d = IoListLogic.DirectionFromName(Path.GetFileName(session.SourcePath));
                    if (d != '\0') return d;
                }
            }
            catch { }
            try { return IoListLogic.DirectionFromName((string)sheet.Name); } catch { return '\0'; }
        }

        // ---- helper-column plumbing ------------------------------------------------------------------

        private static int FindOrCreateHelperColumn(dynamic sheet, SheetGrid g, string marker, int headerSheetRow, UndoEntry undo)
        {
            for (int c = 0; c < g.Cols; c++)
                if (g.Texts[0, c] == marker) return c + 1;

            int col = g.Cols + 1;
            g.Cols = col;   // subsequent creations go one further right

            try { sheet.Columns(col).NumberFormat = "@"; } catch { }
            UndoManager.CaptureCell(undo, sheet, 1, col);
            sheet.Cells(1, col).Value = marker;
            if (headerSheetRow > 1)
            {
                UndoManager.CaptureCell(undo, sheet, headerSheetRow, col);
                sheet.Cells(headerSheetRow, col).Value = marker;
            }
            return col;
        }

        private static void WriteColumn(dynamic sheet, int col, int firstSheetRow, string[] values, UndoEntry undo)
        {
            if (values.Length == 0) return;
            dynamic rng = sheet.Range(sheet.Cells(firstSheetRow, col), sheet.Cells(firstSheetRow + values.Length - 1, col));
            UndoManager.CaptureBlock(undo, rng);

            var data = new object[values.Length, 1];
            for (int i = 0; i < values.Length; i++) data[i, 0] = values[i];
            rng.NumberFormat = "@";
            rng.Value = data;
        }

        // ---- IO-list reading -------------------------------------------------------------------------

        private static int FindIoHeaderRow(SheetGrid g)
        {
            int limit = Math.Min(g.Rows, 20);
            for (int r = 0; r < limit; r++)
                for (int c = 0; c < g.Cols; c++)
                    if (IoListLogic.NormalizeHeader(g.Texts[r, c]) == "mnemonic") return r;
            return -1;
        }

        private static Dictionary<string, List<string[]>> ReadIoList(dynamic app, string path, out List<string> missingFields)
        {
            missingFields = new List<string>();
            dynamic wb = null;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }

            try
            {
                wb = app.Workbooks.Open(path, 0, true);   // UpdateLinks:=0, ReadOnly:=true

                foreach (dynamic ws in wb.Worksheets)
                {
                    var g = ReadSheet(ws);
                    if (g == null) continue;
                    int headerIdx = FindIoHeaderRow(g);
                    if (headerIdx < 0) continue;

                    var headers = RowOf(g, headerIdx);
                    int iBoard = IoListLogic.FindColumn(headers, "rack.slot.card", "pc board");
                    int iBit = IoListLogic.FindColumn(headers, "bit");
                    if (iBoard < 0 || iBit < 0) continue;

                    var fieldCols = new int[PullFields.Length];
                    for (int f = 0; f < PullFields.Length; f++)
                    {
                        fieldCols[f] = IoListLogic.FindColumn(headers, PullFields[f].Fragments);
                        if (fieldCols[f] < 0) missingFields.Add(PullFields[f].Label);
                    }

                    var map = new Dictionary<string, List<string[]>>();
                    for (int r = headerIdx + 1; r < g.Rows; r++)
                    {
                        if (!IoListLogic.TryParseRackSlotCard(g.Texts[r, iBoard], out string node, out string off, out string card))
                            continue;
                        string bit = g.Texts[r, iBit];
                        if (bit.Length == 0) continue;

                        string key = IoListLogic.JoinKey(node, off, card, bit);
                        var fields = new string[PullFields.Length];
                        for (int f = 0; f < PullFields.Length; f++)
                            fields[f] = fieldCols[f] >= 0 ? g.Texts[r, fieldCols[f]] : "";

                        if (!map.TryGetValue(key, out var list)) map[key] = list = new List<string[]>();
                        list.Add(fields);
                    }

                    return map;   // first sheet with a Mnemonic header wins
                }

                Dialogs.Info("No sheet with a 'Mnemonic' header found in:\n" + path);
                return null;
            }
            catch (Exception ex)
            {
                Dialogs.Info("Couldn't read the IO list:\n" + ex.Message);
                return null;
            }
            finally
            {
                try { wb?.Close(false); } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
            }
        }

        private static string JoinField(List<string[]> rows, int fieldIndex)
        {
            if (rows.Count == 1) return rows[0][fieldIndex];
            var seen = new List<string>();
            foreach (var row in rows)
            {
                string v = row[fieldIndex];
                if (v.Length > 0 && !seen.Contains(v)) seen.Add(v);
            }
            return string.Join("/", seen);
        }

        // ---- CSV file reading (headless, via Core) ----------------------------------------------------

        private static int LoadCsvMnemonics(string path, Dictionary<string, List<string>> mnemonics, Dictionary<string, char> directions)
        {
            try
            {
                var doc = CsvFile.Load(File.ReadAllBytes(path));
                var rows = new List<string[]>(doc.Records.Count);
                foreach (var rec in doc.Records) rows.Add(rec.Fields);

                int headerIdx = CsvDetect.DetectHeaderRow(rows, "#");
                if (headerIdx < 0) return -1;

                var headers = rows[headerIdx];
                int iName = IoListLogic.FindColumn(headers, "input name", "output name", "name");
                int iNode = IoListLogic.FindColumn(headers, "node");
                int iOff = IoListLogic.FindColumn(headers, "offset");
                int iBit = IoListLogic.FindColumn(headers, "bit");
                int iBoard = IoListLogic.FindColumn(headers, "board number", "board");
                if (iName < 0 || iNode < 0 || iOff < 0 || iBit < 0 || iBoard < 0) return -1;

                char dir = IoListLogic.DirectionFromName(Path.GetFileName(path));
                int count = 0;
                for (int r = headerIdx + 1; r < rows.Count; r++)
                {
                    string[] f = rows[r];
                    if (f.Length <= iName || f.Length <= iNode || f.Length <= iOff || f.Length <= iBit || f.Length <= iBoard)
                        continue;

                    string name = f[iName].Trim();
                    string node = f[iNode], off = f[iOff], bit = f[iBit], board = f[iBoard];
                    if (name.Length == 0 || node.Trim().Length == 0 || off.Trim().Length == 0
                        || bit.Trim().Length == 0 || board.Trim().Length == 0)
                        continue;

                    string key = IoListLogic.JoinKey(node, off, board, bit);
                    IoListLogic.AddToMap(mnemonics, key, name);
                    if (dir != '\0' && !directions.ContainsKey(key)) directions[key] = dir;
                    count++;
                }
                return count;
            }
            catch
            {
                return -1;
            }
        }

        // ---- dialogs & settings ------------------------------------------------------------------------

        private static string AskIoListPath()
        {
            string last = GetSetting("IoListPath");
            using (var d = new OpenFileDialog
            {
                Filter = "Excel IO list|*.xlsx;*.xlsm;*.xls|All files|*.*",
                CheckFileExists = true,
                FileName = last ?? "",
            })
            {
                if (!string.IsNullOrEmpty(last))
                    try { d.InitialDirectory = Path.GetDirectoryName(last); } catch { }

                if (d.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return null;
                SetSetting("IoListPath", d.FileName);
                return d.FileName;
            }
        }

        private static string[] AskCsvPaths()
        {
            using (var d = new OpenFileDialog
            {
                Filter = "Plant CSVs|*.csv|All files|*.*",
                CheckFileExists = true,
                Multiselect = true,
            })
            {
                string lastDir = GetSetting("CsvDir");
                if (!string.IsNullOrEmpty(lastDir)) d.InitialDirectory = lastDir;

                if (d.ShowDialog(ExcelWindow.Current) != DialogResult.OK) return null;
                try { SetSetting("CsvDir", Path.GetDirectoryName(d.FileNames[0])); } catch { }
                return d.FileNames;
            }
        }

        private static string GetSetting(string name)
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(SettingsKey))
                return k?.GetValue(name) as string;
        }

        private static void SetSetting(string name, string value)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(SettingsKey))
                k?.SetValue(name, value ?? "");
        }
    }
}
