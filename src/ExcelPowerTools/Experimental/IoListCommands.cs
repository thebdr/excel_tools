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
    /// EXPERIMENTAL — IO-list ↔ plant-CSV bridge. The IO list holds one worksheet per card, named
    /// "R{rack}.S{slot}.C{card}" (optionally with a label suffix, "R1.S2.C1 - PC2"); the card comes from
    /// that name and the address "I.node.offset.bit" from the "Bit" column. The CSV supplies
    /// Node/Offset/Bit + Card, so the physical join key is node.offset.card.bit per direction. A card
    /// number is only unique WITHIN a slot, so an index always covers exactly one controller.
    ///
    /// Both pulls match a row two ways — by address and by mnemonic NAME — and report the pair
    /// "//Mnemonic" + "//Mnemonic address", so a signal that moved to a different address is visible
    /// instead of just missing. "//Mnemonic match" flags rows whose two sides disagree and drives a
    /// conditional format on the mnemonic cells.
    ///
    /// Bus-diagnostic CSV rows (Type 11) have no bit address and join on the node alone, against the
    /// module rows (empty "Bit", node in "ID"). Helper columns are marked "//" in row 1 and are never
    /// saved into the CSV. To remove the whole feature: delete this folder, Core/Experimental,
    /// tests/Experimental, and the "IO List [exp]" ribbon group.
    /// </summary>
    public static class IoListCommands
    {
        private const string SettingsKey = @"Software\ExcelPowerTools\Experimental";

        // Written in this order, so the verdict leads and sits next to the column it judges.
        private const string MatchedColumn = "//Mnemonics matched";
        private const string ByAddressColumn = "//Mnemonics by address";
        private const string ByNameColumn = "//Mnemonics by name";

        private const int XlExpression = 2;
        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);   // Excel OLE colour (BGR)
        private static readonly int MismatchFill = Rgb(255, 199, 206);
        private static readonly int MismatchText = Rgb(156, 0, 6);

        // Header fragments for the CSV side. Inputs and Outputs are exported with different schemas
        // (and different delimiters), so every lookup carries both spellings.
        private static readonly string[] CsvNameHeaders = { "signal id", "input name", "output name", "name" };
        private static readonly string[] CsvBoardHeaders = { "card", "board number", "board" };
        private static readonly string[] CsvTypeHeaders = { "type", "input type", "output type" };

        // ---- 1. Build the //Address column on a plant CSV ------------------------------------------

        [ExcelCommand(Name = "PT_XBuildAddress", MenuName = "Power Tools", MenuText = "[exp] Build //Address column")]
        public static void BuildAddressColumn()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            SheetGrid g = ReadSheet(sheet);
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

        /// <summary>
        /// Fields lifted out of the IO list. <c>Extra</c> is an optional second column appended after a
        /// space — the description is split into a permanent and a variable part in the current format.
        /// </summary>
        private static readonly (string Label, string[] Fragments, string[] Extra)[] PullFields =
        {
            ("Mnemonic", new[] { "mnemonic" }, null),
            ("ID", new[] { "id" }, null),
            ("Bit", new[] { "bit" }, null),
            ("Description", new[] { "Description language 1 (1° part", "description english language" },
                            new[] { "Description language 1 (2° part" }),
            ("Functional unit", new[] { "functional unit" }, null),
            ("Location", new[] { "location" }, null),
            ("Device", new[] { "device" }, null),
            ("Type", new[] { "type" }, null),
            ("Drawing", new[] { "drawing name", "drawing" }, null),
            ("Sheet", new[] { "sheet" }, null),
        };

        private const int MnemonicField = 0;   // index into PullFields
        private const int BitField = 2;

        [ExcelCommand(Name = "PT_XPullIoData", MenuName = "Power Tools", MenuText = "[exp] Get data from IO list")]
        public static void PullFromIoList()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            SheetGrid g = ReadSheet(sheet);
            if (g == null) { Dialogs.Info("The sheet is empty."); return; }

            var csv = LocateCsvColumns(g, requireBoard: true);
            if (csv == null) return;

            string path = AskIoListPath();
            if (path == null) return;

            IoIndex index;
            using (new ExcelStateGuard(suppressEvents: true))
            {
                index = ReadIoList(app, path);
            }
            if (index == null) return;

            char dir = DetectDirection(app, sheet);

            var undo = UndoManager.NewEntry("Get data from IO list", sheet);
            int matchedByAddress = 0, matchedByNode = 0, matchedByNameOnly = 0,
                withKey = 0, busRows = 0, virtualSkipped = 0, mismatches = 0, nameAbsent = 0;
            using (new ExcelStateGuard(manualCalculation: true))
            {
                int dataRows = g.Rows - csv.HeaderSheetRow;
                var hits = new List<IoRow>[dataRows];        // union, for the descriptive fields
                var byAddress = new string[dataRows];
                var byName = new string[dataRows];
                var matched = new string[dataRows];

                for (int i = 0; i < dataRows; i++) byAddress[i] = byName[i] = matched[i] = "";

                for (int r = csv.HeaderSheetRow; r < g.Rows; r++)
                {
                    int i = r - csv.HeaderSheetRow;
                    string type = csv.Type >= 0 ? g.Texts[r, csv.Type] : "";

                    if (IoListLogic.IsVirtualType(type))
                    {
                        virtualSkipped++;   // types 5/6: virtual signals — address is not hardware I/O
                        continue;
                    }

                    string node = g.Texts[r, csv.Node], board = g.Texts[r, csv.Board];
                    if (node.Length == 0 || board.Length == 0) continue;

                    var addressRows = new List<IoRow>();
                    bool positional;

                    if (IoListLogic.IsBusDiagnosticType(type))
                    {
                        // Type 11 — bus diagnostics. No bit address: the node alone identifies the device.
                        busRows++;
                        positional = Collect(addressRows, index.ByNode, IoListLogic.NodeKey(board, node));
                        if (positional) matchedByNode++;
                    }
                    else
                    {
                        string off = g.Texts[r, csv.Offset], bit = g.Texts[r, csv.Bit];
                        if (off.Length == 0 || bit.Length == 0) continue;

                        withKey++;
                        positional = false;
                        if (dir != 'O') positional |= Collect(addressRows, index.ByAddress, IoListLogic.JoinKey('I', node, off, board, bit));
                        if (dir != 'I') positional |= Collect(addressRows, index.ByAddress, IoListLogic.JoinKey('O', node, off, board, bit));
                        if (positional) matchedByAddress++;
                    }

                    // ...and by NAME, so a signal that moved to another address is found where it now is.
                    // Only rows the address did NOT already produce: otherwise every agreeing row would
                    // repeat itself here.
                    string ownName = csv.Name >= 0 ? g.Texts[r, csv.Name] : "";
                    var ownKeys = IoListLogic.MnemonicKeys(ownName);
                    var nameRows = new List<IoRow>();
                    foreach (string key in ownKeys)
                        Collect(nameRows, index.ByMnemonic, key);
                    nameRows.RemoveAll(row => addressRows.Contains(row));

                    if (nameRows.Count > 0 && !positional) matchedByNameOnly++;

                    byAddress[i] = JoinMnemonics(addressRows);
                    byName[i] = JoinMnemonicsAtAddress(nameRows, ownKeys);

                    if (csv.Name >= 0)
                    {
                        // The IO list is the reference, so ITS mnemonics set the total — here that is the
                        // address-matched side.
                        matched[i] = IoListLogic.MatchRatio(ownName, byAddress[i], byAddress[i]);
                        if (matched[i].Length > 0)
                        {
                            mismatches++;
                            if (!IoListLogic.MnemonicSetCovers(byAddress[i], ownName)) nameAbsent++;
                        }
                    }

                    var union = new List<IoRow>(addressRows);
                    union.AddRange(nameRows);
                    if (union.Count > 0) hits[i] = union;
                }

                int colMatched = -1;
                if (csv.Name >= 0)
                {
                    colMatched = FindOrCreateHelperColumn(sheet, g, MatchedColumn, csv.HeaderSheetRow, undo);
                    WriteColumn(sheet, colMatched, csv.HeaderSheetRow + 1, matched, undo);
                }

                int colByAddress = FindOrCreateHelperColumn(sheet, g, ByAddressColumn, csv.HeaderSheetRow, undo);
                WriteColumn(sheet, colByAddress, csv.HeaderSheetRow + 1, byAddress, undo);

                int colByName = FindOrCreateHelperColumn(sheet, g, ByNameColumn, csv.HeaderSheetRow, undo);
                WriteColumn(sheet, colByName, csv.HeaderSheetRow + 1, byName, undo);

                for (int f = 0; f < PullFields.Length; f++)
                {
                    if (f == MnemonicField) continue;   // superseded by the two mnemonic columns above
                    int col = FindOrCreateHelperColumn(sheet, g, "//" + PullFields[f].Label, csv.HeaderSheetRow, undo);
                    var values = new string[dataRows];
                    for (int i = 0; i < dataRows; i++)
                        values[i] = hits[i] != null ? JoinField(hits[i], f) : "";
                    WriteColumn(sheet, col, csv.HeaderSheetRow + 1, values, undo);
                }

                if (colMatched > 0)
                    ApplyMismatchHighlight(sheet, colMatched, colByAddress, csv.HeaderSheetRow + 1, g.Rows);
            }
            UndoManager.Push(undo);

            Dialogs.Info($"IO-list data filled from controller {index.Controller}, " +
                         $"{index.SheetsUsed} card sheet(s) ({string.Join(", ", index.SheetNames)}), " +
                         $"{index.RowsIndexed} indexed row(s).\n\n" +
                         $"CSV rows with a complete node.offset.card.bit key: {withKey}\n" +
                         $"Matched by address: {matchedByAddress}\n" +
                         $"Bus-diagnostic rows (type 11): {busRows}, matched by node: {matchedByNode}\n" +
                         $"Found by mnemonic at another address only: {matchedByNameOnly}\n" +
                         $"Virtual rows skipped (type 5/6): {virtualSkipped}\n" +
                         (csv.Name >= 0
                            ? $"Rows carrying a matched/total ratio (highlighted): {mismatches}\n" +
                              $"  · signal name NOT among the address-matched mnemonics: {nameAbsent}\n" +
                              $"  · name present, sharing its address with others: {mismatches - nameAbsent}"
                            : "No name column found — the matched ratio was skipped.") +
                         (dir == '\0' ? "\n(direction unknown — matched both I and O)" : $"\n(direction {dir})"));
        }

        /// <summary>Appends a key's rows to <paramref name="into"/>, skipping ones already there.</summary>
        private static bool Collect(List<IoRow> into, Dictionary<string, List<IoRow>> map, string key)
        {
            if (key == null || !map.TryGetValue(key, out var found)) return false;
            foreach (IoRow row in found)
                if (!into.Contains(row)) into.Add(row);
            return true;
        }

        /// <summary>Every distinct mnemonic across the rows, "/"-joined. No address: it is the row's own.</summary>
        private static string JoinMnemonics(List<IoRow> rows)
        {
            var names = new List<string>();
            foreach (IoRow row in rows)
                foreach (string part in row.Fields[MnemonicField].Split('/'))
                {
                    string mn = part.Trim();
                    if (mn.Length > 0 && !names.Contains(mn)) names.Add(mn);
                }
            return string.Join("/", names);
        }

        /// <summary>
        /// The rows' mnemonics tagged with where they live ("LP100301@O.90.0.1"), restricted to the ones
        /// that actually answer <paramref name="wantedKeys"/> — a companion mnemonic sharing the row is
        /// not what the name lookup was asking about.
        /// </summary>
        private static string JoinMnemonicsAtAddress(List<IoRow> rows, List<string> wantedKeys)
        {
            var entries = new List<string>();
            foreach (IoRow row in rows)
                foreach (string part in row.Fields[MnemonicField].Split('/'))
                {
                    string mn = part.Trim();
                    if (mn.Length == 0 || !wantedKeys.Contains(IoListLogic.MnemonicKey(mn))) continue;
                    string entry = IoListLogic.AtAddress(mn, row.Address);
                    if (!entries.Contains(entry)) entries.Add(entry);
                }
            return string.Join("/", entries);
        }

        // ---- 3. Pull CSV mnemonics into the IO list -------------------------------------------------

        [ExcelCommand(Name = "PT_XPullMnemonics", MenuName = "Power Tools", MenuText = "[exp] Get mnemonics from CSVs")]
        public static void PullMnemonicsIntoIoList()
        {
            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            if (sheet == null) return;

            string sheetName;
            try { sheetName = (string)sheet.Name; } catch { sheetName = ""; }

            if (!IoListLogic.TryParseCardSheetName(sheetName, out string rack, out string slot, out string card))
            {
                Dialogs.Info($"\"{sheetName}\" is not a card sheet.\n\n" +
                             "Run this on a sheet named R{rack}.S{slot}.C{card} (a label may follow, " +
                             "e.g. \"R1.S2.C1 - PC2\") — the card number is what joins to the CSV.");
                return;
            }
            string controller = IoListLogic.ControllerKey(rack, slot);

            SheetGrid g = ReadSheet(sheet);
            if (g == null) { Dialogs.Info("The sheet is empty."); return; }

            int headerIdx = FindIoHeaderRow(g, out int iAddr, out int iMnemonic, out int iId);
            if (headerIdx < 0)
            {
                Dialogs.Info("IO-list layout not recognized: no header row with both a \"Bit\" and a " +
                             "\"Mnemonic\" column was found in the first 20 rows.");
                return;
            }

            string[] csvPaths = AskCsvPaths();
            if (csvPaths == null || csvPaths.Length == 0) return;

            var index = new CsvIndex();
            var fileNotes = new List<string>();
            foreach (string path in csvPaths)
            {
                int n = LoadCsvSignals(path, index, out bool dirUnknown, out int virtualSkipped, out int busRows);
                fileNotes.Add($"{Path.GetFileName(path)}: {(n < 0 ? "no Name/Node/Offset/Bit/Card columns" : n + " row(s)")}" +
                              (busRows > 0 ? $", {busRows} bus-diagnostic (type 11) by node" : "") +
                              (virtualSkipped > 0 ? $", {virtualSkipped} virtual (type 5/6) skipped" : "") +
                              (dirUnknown ? " (direction unknown — indexed as both I and O)" : ""));
            }

            var undo = UndoManager.NewEntry("Get mnemonics from CSVs", sheet);
            int matchedByAddress = 0, matchedByNode = 0, matchedByNameOnly = 0, mismatches = 0, nameAbsent = 0;
            using (new ExcelStateGuard(manualCalculation: true))
            {
                int headerSheetRow = headerIdx + 1;
                int dataRows = g.Rows - headerSheetRow;
                var byAddress = new string[dataRows];
                var byName = new string[dataRows];
                var matched = new string[dataRows];

                for (int r = headerSheetRow; r < g.Rows; r++)
                {
                    int i = r - headerSheetRow;
                    byAddress[i] = byName[i] = matched[i] = "";

                    var addressSignals = new List<CsvSignal>();
                    bool positional = false;

                    if (IoListLogic.TryParseAddress(g.Texts[r, iAddr], out char dir, out string node, out string off, out string bit))
                    {
                        positional = Collect(addressSignals, index.ByAddress, IoListLogic.JoinKey(dir, node, off, card, bit));
                        if (positional) matchedByAddress++;
                    }
                    else if (iId >= 0 && g.Texts[r, iId].Length > 0)
                    {
                        // Module row: no bit address, the node lives in ID — the bus-diagnostic join.
                        positional = Collect(addressSignals, index.ByNode, IoListLogic.NodeKey(card, g.Texts[r, iId]));
                        if (positional) matchedByNode++;
                    }

                    // ...and by NAME, so a mnemonic that moved to another address is found where it now is.
                    // Excluding what the address already produced, so an agreeing row does not repeat itself.
                    string ownMnemonic = g.Texts[r, iMnemonic];
                    var nameSignals = new List<CsvSignal>();
                    foreach (string key in IoListLogic.MnemonicKeys(ownMnemonic))
                        Collect(nameSignals, index.ByMnemonic, key);
                    nameSignals.RemoveAll(s => addressSignals.Contains(s));

                    if (nameSignals.Count > 0 && !positional) matchedByNameOnly++;

                    var names = new List<string>();
                    foreach (CsvSignal s in addressSignals)
                        if (!names.Contains(s.Name)) names.Add(s.Name);
                    byAddress[i] = string.Join("/", names);

                    var tagged = new List<string>();
                    foreach (CsvSignal s in nameSignals)
                    {
                        string entry = IoListLogic.AtAddress(s.Name, s.Address);
                        if (!tagged.Contains(entry)) tagged.Add(entry);
                    }
                    byName[i] = string.Join("/", tagged);

                    // Here the SHEET is the IO list, so its own Mnemonic cell sets the total.
                    matched[i] = IoListLogic.MatchRatio(ownMnemonic, byAddress[i], ownMnemonic);
                    if (matched[i].Length > 0)
                    {
                        mismatches++;
                        if (!IoListLogic.MnemonicSetCovers(byAddress[i], ownMnemonic)) nameAbsent++;
                    }
                }

                int colMatched = FindOrCreateHelperColumn(sheet, g, MatchedColumn, headerSheetRow, undo);
                WriteColumn(sheet, colMatched, headerSheetRow + 1, matched, undo);

                int colByAddress = FindOrCreateHelperColumn(sheet, g, ByAddressColumn, headerSheetRow, undo);
                WriteColumn(sheet, colByAddress, headerSheetRow + 1, byAddress, undo);

                int colByName = FindOrCreateHelperColumn(sheet, g, ByNameColumn, headerSheetRow, undo);
                WriteColumn(sheet, colByName, headerSheetRow + 1, byName, undo);

                ApplyMismatchHighlight(sheet, colMatched, colByAddress, headerSheetRow + 1, g.Rows);
            }
            UndoManager.Push(undo);

            // The card number is scoped to the slot, so the CSVs must be the ones for THIS controller.
            Dialogs.Info($"Mnemonics filled for {controller} card {card} (sheet \"{sheetName}\").\n\n" +
                         string.Join("\n", fileNotes) +
                         $"\n\nIO-list rows matched by address: {matchedByAddress}" +
                         $"\nModule rows matched by node: {matchedByNode}" +
                         $"\nFound by mnemonic at another address only: {matchedByNameOnly}" +
                         $"\nRows carrying a matched/total ratio (highlighted): {mismatches}" +
                         $"\n  · row mnemonic NOT among the CSV names there: {nameAbsent}" +
                         $"\n  · present, sharing the address with others: {mismatches - nameAbsent}" +
                         "\n\nA card number is only unique within its slot — check that the CSVs above are " +
                         $"the ones exported for controller {controller}.");
        }

        /// <summary>Appends a key's signals to <paramref name="into"/>, skipping ones already there.</summary>
        private static bool Collect(List<CsvSignal> into, Dictionary<string, List<CsvSignal>> map, string key)
        {
            if (key == null || !map.TryGetValue(key, out var found)) return false;
            foreach (CsvSignal s in found)
                if (!into.Contains(s)) into.Add(s);
            return true;
        }

        // ---- mismatch highlighting --------------------------------------------------------------------

        /// <summary>
        /// Paints the verdict and the address-matched mnemonics wherever a ratio was written — the ratio
        /// column is blank on agreement, so the rule is simply "not empty". A real conditional format
        /// keyed on that column, so it follows the data: an Undo restores the values and the highlight
        /// goes with them. Re-running replaces only this rule and leaves the CSV view styling alone.
        /// </summary>
        private static void ApplyMismatchHighlight(dynamic sheet, int matchedCol, int throughCol, int firstDataRow, int lastRow)
        {
            if (lastRow < firstDataRow) return;

            try
            {
                int left = Math.Min(matchedCol, throughCol), right = Math.Max(matchedCol, throughCol);
                dynamic target = sheet.Range(sheet.Cells(firstDataRow, left), sheet.Cells(lastRow, right));
                string flagRef = (string)sheet.Cells(firstDataRow, matchedCol).Address(false, true);   // "$AB2"
                string formula = "=" + flagRef + "<>\"\"";

                // Drop only OUR rule from a previous run; other rules on these cells stay untouched.
                for (int i = (int)target.FormatConditions.Count; i >= 1; i--)
                {
                    dynamic existing = target.FormatConditions(i);
                    string f1 = null;
                    try { f1 = (string)existing.Formula1; } catch { }
                    if (f1 == formula) existing.Delete();
                }

                dynamic rule = target.FormatConditions.Add(XlExpression, Type.Missing, formula);
                rule.Interior.Color = MismatchFill;
                rule.Font.Color = MismatchText;
                rule.StopIfTrue = false;   // Add() defaults it to true, which would swallow the CSV zebra
            }
            catch { }
        }

        // ---- CSV-side helpers -----------------------------------------------------------------------

        private sealed class SheetGrid
        {
            public string[,] Texts;   // [Rows, Cols], A1-anchored snapshot — dimensions never change
            public int Rows, Cols;
            public int NextFreeColumn;   // 1-based sheet column where the next new helper column goes
            public readonly Dictionary<string, int> CreatedHelpers = new Dictionary<string, int>();
        }

        private sealed class CsvColumns
        {
            public int HeaderSheetRow;   // 1-based sheet row of the header (== header index + 1)
            public int Node, Offset, Bit, Board;
            public int Name;             // -1 when absent (no mismatch flagging)
            public int Type;             // -1 when absent (no virtual / bus-diagnostic classification)
        }

        private static SheetGrid ReadSheet(dynamic sheet)
        {
            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);
            if (lastRow < 1) return null;

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            object raw = range.Value2;

            var g = new SheetGrid { Rows = (int)lastRow, Cols = (int)lastCol };
            g.NextFreeColumn = g.Cols + 1;
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
                Board = IoListLogic.FindColumn(headers, CsvBoardHeaders),
                Name = IoListLogic.FindColumn(headers, CsvNameHeaders),
                Type = IoListLogic.FindColumn(headers, CsvTypeHeaders),
            };

            var missing = new List<string>();
            if (c.Node < 0) missing.Add("Node");
            if (c.Offset < 0) missing.Add("Offset");
            if (c.Bit < 0) missing.Add("Bit");
            if (requireBoard && c.Board < 0) missing.Add("Card / Board number");
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
                var session = CsvSession.Get(app.ActiveSheet);
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
            // Created earlier in THIS run (not part of the snapshot)?
            if (g.CreatedHelpers.TryGetValue(marker, out int created)) return created;

            // Already on the sheet when it was snapshotted?
            for (int c = 0; c < g.Cols; c++)
                if (g.Texts[0, c] == marker) return c + 1;

            int col = g.NextFreeColumn++;
            g.CreatedHelpers[marker] = col;

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

        /// <summary>One indexed IO-list row: where it sits, and the values pulled from it.</summary>
        private sealed class IoRow
        {
            public string Address;      // "I.90.0.0" for a signal row, "#236" for a module row
            public string[] Fields;     // PullFields values; Fields[MnemonicField] may hold "A/B"
        }

        private sealed class IoIndex
        {
            /// <summary>DIRECTION.node.offset.card.bit → signal rows, one controller only.</summary>
            public readonly Dictionary<string, List<IoRow>> ByAddress = new Dictionary<string, List<IoRow>>();
            /// <summary>card#node → module rows (for bus-diagnostic CSV signals).</summary>
            public readonly Dictionary<string, List<IoRow>> ByNode = new Dictionary<string, List<IoRow>>();
            /// <summary>mnemonic key → rows carrying it, so a moved signal is found by name.</summary>
            public readonly Dictionary<string, List<IoRow>> ByMnemonic = new Dictionary<string, List<IoRow>>();
            public readonly List<string> SheetNames = new List<string>();
            public string Controller;           // "R1.S2" — the slot these sheets belong to
            public int SheetsUsed => SheetNames.Count;
            public int RowsIndexed;
        }

        /// <summary>A card sheet found in the workbook, before any of it is read.</summary>
        private sealed class CardSheet
        {
            public dynamic Worksheet;
            public string Name, Controller, Card;
        }

        /// <summary>
        /// Finds the header row — the first row (within 20) carrying BOTH a "Bit" and a "Mnemonic" column —
        /// and reports those columns plus "ID". Returns the row index, or -1.
        /// </summary>
        private static int FindIoHeaderRow(SheetGrid g, out int bitCol, out int mnemonicCol, out int idCol)
        {
            bitCol = mnemonicCol = idCol = -1;

            int limit = Math.Min(g.Rows, 20);
            for (int r = 0; r < limit; r++)
            {
                var headers = RowOf(g, r);
                int bit = IoListLogic.FindColumn(headers, "bit");
                int mn = IoListLogic.FindColumn(headers, "mnemonic");
                if (bit < 0 || mn < 0) continue;

                bitCol = bit;
                mnemonicCol = mn;
                idCol = IoListLogic.FindColumn(headers, "id");
                return r;
            }
            return -1;
        }

        private static IoIndex ReadIoList(dynamic app, string path)
        {
            var index = new IoIndex();

            dynamic wb = null;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }

            try
            {
                wb = app.Workbooks.Open(path, 0, true);   // UpdateLinks:=0, ReadOnly:=true

                // One sheet per card: the sheet NAME carries rack.slot.card. Everything else is ignored.
                var cardSheets = new List<CardSheet>();
                foreach (dynamic ws in wb.Worksheets)
                {
                    string name;
                    try { name = (string)ws.Name; } catch { continue; }
                    if (!IoListLogic.TryParseCardSheetName(name, out string rack, out string slot, out string card)) continue;

                    cardSheets.Add(new CardSheet
                    {
                        Worksheet = ws,
                        Name = name,
                        Controller = IoListLogic.ControllerKey(rack, slot),
                        Card = card,
                    });
                }

                // A card number is unique only within a slot, and the CSV carries no slot — so the index
                // must cover exactly one controller, otherwise "C1" from two slots would collide.
                string controller = ChooseController(cardSheets, path);
                if (controller == null) return null;
                index.Controller = controller;

                foreach (var cs in cardSheets)
                {
                    if (cs.Controller != controller) continue;
                    dynamic ws = cs.Worksheet;
                    string card = cs.Card;

                    SheetGrid g = (SheetGrid)ReadSheet(ws);
                    if (g == null) continue;

                    int headerIdx = FindIoHeaderRow(g, out int iAddr, out _, out int iId);
                    if (headerIdx < 0) continue;

                    var headers = RowOf(g, headerIdx);
                    var fieldCols = new int[PullFields.Length];
                    var extraCols = new int[PullFields.Length];
                    for (int f = 0; f < PullFields.Length; f++)
                    {
                        fieldCols[f] = IoListLogic.FindColumn(headers, PullFields[f].Fragments);
                        extraCols[f] = PullFields[f].Extra == null ? -1 : IoListLogic.FindColumn(headers, PullFields[f].Extra);
                    }

                    index.SheetNames.Add(cs.Name);
                    for (int r = headerIdx + 1; r < g.Rows; r++)
                    {
                        bool hasAddress = IoListLogic.TryParseAddress(
                            g.Texts[r, iAddr], out char dir, out string node, out string off, out string bit);

                        string key, address;
                        if (hasAddress)
                        {
                            key = IoListLogic.JoinKey(dir, node, off, card, bit);
                            address = IoListLogic.BuildAddress(dir, node, off, bit);
                        }
                        else if (iId >= 0 && g.Texts[r, iId].Length > 0)
                        {
                            key = IoListLogic.NodeKey(card, g.Texts[r, iId]);          // module row
                            address = IoListLogic.ModuleAddress(g.Texts[r, iId]);
                        }
                        else continue;

                        var fields = new string[PullFields.Length];
                        for (int f = 0; f < PullFields.Length; f++)
                            fields[f] = f == BitField && hasAddress
                                ? bit                                   // the parsed bit, not the raw address
                                : CellOf(g, r, fieldCols[f], extraCols[f]);

                        // Module-row mnemonics carry a "BUS_" prefix the CSV never has — drop it here so
                        // everything downstream (pairing, comparison, display) sees one spelling.
                        if (!hasAddress)
                            fields[MnemonicField] = IoListLogic.StripBusPrefix(fields[MnemonicField]);

                        var row = new IoRow { Address = address, Fields = fields };
                        Add(hasAddress ? index.ByAddress : index.ByNode, key, row);
                        foreach (string mk in IoListLogic.MnemonicKeys(fields[MnemonicField]))
                            Add(index.ByMnemonic, mk, row);
                        index.RowsIndexed++;
                    }
                }

                if (index.SheetsUsed == 0)
                {
                    Dialogs.Info($"No readable card sheet for controller {controller}. Sheets must carry a " +
                                 "header row with \"Bit\" and \"Mnemonic\":\n" + path);
                    return null;
                }
                return index;
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

        private static void Add<T>(Dictionary<string, List<T>> map, string key, T value)
        {
            if (!map.TryGetValue(key, out var list)) map[key] = list = new List<T>();
            list.Add(value);
        }

        /// <summary>
        /// Picks the controller (rack + slot) whose cards this CSV belongs to. Silent when the workbook
        /// covers only one; otherwise the user chooses, because nothing in the CSV identifies the slot.
        /// Returns null when there is nothing to choose or the user cancels.
        /// </summary>
        private static string ChooseController(List<CardSheet> cardSheets, string path)
        {
            var controllers = new List<string>();
            foreach (var cs in cardSheets)
                if (!controllers.Contains(cs.Controller)) controllers.Add(cs.Controller);

            if (controllers.Count == 0)
            {
                Dialogs.Info("No card sheet found in this workbook. Sheets must be named " +
                             "R{rack}.S{slot}.C{card} (a label may follow, e.g. \"R1.S2.C1 - PC2\"):\n" + path);
                return null;
            }
            if (controllers.Count == 1) return controllers[0];

            var labels = new string[controllers.Count];
            for (int i = 0; i < controllers.Count; i++)
            {
                var sheets = new List<string>();
                foreach (var cs in cardSheets)
                    if (cs.Controller == controllers[i]) sheets.Add(cs.Name);
                labels[i] = $"{controllers[i]}  —  {string.Join(", ", sheets)}";
            }

            int pick = Dialogs.SelectOne(
                "This IO list covers more than one controller. A card number is only unique within a " +
                "slot, and the CSV does not say which slot it belongs to.\n\nPick the controller this CSV describes:",
                labels, "Select controller");
            return pick < 0 ? null : controllers[pick];
        }

        /// <summary>One field value, optionally continued by a second column (description part 1 + part 2).</summary>
        private static string CellOf(SheetGrid g, int row, int col, int extraCol)
        {
            string main = col >= 0 ? g.Texts[row, col].Trim() : "";
            string extra = extraCol >= 0 ? g.Texts[row, extraCol].Trim() : "";
            if (extra.Length == 0) return main;
            return main.Length == 0 ? extra : main + " " + extra;
        }

        private static string JoinField(List<IoRow> rows, int fieldIndex)
        {
            if (rows.Count == 1) return rows[0].Fields[fieldIndex];
            var seen = new List<string>();
            foreach (IoRow row in rows)
            {
                string v = row.Fields[fieldIndex];
                if (v.Length > 0 && !seen.Contains(v)) seen.Add(v);
            }
            return string.Join("/", seen);
        }

        // ---- CSV file reading (headless, via Core) ----------------------------------------------------

        /// <summary>One plant-CSV signal: its name and where the CSV puts it.</summary>
        private sealed class CsvSignal
        {
            public string Name;
            public string Address;      // "I.90.0.0", or "#236" for a bus-diagnostic row
        }

        private sealed class CsvIndex
        {
            public readonly Dictionary<string, List<CsvSignal>> ByAddress = new Dictionary<string, List<CsvSignal>>();
            public readonly Dictionary<string, List<CsvSignal>> ByNode = new Dictionary<string, List<CsvSignal>>();
            public readonly Dictionary<string, List<CsvSignal>> ByMnemonic = new Dictionary<string, List<CsvSignal>>();
        }

        private static int LoadCsvSignals(
            string path,
            CsvIndex index,
            out bool dirUnknown,
            out int virtualSkipped,
            out int busRows)
        {
            dirUnknown = false;
            virtualSkipped = 0;
            busRows = 0;
            try
            {
                var doc = CsvFile.Load(File.ReadAllBytes(path));
                var rows = new List<string[]>(doc.Records.Count);
                foreach (var rec in doc.Records) rows.Add(rec.Fields);

                int headerIdx = CsvDetect.DetectHeaderRow(rows, "#");
                if (headerIdx < 0) return -1;

                var headers = rows[headerIdx];
                int iName = IoListLogic.FindColumn(headers, CsvNameHeaders);
                int iNode = IoListLogic.FindColumn(headers, "node");
                int iOff = IoListLogic.FindColumn(headers, "offset");
                int iBit = IoListLogic.FindColumn(headers, "bit");
                int iBoard = IoListLogic.FindColumn(headers, CsvBoardHeaders);
                int iType = IoListLogic.FindColumn(headers, CsvTypeHeaders);
                if (iName < 0 || iNode < 0 || iOff < 0 || iBit < 0 || iBoard < 0) return -1;

                char dir = IoListLogic.DirectionFromName(Path.GetFileName(path));
                dirUnknown = dir == '\0';

                int count = 0;
                for (int r = headerIdx + 1; r < rows.Count; r++)
                {
                    string[] f = rows[r];
                    if (CsvDetect.IsCommentFields(f, "#")) continue;   // '#' rows are disabled signals — not pulled
                    if (CsvDetect.IsCommentFields(f, "!")) continue;   // '!' rows are include directives

                    string type = iType >= 0 && f.Length > iType ? f[iType] : "";
                    if (IoListLogic.IsVirtualType(type))
                    {
                        virtualSkipped++;   // types 5/6: virtual signals — not hardware I/O
                        continue;
                    }
                    if (f.Length <= iName || f.Length <= iNode || f.Length <= iBoard) continue;

                    string name = f[iName].Trim();
                    string node = f[iNode], board = f[iBoard];
                    if (name.Length == 0 || node.Trim().Length == 0 || board.Trim().Length == 0) continue;

                    // Type 11 — bus diagnostics: the node alone is the address.
                    if (IoListLogic.IsBusDiagnosticType(type))
                    {
                        busRows++;
                        var busSignal = new CsvSignal { Name = name, Address = IoListLogic.ModuleAddress(node) };
                        Add(index.ByNode, IoListLogic.NodeKey(board, node), busSignal);
                        Add(index.ByMnemonic, IoListLogic.MnemonicKey(name), busSignal);
                        count++;
                        continue;
                    }

                    if (f.Length <= iOff || f.Length <= iBit) continue;
                    string off = f[iOff], bit = f[iBit];
                    if (off.Trim().Length == 0 || bit.Trim().Length == 0) continue;

                    // Direction is part of the key; an unknown direction is indexed as both so it still matches.
                    var signal = new CsvSignal { Name = name, Address = IoListLogic.BuildAddress(dir, node, off, bit) };
                    if (dir != 'O') Add(index.ByAddress, IoListLogic.JoinKey('I', node, off, board, bit), signal);
                    if (dir != 'I') Add(index.ByAddress, IoListLogic.JoinKey('O', node, off, board, bit), signal);
                    Add(index.ByMnemonic, IoListLogic.MnemonicKey(name), signal);
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
