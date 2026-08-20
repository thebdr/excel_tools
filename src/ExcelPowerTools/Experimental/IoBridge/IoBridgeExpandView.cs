using System;
using System.Collections.Generic;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Experimental.IoBridge;
using ExcelPowerTools.Interop;
using CoreIoListLogic = ExcelPowerTools.Core.Experimental.IoListLogic;

namespace ExcelPowerTools.Experimental.IoBridge
{
    /// <summary>What the _Expand sheet currently unfolds — the anchor every expand action queues against.</summary>
    internal sealed class IoBridgeExpandContext
    {
        public string CardSheetName = "";
        public BridgeRow Row;
    }

    /// <summary>
    /// The expanded view: a temporary _Expand sheet unfolding ONE multi-mnemonic point into one row
    /// per constituent, in the exact card-sheet column format — MATCHED pairs first (the IO mnemonic
    /// column split, each part beside its CSV record), then IO-only parts, then CSV-only records.
    /// The sheet is a SCRATCHPAD: delete or edit rows freely, then run the row actions there to queue
    /// the reconciliation — merge the CSV names into the IO list, make the CSV records match the
    /// sheet (delete the missing, append the new), or update the IO address from the CSV side.
    /// </summary>
    internal static class IoBridgeExpandView
    {
        internal const string SheetName = "_Expand";
        private const int HeaderRow = 2, FirstDataRow = 3;

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int HeaderBlue = Rgb(180, 210, 240);

        // ---- opening ---------------------------------------------------------------------------------

        public static void ShowForSelection()
        {
            IoBridgeSession session = IoBridgeSession.Current;
            if (session == null)
            {
                Dialogs.Info("No live IO Bridge workspace. Use \"IO Bridge: open workspace\" first.");
                return;
            }

            dynamic app = ExcelApp.App;
            dynamic sheet = app.ActiveSheet;
            string sheetName;
            try { sheetName = (string)sheet.Name; } catch { sheetName = ""; }

            if (!session.SheetMaps.TryGetValue(sheetName, out BridgeSheetMap map))
            {
                Dialogs.Info("Select a row on a bridge card sheet first.");
                return;
            }

            int rowNumber;
            try { rowNumber = (int)app.ActiveCell.Row; } catch { rowNumber = -1; }
            if (rowNumber < 0 || !map.RowByWorksheetRow.TryGetValue(rowNumber, out BridgeRow row))
            {
                Dialogs.Info("The selected row is not a signal row.");
                return;
            }

            List<string> ioParts = SplitParts(row.Signal.Mnemonic);
            bool multiCsv = row.Csv != null && row.Csv.IsMulti;
            if (!multiCsv && ioParts.Count <= 1)
            {
                Dialogs.Info("This point has a single mnemonic and a single CSV row — nothing to expand.");
                return;
            }

            session.ExpandContext = new IoBridgeExpandContext { CardSheetName = sheetName, Row = row };
            IoBridgeEditCapture.RunSuppressed(() => Render(app, session, map, row));
        }

        internal static List<string> SplitParts(string mnemonicList)
        {
            var parts = new List<string>();
            foreach (string part in (mnemonicList ?? "").Split('/'))
            {
                string p = part.Trim();
                if (p.Length > 0) parts.Add(p);
            }
            return parts;
        }

        // ---- rendering -------------------------------------------------------------------------------

        private sealed class ExpandLine
        {
            public string IoPart = "";       // this line's IO mnemonic constituent ("" for CSV-only)
            public string[] CsvFields;       // the CSV record behind the line, null for IO-only
            public string Label = "";
        }

        private static void Render(dynamic app, IoBridgeSession session, BridgeSheetMap map, BridgeRow row)
        {
            dynamic workbook = IoBridgeWorkspace.FindLiveWorkbook(app);
            if (workbook == null) return;

            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
            dynamic expand = null;
            try
            {
                foreach (dynamic ws in workbook.Worksheets)
                    if ((string)ws.Name == SheetName) { expand = ws; break; }
                if (expand == null)
                {
                    expand = workbook.Worksheets.Add(
                        Type.Missing, workbook.Worksheets(workbook.Worksheets.Count));
                    expand.Name = SheetName;
                }
                expand.Cells.Clear();
            }
            finally
            {
                try { app.DisplayAlerts = prevAlerts; } catch { }
            }

            int columns = session.TotalSheetColumns;
            expand.Cells(1, 1).Value =
                "Expanded view · " + map.SheetName + " · " +
                (row.Signal.Address.Length > 0 ? row.Signal.Address + " · " : "") + row.Signal.Mnemonic +
                " — edit this sheet freely, then right-click ▸ IO Bridge row actions to queue the reconciliation";
            expand.Cells(1, 1).Font.Bold = true;

            var headers = new object[1, columns];
            for (int c = 0; c < session.Layout.Columns.Count; c++)
                headers[0, session.WorksheetColumnOf(c) - 1] = session.Layout.Columns[c].Def.SampleHeader;
            dynamic headerRange = expand.Range(expand.Cells(HeaderRow, 1), expand.Cells(HeaderRow, columns));
            headerRange.NumberFormat = "@";
            IoBridgeBuilder.WriteBlock(expand, HeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            IoBridgeBuilder.PaintCsvHeaderBlock(expand, session, HeaderRow);
            headerRange.WrapText = true;
            try { headerRange.Font.Name = "Arial Narrow"; } catch { }

            List<ExpandLine> lines = PairConstituents(session, row);
            if (lines.Count > 0)
            {
                CsvColumnsMap csvMap = CsvMapOf(session, row);
                var data = new object[lines.Count, columns];
                for (int r = 0; r < lines.Count; r++)
                    FillLine(data, r, lines[r], session, map, row, csvMap);

                dynamic block = expand.Range(
                    expand.Cells(FirstDataRow, 1), expand.Cells(FirstDataRow + lines.Count - 1, columns));
                block.NumberFormat = "@";
                IoBridgeBuilder.WriteBlock(expand, FirstDataRow, 1, data);
            }

            // Autofit from the DATA rows only — the meta line in A1 must not stretch column 1 (and the
            // wrapped headers must not stretch the rest).
            try
            {
                if (lines.Count > 0)
                    expand.Range(
                        expand.Cells(FirstDataRow, 1),
                        expand.Cells(FirstDataRow + lines.Count - 1, columns)).Columns.AutoFit();

                const double minColumnWidth = 8.5;
                for (int c = 1; c <= columns; c++)
                    if ((double)expand.Columns(c).ColumnWidth < minColumnWidth)
                        expand.Columns(c).ColumnWidth = minColumnWidth;

                expand.Rows(HeaderRow).AutoFit();
                if ((double)expand.Rows(HeaderRow).RowHeight > 48) expand.Rows(HeaderRow).RowHeight = 48;
            }
            catch { }
            try { expand.Activate(); } catch { }
        }

        /// <summary>Matched pairs first (IO order), then IO-only parts, then unclaimed CSV records.</summary>
        private static List<ExpandLine> PairConstituents(IoBridgeSession session, BridgeRow row)
        {
            List<string> ioParts = SplitParts(row.Signal.Mnemonic);
            var csvRecords = row.Csv != null ? row.Csv.AllFields : new List<string[]>();
            CsvColumnsMap csvMap = CsvMapOf(session, row);

            var claimed = new bool[csvRecords.Count];
            var matched = new List<ExpandLine>();
            var ioOnly = new List<ExpandLine>();

            foreach (string part in ioParts)
            {
                string key = CoreIoListLogic.MnemonicKey(part);
                int found = -1;
                for (int i = 0; i < csvRecords.Count && found < 0; i++)
                    if (!claimed[i] && CoreIoListLogic.MnemonicKey(NameOf(csvRecords[i], csvMap)) == key)
                        found = i;

                if (found >= 0)
                {
                    claimed[found] = true;
                    matched.Add(new ExpandLine { IoPart = part, CsvFields = csvRecords[found], Label = "matched" });
                }
                else
                {
                    ioOnly.Add(new ExpandLine { IoPart = part, Label = "IO only" });
                }
            }

            var lines = new List<ExpandLine>(matched);
            lines.AddRange(ioOnly);
            for (int i = 0; i < csvRecords.Count; i++)
                if (!claimed[i])
                    lines.Add(new ExpandLine { CsvFields = csvRecords[i], Label = "CSV only" });
            return lines;
        }

        private static void FillLine(
            object[,] data, int r, ExpandLine line, IoBridgeSession session, BridgeSheetMap map,
            BridgeRow row, CsvColumnsMap csvMap)
        {
            for (int c = 0; c < session.Layout.Columns.Count; c++)
            {
                LayoutColumn column = session.Layout.Columns[c];
                int sheetCol = session.WorksheetColumnOf(c) - 1;
                switch (column.Def.Side)
                {
                    case "IO":
                        // IO context only on IO-side lines; the mnemonic column carries THIS part alone.
                        if (line.IoPart.Length > 0)
                        {
                            if (column.Def.NormalizedHeader == "mnemonic")
                            {
                                data[r, sheetCol] = line.IoPart;
                            }
                            else
                            {
                                int ioCol = map.IoColOfLayout[c];
                                if (row.Kind != BridgeRowKind.CsvOnly && ioCol >= 0 && ioCol < row.IoValues.Length)
                                    data[r, sheetCol] = row.IoValues[ioCol];
                            }
                        }
                        break;
                    case "CSV":
                        if (line.CsvFields != null)
                        {
                            int canonical = IoBridgeSession.CanonicalIndexOf(column.Def.NormalizedHeader);
                            int field = canonical >= 0 ? CsvCanonical.FieldIndex(csvMap, canonical) : -1;
                            if (field >= 0 && field < line.CsvFields.Length && line.CsvFields[field] != null)
                                data[r, sheetCol] = line.CsvFields[field];
                        }
                        break;
                    default:
                        data[r, sheetCol] = column.Def.NormalizedHeader == "status.row" ? line.Label : "";
                        break;
                }
            }
        }

        internal static CsvColumnsMap CsvMapOf(IoBridgeSession session, BridgeRow row) =>
            row.Csv != null && row.Csv.File == BridgeTarget.OutputsCsv
                ? session.Model.OutputsMap
                : session.Model.InputsMap;

        internal static string NameOf(string[] fields, CsvColumnsMap map) =>
            map != null && map.Name >= 0 && map.Name < fields.Length && fields[map.Name] != null
                ? fields[map.Name].Trim()
                : "";
    }

    /// <summary>
    /// The action menu of the _Expand scratchpad. Everything reads the SHEET AS IT STANDS — rows the
    /// user deleted or renamed count — and queues ordinary bridge actions against the original point.
    /// </summary>
    internal static class IoBridgeExpandActions
    {
        private const int FirstDataRow = 3;

        public static void ShowMenu()
        {
            IoBridgeSession session = IoBridgeSession.Current;
            IoBridgeExpandContext context = session?.ExpandContext;
            if (session == null || context?.Row == null)
            {
                Dialogs.Info("The expanded view is empty — open it from a card-sheet row first.");
                return;
            }
            if (!session.SheetMaps.TryGetValue(context.CardSheetName, out BridgeSheetMap map))
            {
                Dialogs.Info("The expanded point's card sheet is gone — refresh and expand again.");
                return;
            }

            dynamic app = ExcelApp.App;
            dynamic expand = app.ActiveSheet;
            BridgeRow row = context.Row;

            // The sheet as it stands: per line, the IO part and the CSV name (+ address cells).
            List<SheetLine> lines = ReadSheet(expand, session);
            var csvNames = new List<string>();
            foreach (SheetLine line in lines)
                if (line.CsvName.Length > 0 && !csvNames.Contains(line.CsvName)) csvNames.Add(line.CsvName);

            var labels = new List<string>();
            var actions = new List<Action>();

            // 1. IO-list mnemonics ← every CSV name on the sheet, joined.
            string ioMnemonicCell = CurrentIoMnemonic(session, map, row);
            string merged = string.Join("/", csvNames);
            if (csvNames.Count > 0 && !CoreIoListLogic.MnemonicSetsMatch(merged, ioMnemonicCell))
            {
                labels.Add("M — Update IO-list mnemonics ← the sheet's CSV names  (\"" + merged + "\")");
                actions.Add(() => QueueSet(session, row, BridgeTarget.IoList, "mnemonic",
                    LinkedField.Mnemonic, ioMnemonicCell, merged));
            }

            // 2. CSV ← the sheet: delete records whose name vanished, append rows for new names.
            List<Action> syncActions = BuildCsvSync(session, map, row, lines, out string syncSummary);
            if (syncActions.Count > 0)
            {
                labels.Add("U — Update CSV ← this sheet  (" + syncSummary + ")");
                actions.Add(() => { foreach (Action a in syncActions) a(); });
            }

            // 3. IO address ← the CSV side.
            if (TryAddressUpdate(session, map, row, lines, out string addressLabel, out Action addressAction))
            {
                labels.Add(addressLabel);
                actions.Add(addressAction);
            }

            if (labels.Count == 0)
            {
                Dialogs.Info("The sheet already agrees with both documents — nothing to queue.");
                return;
            }

            int pick = Dialogs.SelectOne(
                "Expanded point: " + (row.Signal.Address.Length > 0 ? row.Signal.Address + " · " : "") +
                row.Signal.Mnemonic + "\n\nQueue which reconciliation?",
                labels.ToArray(), "IO Bridge — expanded view");
            if (pick < 0) return;

            actions[pick]();
            session.SaveSidecarNow();

            // Live feedback on the ORIGINAL card sheet.
            try
            {
                dynamic workbook = IoBridgeWorkspace.FindLiveWorkbook(app);
                dynamic cardSheet = null;
                foreach (dynamic ws in workbook.Worksheets)
                    if ((string)ws.Name == context.CardSheetName) { cardSheet = ws; break; }
                if (cardSheet != null)
                    IoBridgeEditCapture.RunSuppressed(() =>
                        IoBridgeBuilder.RenderActionHighlights(cardSheet, map, session));
            }
            catch { }
            IoBridgeQueueSheet.Refresh(session);
        }

        // ---- reading the scratchpad ------------------------------------------------------------------

        private sealed class SheetLine
        {
            public string IoPart = "", CsvName = "";
            public string Node = "", Offset = "", Bit = "", Description = "";
        }

        private static List<SheetLine> ReadSheet(dynamic expand, IoBridgeSession session)
        {
            var lines = new List<SheetLine>();
            int lastRow = (int)RangeHelpers.LastRow(expand);

            int ioMnemonicCol = SheetColumnOfIo(session, "mnemonic");
            int nameCol = SheetColumnOfCanonical(session, CsvCanonical.Name);
            int nodeCol = SheetColumnOfCanonical(session, CsvCanonical.Node);
            int offCol = SheetColumnOfCanonical(session, CsvCanonical.Offset);
            int bitCol = SheetColumnOfCanonical(session, CsvCanonical.Bit);
            int descCol = SheetColumnOfCanonical(session, CsvCanonical.Description);

            for (int r = FirstDataRow; r <= lastRow; r++)
            {
                var line = new SheetLine
                {
                    IoPart = CellText(expand, r, ioMnemonicCol),
                    CsvName = CellText(expand, r, nameCol),
                    Node = CellText(expand, r, nodeCol),
                    Offset = CellText(expand, r, offCol),
                    Bit = CellText(expand, r, bitCol),
                    Description = CellText(expand, r, descCol),
                };
                if (line.IoPart.Length > 0 || line.CsvName.Length > 0) lines.Add(line);
            }
            return lines;
        }

        private static string CellText(dynamic sheet, int row, int column)
        {
            if (column <= 0) return "";
            try { return ((string)CsvSheet.CellText(sheet.Cells(row, column).Value2, sheet, row, column)).Trim(); }
            catch { return ""; }
        }

        private static int SheetColumnOfIo(IoBridgeSession session, string normalizedHeader)
        {
            for (int c = 0; c < session.Layout.Columns.Count; c++)
                if (session.Layout.Columns[c].Def.Side == "IO" &&
                    session.Layout.Columns[c].Def.NormalizedHeader == normalizedHeader)
                    return session.WorksheetColumnOf(c);
            return -1;
        }

        private static int SheetColumnOfCanonical(IoBridgeSession session, int canonical)
        {
            string key = CsvCanonical.LayoutKey(canonical);
            for (int c = 0; c < session.Layout.Columns.Count; c++)
                if (session.Layout.Columns[c].Def.NormalizedHeader == key)
                    return session.WorksheetColumnOf(c);
            return -1;
        }

        // ---- the three reconciliations ---------------------------------------------------------------

        private static string CurrentIoMnemonic(IoBridgeSession session, BridgeSheetMap map, BridgeRow row)
        {
            for (int c = 0; c < session.Layout.Columns.Count; c++)
                if (session.Layout.Columns[c].Def.Side == "IO" &&
                    session.Layout.Columns[c].Def.NormalizedHeader == "mnemonic")
                {
                    int ioCol = map.IoColOfLayout[c];
                    if (ioCol >= 0 && ioCol < row.IoValues.Length && row.IoValues[ioCol] != null)
                        return row.IoValues[ioCol].Trim();
                }
            return row.Signal.Mnemonic;
        }

        private static List<Action> BuildCsvSync(
            IoBridgeSession session, BridgeSheetMap map, BridgeRow row, List<SheetLine> lines,
            out string summary)
        {
            var actions = new List<Action>();
            summary = "";
            CsvColumnsMap csvMap = IoBridgeExpandView.CsvMapOf(session, row);
            if (csvMap == null) { return actions; }

            var existing = row.Csv != null ? row.Csv.AllFields : new List<string[]>();
            var sheetKeys = new List<string>();
            foreach (SheetLine line in lines)
                if (line.CsvName.Length > 0) sheetKeys.Add(CoreIoListLogic.MnemonicKey(line.CsvName));

            // Deletes: claimed records whose name no longer appears on the sheet.
            int deletes = 0;
            var existingKeys = new List<string>();
            foreach (string[] record in existing)
            {
                string name = IoBridgeExpandView.NameOf(record, csvMap);
                string key = CoreIoListLogic.MnemonicKey(name);
                existingKeys.Add(key);
                if (name.Length == 0 || sheetKeys.Contains(key)) continue;

                deletes++;
                actions.Add(() => session.Queue.Record(new BridgeAction
                {
                    Kind = BridgeActionKind.DeleteCsvRow,
                    Target = TargetOf(session, row),
                    Signal = PerRecordSignal(row, name),
                    OldValue = "",
                    NewValue = "+",
                    QueuedUtc = DateTime.UtcNow.ToString("O"),
                }));
            }

            // Adds: sheet rows whose CSV name is not among the claimed records — appended at the end.
            int adds = 0;
            string promptedType = null;
            foreach (SheetLine line in lines)
            {
                if (line.CsvName.Length == 0) continue;
                if (existingKeys.Contains(CoreIoListLogic.MnemonicKey(line.CsvName))) continue;

                string[] fields = ComposeRecord(session, map, row, csvMap, line, existing, ref promptedType);
                if (fields == null) continue;

                adds++;
                string name = line.CsvName;
                actions.Add(() => session.Queue.Record(new BridgeAction
                {
                    Kind = BridgeActionKind.CreateCsvRow,
                    Target = TargetOf(session, row),
                    Signal = PerRecordSignal(row, name),
                    OldValue = "",
                    NewValue = "+",
                    RowFields = fields,
                    QueuedUtc = DateTime.UtcNow.ToString("O"),
                }));
            }

            if (actions.Count > 0)
                summary = deletes + " delete(s), " + adds + " new row(s) at the end";
            return actions;
        }

        /// <summary>
        /// A new record: cloned from a sibling of the same point when one exists (same address, type,
        /// flags — only the name and description vary), else composed like Integrate with the type
        /// asked once for the whole sync.
        /// </summary>
        private static string[] ComposeRecord(
            IoBridgeSession session, BridgeSheetMap map, BridgeRow row, CsvColumnsMap csvMap,
            SheetLine line, List<string[]> siblings, ref string promptedType)
        {
            string[] fields;
            if (siblings.Count > 0)
            {
                fields = (string[])siblings[0].Clone();
            }
            else
            {
                if (!CoreIoListLogic.TryParseAddress(row.RawBitText, out char dir, out string node, out string off, out string bit))
                    return null;
                if (promptedType == null)
                {
                    promptedType = PromptType(dir == 'I', row);
                    if (promptedType == null) return null;
                }

                fields = new string[csvMap.Headers.Count];
                for (int i = 0; i < fields.Length; i++) fields[i] = "0";
                fields[0] = "@";
                if (csvMap.Type >= 0) fields[csvMap.Type] = promptedType;
                fields[csvMap.Node] = node;
                fields[csvMap.Offset] = off;
                fields[csvMap.Bit] = bit;
                fields[csvMap.Board] = map.Model.Card;
            }

            fields[0] = "@";
            fields[csvMap.Name] = line.CsvName;
            SetIf(fields, csvMap.Node, line.Node);
            SetIf(fields, csvMap.Offset, line.Offset);
            SetIf(fields, csvMap.Bit, line.Bit);

            // Description: what the user typed on the sheet wins; else the composed one, automatically.
            int descField = CsvCanonical.FieldIndex(csvMap, CsvCanonical.Description);
            if (descField >= 0)
            {
                string description = line.Description.Length > 0
                    ? line.Description
                    : IoBridgeRowActions.ComposeCsvDescription(map, row);
                if (description.Length > 0) fields[descField] = description;
            }
            return fields;
        }

        private static string PromptType(bool isInput, BridgeRow row) =>
            IoBridgeRowActions.PromptCsvType(isInput, "the new rows of " + row.Signal.Mnemonic);

        private static void SetIf(string[] fields, int index, string value)
        {
            if (index >= 0 && index < fields.Length && value.Length > 0) fields[index] = value;
        }

        private static bool TryAddressUpdate(
            IoBridgeSession session, BridgeSheetMap map, BridgeRow row, List<SheetLine> lines,
            out string label, out Action action)
        {
            label = null;
            action = null;
            if (row.Kind != BridgeRowKind.Signal) return false;
            if (!CoreIoListLogic.TryParseAddress(row.RawBitText, out char dir, out _, out _, out _)) return false;

            // The first sheet line carrying a full CSV address is the source.
            foreach (SheetLine line in lines)
            {
                if (line.Node.Length == 0 || line.Offset.Length == 0 || line.Bit.Length == 0) continue;
                string candidate = CoreIoListLogic.BuildAddress(dir, line.Node, line.Offset, line.Bit);
                if (candidate == row.RawBitText.Trim()) return false;   // already agrees

                label = "A — Update IO address ← the CSV side  (\"" + row.RawBitText.Trim() + "\" → \"" + candidate + "\")";
                string oldValue = row.RawBitText;
                action = () => QueueSet(session, row, BridgeTarget.IoList, "bit",
                    LinkedField.Address, oldValue, candidate);
                return true;
            }
            return false;
        }

        // ---- queue plumbing --------------------------------------------------------------------------

        private static BridgeTarget TargetOf(IoBridgeSession session, BridgeRow row) =>
            row.Csv != null ? row.Csv.File
            : CoreIoListLogic.TryParseAddress(row.RawBitText, out char dir, out _, out _, out _) && dir == 'O'
                ? BridgeTarget.OutputsCsv
                : BridgeTarget.InputsCsv;

        /// <summary>
        /// Per-record identity: name only. Several records share the point's address, so the address
        /// would be ambiguous — the empty Address routes resolution through the mnemonic.
        /// </summary>
        private static SignalRef PerRecordSignal(BridgeRow row, string name) => new SignalRef
        {
            CardKey = row.Signal.CardKey,
            Address = "",
            Mnemonic = name,
        };

        private static void QueueSet(
            IoBridgeSession session, BridgeRow row, BridgeTarget target, string fieldHeader,
            LinkedField linked, string oldValue, string newValue)
        {
            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = target,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                Linked = linked,
                FieldHeader = fieldHeader,
                OldValue = oldValue,
                NewValue = newValue,
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }
    }
}
