using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using ExcelPowerTools.Interop;
using CoreIoListLogic = ExcelPowerTools.Core.IoPipeline.IoListLogic;
using CsvSignalType = ExcelPowerTools.Core.IoPipeline.CsvSignalType;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// Row-level actions on the selected bridge row, offered through one picker: the four linked-pair
    /// updates in either direction (address, mnemonic, description, normal condition), integrate an
    /// IO-only signal into its CSV, disable/enable, delete. Every choice queues — nothing applies here.
    ///
    /// Every entry leads with a FIXED accelerator letter, one per action type: typing it in the picker
    /// selects the entry (the list's type-ahead), Enter queues it. CSV-bound and IO-bound variants of
    /// a pair use adjacent letters (M/N mnemonic, A/B address, C/D condition).
    /// </summary>
    internal static class IoBridgeRowActions
    {
        private const string KeyMnemonicCsv = "M", KeyMnemonicIo = "N";
        private const string KeyAddressCsv = "A", KeyAddressIo = "B";
        private const string KeyConditionCsv = "C", KeyConditionIo = "D";
        private const string KeyDescription = "E", KeyIntegrate = "I";
        private const string KeyDisable = "X", KeyEnable = "V", KeyDelete = "Z";
        private const string KeyType = "T";
        private sealed class Choice
        {
            public string Label;
            public Action Perform;
        }

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

            // On the _Expand scratchpad the same command opens the reconciliation menu.
            if (sheetName == IoBridgeExpandView.SheetName)
            {
                IoBridgeExpandActions.ShowMenu();
                return;
            }

            if (!session.SheetMaps.TryGetValue(sheetName, out BridgeSheetMap map))
            {
                Dialogs.Info("Select a row on a bridge card sheet first.");
                return;
            }

            List<BridgeRow> rows = SelectedModelRows(app, map);
            if (rows.Count == 0)
            {
                Dialogs.Info("The selection contains no signal rows.");
                return;
            }

            List<Choice> choices = rows.Count == 1
                ? BuildChoices(session, map, rows[0])
                : BuildBulkChoices(session, map, rows);
            if (choices.Count == 0)
            {
                Dialogs.Info("No row action applies to the selection.");
                return;
            }

            var labels = new string[choices.Count];
            for (int i = 0; i < choices.Count; i++) labels[i] = choices[i].Label;

            string prompt = (rows.Count == 1
                ? RowDescription(rows[0]) + "\n\nQueue which action?"
                : rows.Count + " selected signal rows.\n\nQueue which action for each applicable row?") +
                "  (press an action's letter, Enter to queue)";
            int pick = Dialogs.SelectOne(prompt, labels, "IO Bridge — row actions", monospaced: true);
            if (pick < 0) return;

            choices[pick].Perform();
            session.SaveSidecarNow();

            // Live feedback: the affected cells light up and show their queued values right away.
            IoBridgeEditCapture.RunSuppressed(() =>
                IoBridgeBuilder.RenderActionHighlights(sheet, map, session));
            IoBridgeQueueSheet.Refresh(session);
        }

        /// <summary>
        /// Distinct model rows behind the current selection, in worksheet order (capped) — VISIBLE
        /// rows only. A selection drawn over a filtered sheet spans the hidden rows too, and an
        /// action the user aimed at the rows they can see must never land on the ones the filter
        /// took away (field report 2026-11-03).
        /// </summary>
        private static List<BridgeRow> SelectedModelRows(dynamic app, BridgeSheetMap map)
        {
            const int MaxRows = 5000;
            var seen = new HashSet<int>();
            var rows = new List<BridgeRow>();
            try
            {
                dynamic sheet = app.ActiveSheet;
                dynamic areas = app.Selection.Areas;
                int areaCount = (int)areas.Count;
                for (int a = 1; a <= areaCount; a++)
                {
                    dynamic area = areas(a);
                    int first = (int)area.Row;
                    int count = (int)area.Rows.Count;
                    for (int r = first; r < first + count && seen.Count <= MaxRows; r++)
                    {
                        if (!seen.Add(r) || !map.RowByWorksheetRow.TryGetValue(r, out BridgeRow row)) continue;
                        if (IsRowHidden(sheet, r)) continue;
                        rows.Add(row);
                    }
                }
            }
            catch
            {
                try
                {
                    int r = (int)app.ActiveCell.Row;
                    if (map.RowByWorksheetRow.TryGetValue(r, out BridgeRow row)) rows.Add(row);
                }
                catch { }
            }
            return rows;
        }

        /// <summary>
        /// The multi-row menu: each entry queues the action for EVERY selected row it applies to,
        /// silently skipping the rest. The label carries the applicable-row count.
        /// </summary>
        private static List<Choice> BuildBulkChoices(IoBridgeSession session, BridgeSheetMap map, List<BridgeRow> rows)
        {
            var choices = new List<Choice>();

            AddBulkLinked(choices, session, map, rows, LinkedField.Mnemonic, "Mnemonic");
            AddBulkLinked(choices, session, map, rows, LinkedField.Address, "Address");
            AddBulkLinked(choices, session, map, rows, LinkedField.NormalCondition, "Normal condition");

            var describable = rows.FindAll(r =>
                r.Kind != BridgeRowKind.CsvOnly && r.Csv != null && !r.Csv.IsMulti);
            if (describable.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeyDescription, "Description", "IOL → CSV", Count(describable.Count)),
                    Perform = () =>
                    {
                        foreach (BridgeRow row in describable)
                        {
                            var single = new List<Choice>();
                            AddDescriptionToCsv(single, session, map, row);
                            foreach (Choice c in single) c.Perform();
                        }
                    },
                });

            // Empty-mnemonic points are unused spares — integrating them would append nameless lines.
            // Bus modules integrate too (auto-typed 11, no prompt).
            var integrable = rows.FindAll(r =>
                (r.Kind == BridgeRowKind.Signal || r.Kind == BridgeRowKind.Module) &&
                r.Csv == null && r.Signal.Mnemonic.Trim().Length > 0);
            if (integrable.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeyIntegrate, "Create CSV record", "from IOL", Count(integrable.Count)),
                    Perform = () => QueueIntegrateMany(session, map, integrable),
                });

            // Virtual twins: the record carries the mnemonic but acquires nothing — make it real.
            var twins = rows.FindAll(r => r.Csv != null && !r.Csv.IsMulti && r.Csv.Virtual);
            if (twins.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeyType, "CSV type", "virtual twin → real", Count(twins.Count)),
                    Perform = () => QueueTypeChangeMany(session, twins),
                });

            var enabled = rows.FindAll(r => r.Csv != null && !r.Csv.IsMulti && !r.Csv.Disabled);
            if (enabled.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeyDisable, "Disable CSV record", "", Values("@", "#") + Count(enabled.Count)),
                    Perform = () => { foreach (BridgeRow r in enabled) QueueRowState(session, r, currentlyDisabled: false); },
                });

            var disabled = rows.FindAll(r => r.Csv != null && !r.Csv.IsMulti && r.Csv.Disabled);
            if (disabled.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeyEnable, "Enable CSV record", "", Values("#", "@") + Count(disabled.Count)),
                    Perform = () => { foreach (BridgeRow r in disabled) QueueRowState(session, r, currentlyDisabled: true); },
                });

            var deletable = rows.FindAll(r => r.Csv != null && !r.Csv.IsMulti);
            if (deletable.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeyDelete, "Delete CSV record", "", Count(deletable.Count)),
                    Perform = () => { foreach (BridgeRow r in deletable) QueueDelete(session, r); },
                });

            return choices;
        }

        private static void AddBulkLinked(
            List<Choice> choices, IoBridgeSession session, BridgeSheetMap map, List<BridgeRow> rows,
            LinkedField pair, string label)
        {
            var csvSide = new List<LinkedValues>();
            var ioSide = new List<LinkedValues>();
            foreach (BridgeRow row in rows)
            {
                if (row.Kind == BridgeRowKind.CsvOnly || row.Csv == null) continue;
                bool singleCsv = !row.Csv.IsMulti;
                if (!TryLinkedValues(session, map, row, pair, out LinkedValues values)) continue;
                if (singleCsv && values.CsvNew != values.CsvOld && CsvAlignAllowed(row, pair)) csvSide.Add(values);
                if (values.IoNew != values.IoOld && IoAlignAllowed(row, pair)) ioSide.Add(values);
            }

            if (csvSide.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeysFor(pair).Csv, label, "IOL → CSV", Count(csvSide.Count)),
                    Perform = () => { foreach (LinkedValues v in csvSide) v.QueueCsv(); },
                });
            if (ioSide.Count > 0)
                choices.Add(new Choice
                {
                    Label = Line(KeysFor(pair).Io, label, "CSV → IOL", Count(ioSide.Count)),
                    Perform = () => { foreach (LinkedValues v in ioSide) v.QueueIo(); },
                });
        }

        /// <summary>
        /// Integrates many rows: bus modules go straight in as type 11; signals ask the CSV Type ONCE
        /// per direction and apply it to all of that direction.
        /// </summary>
        private static void QueueIntegrateMany(IoBridgeSession session, BridgeSheetMap map, List<BridgeRow> rows)
        {
            var byDirection = new Dictionary<char, List<BridgeRow>>();
            foreach (BridgeRow row in rows)
            {
                if (row.Kind == BridgeRowKind.Module)
                {
                    QueueIntegrate(session, map, row, "11");
                    continue;
                }
                if (!CoreIoListLogic.TryParseAddress(row.RawBitText, out char dir, out _, out _, out _)) continue;
                if (!byDirection.TryGetValue(dir, out List<BridgeRow> list)) byDirection[dir] = list = new List<BridgeRow>();
                list.Add(row);
            }

            foreach (KeyValuePair<char, List<BridgeRow>> group in byDirection)
            {
                string type = PromptCsvType(group.Key == 'I',
                    group.Value.Count + " signal(s), direction " + (group.Key == 'I' ? "Input" : "Output"));
                if (type == null) continue;   // this direction skipped, others still offered
                foreach (BridgeRow row in group.Value)
                    QueueIntegrate(session, map, row, type);
            }
        }

        private static string RowDescription(BridgeRow row) =>
            row.Signal.Address.Length > 0
                ? row.Signal.Address + "  ·  " + row.Signal.Mnemonic
                : row.Signal.Mnemonic;

        private static List<Choice> BuildChoices(IoBridgeSession session, BridgeSheetMap map, BridgeRow row)
        {
            var choices = new List<Choice>();
            bool hasCsv = row.Csv != null;
            bool singleCsv = hasCsv && !row.Csv.IsMulti;   // CSV-side writes need ONE unambiguous record
            bool isSpine = row.Kind != BridgeRowKind.CsvOnly;

            if (isSpine && singleCsv)
            {
                AddLinkedPair(choices, session, map, row, LinkedField.Mnemonic, "Mnemonic");
                AddLinkedPair(choices, session, map, row, LinkedField.Address, "Address");
                AddLinkedPair(choices, session, map, row, LinkedField.NormalCondition, "Normal condition");
                AddDescriptionToCsv(choices, session, map, row);
            }
            else if (isSpine && hasCsv)
            {
                // Multi-record point: IO-side updates still make sense; CSV-side ones are ambiguous.
                AddLinkedPair(choices, session, map, row, LinkedField.Mnemonic, "Mnemonic", ioSideOnly: true);
                AddLinkedPair(choices, session, map, row, LinkedField.NormalCondition, "Normal condition", ioSideOnly: true);
            }

            if (isSpine && !hasCsv && row.Signal.Mnemonic.Trim().Length > 0)
            {
                if (row.Kind == BridgeRowKind.Signal)
                    choices.Add(new Choice
                    {
                        Label = Line(KeyIntegrate, "Create CSV record", "from IOL",
                            "(" + row.Signal.Mnemonic + " @ " + row.Signal.Address + ")"),
                        Perform = () =>
                        {
                            if (!CoreIoListLogic.TryParseAddress(row.RawBitText, out char dir, out _, out _, out _))
                            {
                                Dialogs.Info("This row's address cannot be parsed — nothing to integrate.");
                                return;
                            }
                            string type = PromptCsvType(dir == 'I', row.Signal.Mnemonic + " (" + row.Signal.Address + ")");
                            if (type != null) QueueIntegrate(session, map, row, type);
                        },
                    });
                else if (row.Kind == BridgeRowKind.Module)
                    choices.Add(new Choice
                    {
                        // A BUS module IS type 11 by definition — no prompt.
                        Label = Line(KeyIntegrate, "Create CSV record", "from IOL",
                            "(" + CoreIoListLogic.StripBusPrefix(row.Signal.Mnemonic) + " @ " + row.Signal.Address + ", type 11)"),
                        Perform = () => QueueIntegrate(session, map, row, "11"),
                    });
            }

            if (singleCsv && row.Csv.Virtual)
                choices.Add(new Choice
                {
                    // The twin looks like the signal but acquires nothing from the field: give it a
                    // real type (same picker as Integrate); the address update then follows as usual.
                    Label = Line(KeyType, "CSV type", "virtual twin → real", Values(row.CsvValues[CsvCanonical.Type], "…")),
                    Perform = () =>
                    {
                        string type = PromptCsvType(row.Csv.File == BridgeTarget.InputsCsv,
                            row.Signal.Mnemonic + " (virtual twin)", "Change CSV type");
                        if (type != null) QueueTypeChange(session, row, type);
                    },
                });

            if (singleCsv)
            {
                bool disabled = row.Csv.Disabled;
                choices.Add(new Choice
                {
                    Label = disabled
                        ? Line(KeyEnable, "Enable CSV record", "", Values("#", "@"))
                        : Line(KeyDisable, "Disable CSV record", "", Values("@", "#")),
                    Perform = () => QueueRowState(session, row, disabled),
                });
                choices.Add(new Choice
                {
                    Label = Line(KeyDelete, "Delete CSV record", "", ""),
                    Perform = () => QueueDelete(session, row),
                });
            }

            return choices;
        }

        // ---- linked-pair updates ---------------------------------------------------------------------

        private static (string Csv, string Io) KeysFor(LinkedField pair)
        {
            switch (pair)
            {
                case LinkedField.Mnemonic: return (KeyMnemonicCsv, KeyMnemonicIo);
                case LinkedField.Address: return (KeyAddressCsv, KeyAddressIo);
                default: return (KeyConditionCsv, KeyConditionIo);
            }
        }

        private static void AddLinkedPair(
            List<Choice> choices, IoBridgeSession session, BridgeSheetMap map, BridgeRow row,
            LinkedField pair, string label, bool ioSideOnly = false)
        {
            (string csvKey, string ioKey) = KeysFor(pair);
            if (TryLinkedValues(session, map, row, pair, out LinkedValues values))
            {
                if (!ioSideOnly && values.CsvNew != values.CsvOld && CsvAlignAllowed(row, pair))
                    choices.Add(new Choice
                    {
                        Label = Line(csvKey, label, "IOL → CSV", Values(values.CsvOld, values.CsvNew)),
                        Perform = () => values.QueueCsv(),
                    });
                if (values.IoNew != values.IoOld && IoAlignAllowed(row, pair))
                    choices.Add(new Choice
                    {
                        Label = Line(ioKey, label, "CSV → IOL", Values(values.IoOld, values.IoNew)),
                        Perform = () => values.QueueIo(),
                    });
            }
        }

        /// <summary>
        /// Aligning a MOVED row's address (the CSV record carries the IO row's name at another
        /// address) is refused on the side where the destination address is already taken — moving
        /// there would collide with another record (CSV side) or duplicate another row's address
        /// (IO-list side), as in a swapped pair or a shifted block. The other side stays offered.
        /// </summary>
        private static bool CsvAlignAllowed(BridgeRow row, LinkedField pair) =>
            pair != LinkedField.Address || row.Csv == null || !row.Csv.Moved || !row.Csv.IoAddressOccupiedInCsv;

        private static bool IoAlignAllowed(BridgeRow row, LinkedField pair) =>
            pair != LinkedField.Address || row.Csv == null || !row.Csv.Moved ||
            (!row.Csv.CsvAddressOccupiedInIo && !row.Csv.OtherCard);   // an IO address cannot name another card

        private sealed class LinkedValues
        {
            public string CsvOld, CsvNew, IoOld, IoNew;
            public Action QueueCsv, QueueIo;
        }

        /// <summary>
        /// Resolves both sides of a linked pair for one row: each side's current value and what the
        /// OTHER side would write into it. False when either column cannot be located.
        /// </summary>
        private static bool TryLinkedValues(
            IoBridgeSession session, BridgeSheetMap map, BridgeRow row, LinkedField pair, out LinkedValues values)
        {
            values = null;

            switch (pair)
            {
                case LinkedField.Mnemonic:
                {
                    int ioCol = IoColumnByKey(map, row, session, "mnemonic");
                    if (ioCol < 0) return false;
                    string ioValue = IoValue(row, ioCol);
                    string csvValue = row.CsvValues[CsvCanonical.Name];

                    // The comparison uses the row's EFFECTIVE mnemonic (the exclude regex may have
                    // filtered generator keys out of the cell); an IO-side write composes the cell
                    // back the way the IO list spells it — BUS_ prefix on module rows, filtered keys
                    // re-attached — so the list's own conventions are never overwritten.
                    bool busPrefixed = row.Kind == BridgeRowKind.Module &&
                        ioValue.Trim().StartsWith("BUS_", StringComparison.OrdinalIgnoreCase);
                    string ioNew = (busPrefixed ? "BUS_" : "") + csvValue +
                        (row.ExcludedMnemonics.Length > 0 ? "/" + row.ExcludedMnemonics : "");

                    values = MakeSetPair(session, map, row, pair,
                        csvFieldCanonical: CsvCanonical.Name, csvOld: csvValue,
                        csvNew: CoreIoListLogic.StripBusPrefix(row.Signal.Mnemonic),
                        ioFieldHeader: "mnemonic", ioOld: ioValue, ioNew: ioNew);
                    return true;
                }
                case LinkedField.NormalCondition:
                {
                    int ioCol = IoColumnByKey(map, row, session, "normalcondition");
                    if (ioCol < 0) return false;
                    string ioValue = IoValue(row, ioCol);
                    string csvValue = row.CsvValues[CsvCanonical.Logic];

                    // The two sides compare through the CSV's spelling: an empty IO-list cell IS
                    // Logic 0, so "" vs "0" is agreement, never a pair of opposite updates. A side
                    // holding something neither file accepts offers nothing in its direction —
                    // the CSV takes only 0/1, and junk never crosses over.
                    bool ioReadable = CoreIoListLogic.TryNormalConditionToLogic(ioValue, out string ioAsLogic);
                    bool csvReadable = CoreIoListLogic.IsValidLogic(csvValue);
                    string csvNew = ioReadable ? ioAsLogic : csvValue;                      // unreadable IO → no CSV update
                    string ioNew = csvReadable && (!ioReadable || ioAsLogic != csvValue.Trim())
                        ? csvValue.Trim() : ioValue;                                        // unreadable CSV → no IO update

                    values = MakeSetPair(session, map, row, pair,
                        csvFieldCanonical: CsvCanonical.Logic, csvOld: csvValue, csvNew: csvNew,
                        ioFieldHeader: "normalcondition", ioOld: ioValue, ioNew: ioNew);
                    return true;
                }
                case LinkedField.Address:
                {
                    // IO side: the Bit column holds "I.n.o.b". CSV side: three fields.
                    if (row.Kind != BridgeRowKind.Signal) return false;
                    string ioAddress = row.RawBitText.Trim();
                    if (!CoreIoListLogic.TryParseAddress(ioAddress, out char dir, out string n, out string o, out string b))
                        return false;

                    string csvNode = row.CsvValues[CsvCanonical.Node];
                    string csvOff = row.CsvValues[CsvCanonical.Offset];
                    string csvBit = row.CsvValues[CsvCanonical.Bit];
                    string csvComposite = CoreIoListLogic.BuildAddress(
                        '\0', csvNode.Trim(), csvOff.Trim(), csvBit.Trim());
                    string ioComposite = CoreIoListLogic.BuildAddress('\0', n, o, b);

                    // An off-card moved row also differs by CARD: the composites carry it, so the CSV
                    // side is offered even when node/offset/bit agree.
                    bool offCard = row.Csv != null && row.Csv.Moved && row.Csv.OtherCard;
                    if (offCard)
                    {
                        csvComposite += " · card " + row.Csv.CsvCard;
                        ioComposite += " · card " + map.Model.Card;
                    }

                    var result = new LinkedValues
                    {
                        CsvOld = csvComposite,
                        CsvNew = ioComposite,
                        IoOld = ioAddress,
                        IoNew = CoreIoListLogic.BuildAddress(dir, csvNode, csvOff, csvBit),
                    };
                    result.QueueCsv = () =>
                    {
                        // Only the parts that differ — a moved record usually changes just the offset.
                        if (Norm(csvNode) != Norm(n)) QueueCsvSet(session, row, LinkedField.Address, CsvCanonical.Node, csvNode, n);
                        if (Norm(csvOff) != Norm(o)) QueueCsvSet(session, row, LinkedField.Address, CsvCanonical.Offset, csvOff, o);
                        if (Norm(csvBit) != Norm(b)) QueueCsvSet(session, row, LinkedField.Address, CsvCanonical.Bit, csvBit, b);
                        if (offCard) QueueCsvSet(session, row, LinkedField.Address, CsvCanonical.Card, row.CsvValues[CsvCanonical.Card], map.Model.Card);
                        AutoQueueDescription(session, map, row);
                    };
                    result.QueueIo = () =>
                        QueueIoSet(session, row, LinkedField.Address, "bit", ioAddress, result.IoNew);
                    values = result;
                    return true;
                }
                default:
                    return false;
            }
        }

        /// <summary>
        /// The CSV description composed EXACTLY as the export tool writes it, from this row's IO-list
        /// fields: language-1 parts + Location&amp;Device + [drawing code]. "" when the row carries no
        /// IO side or the sheet lacks the columns.
        /// </summary>
        internal static string ComposeCsvDescription(BridgeSheetMap map, BridgeRow row)
        {
            if (row == null || row.Kind == BridgeRowKind.CsvOnly) return "";
            string[] headers = map.Model.IoHeaders;

            int part1 = CoreIoListLogic.FindColumn(headers, "Description language 1 (1° part", "description english language");
            int part2 = CoreIoListLogic.FindColumn(headers, "Description language 1 (2° part");
            int location = CoreIoListLogic.FindColumn(headers, "location");
            int device = CoreIoListLogic.FindColumn(headers, "device");
            int drawing = CoreIoListLogic.FindColumn(headers, "drawing name", "drawing");

            return CsvDescription.Compose(
                IoValue(row, part1), IoValue(row, part2),
                IoValue(row, location), IoValue(row, device), IoValue(row, drawing));
        }

        private static void AddDescriptionToCsv(
            List<Choice> choices, IoBridgeSession session, BridgeSheetMap map, BridgeRow row)
        {
            // Description flows IO → CSV only (the CSV never feeds the electrical descriptions).
            string description = ComposeCsvDescription(map, row);
            if (description.Length == 0) return;

            string csvValue = row.CsvValues[CsvCanonical.Description];
            if (description == csvValue) return;

            choices.Add(new Choice
            {
                Label = Line(KeyDescription, "Description", "IOL → CSV", Values(csvValue, description)),
                Perform = () => QueueCsvSet(session, row, LinkedField.Description, CsvCanonical.Description, csvValue, description),
            });
        }

        /// <summary>
        /// Every CSV content update refreshes the description alongside — queued automatically so the
        /// composed text follows whatever field just changed.
        /// </summary>
        private static void AutoQueueDescription(IoBridgeSession session, BridgeSheetMap map, BridgeRow row)
        {
            if (row?.Csv == null || row.Csv.IsMulti) return;
            string description = ComposeCsvDescription(map, row);
            if (description.Length == 0) return;

            string csvValue = row.CsvValues[CsvCanonical.Description];
            if (description == csvValue) return;
            QueueCsvSet(session, row, LinkedField.Description, CsvCanonical.Description, csvValue, description);
        }

        // ---- queue helpers ---------------------------------------------------------------------------

        private static LinkedValues MakeSetPair(
            IoBridgeSession session, BridgeSheetMap map, BridgeRow row, LinkedField pair,
            int csvFieldCanonical, string csvOld, string csvNew,
            string ioFieldHeader, string ioOld, string ioNew)
        {
            return new LinkedValues
            {
                CsvOld = csvOld,
                CsvNew = csvNew,
                IoOld = ioOld,
                IoNew = ioNew,
                QueueCsv = () =>
                {
                    QueueCsvSet(session, row, pair, csvFieldCanonical, csvOld, csvNew);
                    AutoQueueDescription(session, map, row);   // a CSV row update carries its description
                },
                QueueIo = () => QueueIoSet(session, row, pair, ioFieldHeader, ioOld, ioNew),
            };
        }

        private static void QueueCsvSet(
            IoBridgeSession session, BridgeRow row, LinkedField pair, int canonical, string oldValue, string newValue)
        {
            string fieldHeader = session.CsvFieldHeaderFor(row, canonical);
            if (fieldHeader.Length == 0) return;
            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = row.Csv.File,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                Linked = pair,
                FieldHeader = fieldHeader,
                OldValue = oldValue,
                NewValue = newValue,
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }

        private static void QueueIoSet(
            IoBridgeSession session, BridgeRow row, LinkedField pair, string fieldHeader, string oldValue, string newValue)
        {
            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = BridgeTarget.IoList,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                Linked = pair,
                FieldHeader = fieldHeader,
                OldValue = oldValue,
                NewValue = newValue,
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }

        private static void QueueRowState(IoBridgeSession session, BridgeRow row, bool currentlyDisabled)
        {
            session.Queue.Record(new BridgeAction
            {
                Kind = currentlyDisabled ? BridgeActionKind.EnableCsvRow : BridgeActionKind.DisableCsvRow,
                Target = row.Csv.File,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                OldValue = currentlyDisabled ? "1" : "0",
                NewValue = currentlyDisabled ? "0" : "1",
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }

        private static void QueueDelete(IoBridgeSession session, BridgeRow row)
        {
            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.DeleteCsvRow,
                Target = row.Csv.File,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                OldValue = "",
                NewValue = "+",   // "+" = gone, per the re-anchor encoding
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }

        /// <summary>
        /// The CSV Type is not derivable from the IO list — one picker, shared by all callers,
        /// listing the controller's own taxonomy (<see cref="CsvSignalType"/>) for the direction.
        /// </summary>
        internal static string PromptCsvType(bool isInput, string subject, string caption = "Integrate into CSV")
        {
            List<CsvSignalType.Entry> entries = CsvSignalType.ForDirection(isInput);
            var typeOptions = new string[entries.Count];
            for (int i = 0; i < entries.Count; i++) typeOptions[i] = CsvSignalType.Label(entries[i]);
            int pick = Dialogs.SelectOne("CSV Type for " + subject + ":", typeOptions, caption);
            return pick < 0 ? null : entries[pick].Number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// A virtual twin becomes a real signal through its Type field — a plain SetField on the
        /// record, which the apply resolves by mnemonic (the twin's address fields are stale).
        /// </summary>
        private static void QueueTypeChange(IoBridgeSession session, BridgeRow row, string type)
        {
            string header = session.CsvFieldHeaderFor(row, CsvCanonical.Type);
            if (header.Length == 0 || row.Csv == null) return;
            string current = CsvCanonical.Type < row.CsvValues.Length ? row.CsvValues[CsvCanonical.Type] : "";
            if (current == type) return;

            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = row.Csv.File,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                FieldHeader = header,
                OldValue = current,
                NewValue = type,
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }

        /// <summary>One type prompt per CSV file, applied to every twin of that file.</summary>
        private static void QueueTypeChangeMany(IoBridgeSession session, List<BridgeRow> twins)
        {
            foreach (BridgeTarget file in new[] { BridgeTarget.InputsCsv, BridgeTarget.OutputsCsv })
            {
                var group = twins.FindAll(r => r.Csv.File == file);
                if (group.Count == 0) continue;
                string type = PromptCsvType(file == BridgeTarget.InputsCsv,
                    group.Count + " virtual twin(s) in " + (file == BridgeTarget.InputsCsv ? "Inputs" : "Outputs"),
                    "Change CSV type");
                if (type == null) continue;
                foreach (BridgeRow row in group) QueueTypeChange(session, row, type);
            }
        }

        private static void QueueIntegrate(IoBridgeSession session, BridgeSheetMap map, BridgeRow row, string type)
        {
            string node, off, bit;
            bool isInput;
            if (row.Kind == BridgeRowKind.Module)
            {
                // A bus module: the node IS the address ("#236"); type-11 records live in Inputs.
                if (!row.Signal.Address.StartsWith("#", StringComparison.Ordinal)) return;
                node = row.Signal.Address.Substring(1);
                off = "0";
                bit = "0";
                isInput = true;
            }
            else
            {
                if (!CoreIoListLogic.TryParseAddress(row.RawBitText, out char dir, out node, out off, out bit))
                    return;
                isInput = dir == 'I';
            }

            // The CSV never carries the IO list's BUS_ prefix on type-11 signals.
            string name = type == "11"
                ? CoreIoListLogic.StripBusPrefix(row.Signal.Mnemonic).Trim()
                : row.Signal.Mnemonic.Trim();
            if (name.Length == 0) return;   // an unused spare — never append a nameless CSV line

            BridgeTarget target = isInput ? BridgeTarget.InputsCsv : BridgeTarget.OutputsCsv;
            CsvColumnsMap csvMap = isInput ? session.Model.InputsMap : session.Model.OutputsMap;

            // Normal condition from the IO list ("" and "0" are Logic 0); anything unreadable, or
            // no column at all, takes the neutral 0 — the CSV accepts nothing but 0/1.
            int ncCol = IoColumnByKey(map, row, session, "normalcondition");
            if (ncCol < 0 || !CoreIoListLogic.TryNormalConditionToLogic(IoValue(row, ncCol), out string logic)) logic = "0";

            var fields = new string[csvMap.Headers.Count];
            for (int i = 0; i < fields.Length; i++) fields[i] = "0";
            fields[0] = "@";
            fields[csvMap.Name] = name;
            if (csvMap.Type >= 0) fields[csvMap.Type] = type;
            int logicField = CsvCanonical.FieldIndex(csvMap, CsvCanonical.Logic);
            if (logicField >= 0) fields[logicField] = logic;
            fields[csvMap.Node] = node;
            fields[csvMap.Offset] = off;
            fields[csvMap.Bit] = bit;
            fields[csvMap.Board] = SheetCardOf(session, map);
            // A created row carries the FULL composed description automatically.
            int descField = CsvCanonical.FieldIndex(csvMap, CsvCanonical.Description);
            if (descField >= 0) fields[descField] = ComposeCsvDescription(map, row);

            session.Queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.CreateCsvRow,
                Target = target,
                Signal = IoBridgeEditCapture.CloneSignal(row.Signal),
                OldValue = "",
                NewValue = "+",   // "+" = the row exists
                RowFields = fields,
                QueuedUtc = DateTime.UtcNow.ToString("O"),
            });
        }

        // ---- lookups ---------------------------------------------------------------------------------

        private static string SheetCardOf(IoBridgeSession session, BridgeSheetMap map) => map.Model.Card;

        private static string IoValue(BridgeRow row, int ioColumn) =>
            ioColumn >= 0 && ioColumn < row.IoValues.Length && row.IoValues[ioColumn] != null
                ? row.IoValues[ioColumn] : "";

        private static int IoColumnByKey(BridgeSheetMap map, BridgeRow row, IoBridgeSession session, string normalizedKey)
        {
            string[] headers = map.Model.IoHeaders;
            for (int c = 0; c < headers.Length; c++)
                if (CoreIoListLogic.NormalizeHeader(headers[c]) == normalizedKey) return c;
            return -1;
        }


        private static string Norm(string value) => CoreIoListLogic.Norm(value);

        /// <summary>Hidden by an AutoFilter, a manual hide or an outline collapse — all are "not shown".</summary>
        private static bool IsRowHidden(dynamic sheet, int row)
        {
            try { return (bool)sheet.Rows(row).Hidden; }
            catch { return false; }   // unreadable → treat as visible rather than silently drop it
        }

        // ---- labels: one shape for every entry ---------------------------------------------------
        //   K — Field              IOL → CSV   ("old" → "new")
        // The field is the column's own word, the two place names are "IO list" and "CSV" in the
        // order the write goes, and the tail is what changes. Padded so the columns line up in the
        // picker's list, which this menu opens in a monospaced face.

        private static string Line(string key, string field, string direction, string tail) =>
            key + " — " + field.PadRight(20) + direction.PadRight(12) + tail;

        private static string Values(string oldValue, string newValue) =>
            "(\"" + Shorten(oldValue) + "\" → \"" + Shorten(newValue) + "\")";

        private static string Count(int rows) => "— " + rows + " row(s)";

        private static string Shorten(string value) =>
            value == null ? "" : value.Length <= 24 ? value : value.Substring(0, 24) + "…";
    }
}
