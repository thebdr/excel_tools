using System.Collections.Generic;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    public enum SwpCsvVerdict { NotChecked, Ok, NameMissing, AddressMismatch }
    public enum SwpBusfieldVerdict { NotChecked, Ok, NodeMissing, DataMismatch }
    public enum SwpAlarmVerdict { NotChecked, Ok, NotFound, DescriptionMismatch }
    public enum SwpBusExtVerdict { NotChecked, Ok, NotFound, NodeMismatch, SubDeviceMismatch }

    /// <summary>
    /// The legacy DiagCheck coherence, as spine verdict columns: signal rows are looked up in
    /// Plant_Inputs.csv (the mnemonic must exist and its node/offset/bit — fields 11/12/13 — must
    /// match the IO-list address); module rows in Busfield.csv (node + card present; element, bit
    /// and warning flag consistent with what the mnemonic encodes). The legacy DBClassID 60/21/11
    /// class check is deprecated and not ported. Both inputs are optional — absent documents
    /// leave the verdicts NotChecked.
    /// </summary>
    public static class SwpCoherence
    {
        private sealed class BusfieldRow
        {
            public string Card = "", Node = "";
            public bool IsWarning;
            public string WarnRaw = "";       // the raw warning-on-fault flag (bit 1 = ignore error)
            public string LinkedAddress = ""; // AS-I slaves name their master node here
            public int ReferenceIndex, Element, DiagBit;
            public int Line;                  // 1-based line in the busfield file
        }

        /// <summary>An active CSV record with its 1-based line — the spine keeps the line so a
        /// double-click can open the document there.</summary>
        private sealed class LineRecord
        {
            public CsvRecord Record;
            public int Line;
        }

        public static void Annotate(SwpWorkspaceModel model, CsvDocument plantInputs,
            CsvDocument busfield, SwpPlantType plantType, List<SwpLogEntry> log,
            IReadOnlyList<SwpAlarmString> alarmStrings = null,
            int firstLineId = 1, int groupIndex = 1, SwpIndexRules rules = null,
            CsvDocument busExtended = null)
        {
            if (model == null) return;
            rules = rules ?? SwpIndexRules.Official();

            List<LineRecord> inputRecords = ActiveRecords(plantInputs);
            List<BusfieldRow> busfieldRows = busfield == null ? null : ParseBusfield(busfield, log);
            if (busfieldRows != null) CheckBusfieldGlobals(busfieldRows, log);
            List<LineRecord> extendedRecords = ActiveRecords(busExtended);

            var cardBySheet = new Dictionary<string, string>();

            foreach (SwpSpineRow row in model.Rows)
            {
                DiagnosticEntry entry = row.Entry;
                if (entry == null || entry.GeneratedBy.Length > 0) continue;

                if (!entry.IsBusModule)
                {
                    if (inputRecords != null) CheckPlantInputs(row, entry, inputRecords, log);
                }
                else if (entry.AsiNode.Length > 0)
                {
                    // AS-I slaves are not Profinet nodes — the busfield never lists them; their
                    // binding lives in BUS_ExtendedDiagnostic.
                    if (extendedRecords != null)
                        CheckBusExtended(row, entry, extendedRecords, CardOf(entry, cardBySheet), log);
                }
                else if (busfieldRows != null && row.Mapped != null && row.Mapped.Ok)
                {
                    CheckBusfield(row, entry, busfieldRows, plantType, CardOf(entry, cardBySheet), log);
                }

                // Only objects whose tag family the alarm parser knows are checked — the rules
                // table says which (CLASSDIAG objects export different tags and stay NotChecked).
                if (alarmStrings != null && alarmStrings.Count > 0 && row.Mapped != null && row.Mapped.Ok &&
                    rules.AlarmCheckEnabled(plantType, SwpIndexRules.ObjectKeyFor(plantType, row.Mapped)))
                    CheckAlarmStrings(row, entry, alarmStrings, plantType, firstLineId, groupIndex, rules, log);
            }
        }

        /// <summary>
        /// Whole-file busfield rules: node 0 (the Profinet controller) must be present (§3.2),
        /// and every node another row names as its master (an AS-I master) must carry the
        /// ignore-error bit in its warning-on-fault flag — value 2 or 3 (§4.8).
        /// </summary>
        private static void CheckBusfieldGlobals(List<BusfieldRow> rows, List<SwpLogEntry> log)
        {
            bool hasNodeZero = false;
            var masterNodes = new HashSet<string>();
            foreach (BusfieldRow row in rows)
            {
                if (row.Node == "0") hasNodeZero = true;
                if (row.LinkedAddress.Length > 0 && row.LinkedAddress != "0")
                    masterNodes.Add(IoListLogic.Norm(row.LinkedAddress));
            }

            if (!hasNodeZero)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "busfield has no node 0 — the Profinet controller must always be included " +
                    "to diagnose it (rules §3.2)"));

            foreach (BusfieldRow row in rows)
            {
                if (!masterNodes.Contains(row.Node)) continue;
                if (row.WarnRaw == "2" || row.WarnRaw == "3") continue;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "busfield node " + row.Node + " is an AS-I master but its warning-on-fault " +
                    "flag is \"" + row.WarnRaw + "\" — masters need the ignore-error bit " +
                    "(write 2 or 3), or one slave fault takes the whole network down (rules §4.8)"));
            }
        }

        /// <summary>
        /// The exported alarm must exist (matched with the official v49 semantics) and its texts
        /// must carry the row's Location and Device — the operator finds the machine by them.
        /// </summary>
        private static void CheckAlarmStrings(SwpSpineRow row, DiagnosticEntry entry,
            IReadOnlyList<SwpAlarmString> alarms, SwpPlantType plantType,
            int firstLineId, int groupIndex, SwpIndexRules rules, List<SwpLogEntry> log)
        {
            SwpAlarmString alarm = SwpAlarmStrings.Match(row.Mapped, plantType, firstLineId, groupIndex, alarms, rules);
            if (alarm == null)
            {
                row.AlarmVerdict = SwpAlarmVerdict.NotFound;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    "not found in the exported alarm strings",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                return;
            }

            row.AlarmLine = alarm.Line;
            row.AlarmSource = alarm.SourceLabel;

            bool mismatch = false;
            void Contains(string what, string value, string description, string which)
            {
                if (value.Length == 0 || description.IndexOf(value, System.StringComparison.Ordinal) >= 0) return;
                mismatch = true;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    what + " \"" + value + "\" missing from the exported " + which +
                    ": \"" + description + "\"",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            }

            Contains("location", entry.Location, alarm.Description1, "description 1");
            Contains("device", entry.Device, alarm.Description1, "description 1");
            if (alarm.Description2.Trim().Length > 0)
            {
                Contains("location", entry.Location, alarm.Description2, "description 2");
                Contains("device", entry.Device, alarm.Description2, "description 2");
            }

            row.AlarmVerdict = mismatch ? SwpAlarmVerdict.DescriptionMismatch : SwpAlarmVerdict.Ok;
        }

        /// <summary>
        /// BUS_ExtendedDiagnostic.csv binds sub-module states to mnemonics — on the sorter, the
        /// AS-I slaves' CLASSDIAG alarms: "@;mnemonic;card;gateway node;sub-device;type;mask;#
        /// comment". The slave's mnemonic must be bound, to the gateway it hangs off (when the IO
        /// list's row order tells) and to the sub-device its line + address fix
        /// (<see cref="SwpAsiAddressing"/>).
        /// </summary>
        private static void CheckBusExtended(SwpSpineRow row, DiagnosticEntry entry,
            List<LineRecord> records, string card, List<SwpLogEntry> log)
        {
            int expectedSub = SwpAsiAddressing.SubDevice(entry.AsiNode);
            LineRecord found = null;
            foreach (LineRecord item in records)
            {
                if (!string.Equals(Field(item.Record, 1), entry.Mnemonic, System.StringComparison.OrdinalIgnoreCase)) continue;
                found = item;
                break;
            }

            if (found == null)
            {
                row.BusExtVerdict = SwpBusExtVerdict.NotFound;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    "AS-I slave " + entry.AsiNode + " (sub-device " + expectedSub +
                    "): mnemonic not bound in BUS_ExtendedDiagnostic",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                return;
            }

            row.BusExtLine = found.Line;
            string boundCard = IoListLogic.Norm(Field(found.Record, 2));
            string boundNode = IoListLogic.Norm(Field(found.Record, 3));
            string boundSub = IoListLogic.Norm(Field(found.Record, 4));

            if (entry.AsiMasterNode.Length > 0 && boundNode != IoListLogic.Norm(entry.AsiMasterNode))
            {
                row.BusExtVerdict = SwpBusExtVerdict.NodeMismatch;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    "BUS_ExtendedDiagnostic binds it to gateway node " + boundNode +
                    "; the IO list hangs slave " + entry.AsiNode + " off node " + entry.AsiMasterNode,
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                return;
            }
            if (card.Length > 0 && boundCard.Length > 0 && boundCard != card)
            {
                row.BusExtVerdict = SwpBusExtVerdict.NodeMismatch;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    "BUS_ExtendedDiagnostic binds it to card " + boundCard + "; the IO list has slave " +
                    entry.AsiNode + " on card " + card,
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                return;
            }
            if (boundSub != expectedSub.ToString(System.Globalization.CultureInfo.InvariantCulture))
            {
                row.BusExtVerdict = SwpBusExtVerdict.SubDeviceMismatch;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    "BUS_ExtendedDiagnostic sub-device " + boundSub + " vs " + expectedSub +
                    " for AS-I slave " + entry.AsiNode + " (line 1: A 1-31, B 33-63; line 2: A 65-95, B 97-127)",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                return;
            }

            if (entry.AsiMasterNode.Length == 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    "AS-I slave " + entry.AsiNode + " has no gateway row above it in the IO list — " +
                    "the bound gateway node " + boundNode + " was not verified",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            row.BusExtVerdict = SwpBusExtVerdict.Ok;
        }

        private static void CheckPlantInputs(SwpSpineRow row, DiagnosticEntry entry,
            List<LineRecord> records, List<SwpLogEntry> log)
        {
            foreach (LineRecord item in records)
            {
                // Legacy: a containment match — the CSV signal id carries the mnemonic.
                CsvRecord record = item.Record;
                string name = Field(record, 1);
                if (name.Length == 0 || !name.Contains(entry.Mnemonic)) continue;

                row.CsvLine = item.Line;
                bool addressMatches =
                    IoListLogic.Norm(Field(record, 11)) == entry.Node &&
                    IoListLogic.Norm(Field(record, 12)) == entry.Offset &&
                    IoListLogic.Norm(Field(record, 13)) == entry.Bit;
                if (addressMatches)
                {
                    row.CsvVerdict = SwpCsvVerdict.Ok;
                }
                else
                {
                    row.CsvVerdict = SwpCsvVerdict.AddressMismatch;
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "Plant_Inputs address differs from the IO list (" +
                        Field(record, 11) + "." + Field(record, 12) + "." + Field(record, 13) +
                        " vs " + entry.Node + "." + entry.Offset + "." + entry.Bit + ")",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                }
                return;
            }

            row.CsvVerdict = SwpCsvVerdict.NameMissing;
            log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                "mnemonic not found in Plant_Inputs",
                entry.CardSheet, entry.SourceRow, entry.Mnemonic));
        }

        private static void CheckBusfield(SwpSpineRow row, DiagnosticEntry entry,
            List<BusfieldRow> rows, SwpPlantType plantType, string card, List<SwpLogEntry> log)
        {
            BusfieldRow found = null;
            foreach (BusfieldRow candidate in rows)
            {
                if (candidate.Node != entry.Node || candidate.Card != card) continue;

                if (plantType == SwpPlantType.Ispc)
                {
                    // Reference index ≥ 10 is a line bitmask — accepted, decoded into the log.
                    if (candidate.ReferenceIndex >= 10)
                    {
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                            "busfield node " + entry.Node + " carries a line mask: " + DecodeMask(candidate.ReferenceIndex),
                            entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                    }
                    else if (candidate.ReferenceIndex != row.Mapped.ReferenceIndex) continue;
                }
                else if (plantType == SwpPlantType.Ups &&
                         candidate.ReferenceIndex != row.Mapped.ReferenceIndex) continue;

                found = candidate;
                break;
            }

            if (found == null)
            {
                row.BusfieldVerdict = SwpBusfieldVerdict.NodeMissing;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "node " + entry.Node + " (card " + card + ") not found in the busfield",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                return;
            }
            row.BusfieldLine = found.Line;

            // Two valid configurations exist. Modern (rules §3.7): the diagnostic linkage lives
            // in a type-11 BUS input, and DB class ID / Element / Bit in the busfield stay 0.
            // Legacy (deprecated): the busfield itself carries element/bit/warning-flag and they
            // must agree with what the mnemonic encodes.
            bool consistent =
                (found.Element == 0 && found.DiagBit == 0) ||
                (found.Element == row.Mapped.RawDeviceNumber &&
                 found.DiagBit == row.Mapped.Coordinate.Bit &&
                 found.IsWarning == row.Mapped.IsWarning);
            if (consistent)
            {
                row.BusfieldVerdict = SwpBusfieldVerdict.Ok;
            }
            else
            {
                row.BusfieldVerdict = SwpBusfieldVerdict.DataMismatch;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    "busfield node " + entry.Node + " disagrees: element " + found.Element +
                    " bit " + found.DiagBit + " warning " + found.IsWarning +
                    " vs mnemonic element " + row.Mapped.RawDeviceNumber +
                    " bit " + row.Mapped.Coordinate.Bit + " warning " + row.Mapped.IsWarning,
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            }
        }

        private static List<BusfieldRow> ParseBusfield(CsvDocument document, List<SwpLogEntry> log)
        {
            var rows = new List<BusfieldRow>();
            int line = 0;
            foreach (CsvRecord record in document.Records)
            {
                line++;
                if (!Field(record, 0).Contains("@")) continue;
                string node = Field(record, 2);
                if (node.Length == 0 || !AllDigits(node)) continue;
                if (record.FieldCount < 12)
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "busfield row for node " + node + " is missing columns"));
                    continue;
                }
                rows.Add(new BusfieldRow
                {
                    Card = IoListLogic.Norm(Field(record, 1)),
                    Node = IoListLogic.Norm(node),
                    IsWarning = Field(record, 4) == "1" || Field(record, 4) == "3",
                    WarnRaw = Field(record, 4),
                    LinkedAddress = Field(record, 5),
                    ReferenceIndex = IntOf(Field(record, 8)),
                    Element = IntOf(Field(record, 10)),
                    DiagBit = IntOf(Field(record, 11)),
                    Line = line,
                });
            }
            return rows;
        }

        private static List<LineRecord> ActiveRecords(CsvDocument document)
        {
            if (document == null) return null;
            var records = new List<LineRecord>();
            int line = 0;
            foreach (CsvRecord record in document.Records)
            {
                line++;
                if (Field(record, 0).Contains("@")) records.Add(new LineRecord { Record = record, Line = line });
            }
            return records;
        }

        /// <summary>The card number of the entry's R.S.C sheet — the busfield "Board Number".</summary>
        private static string CardOf(DiagnosticEntry entry, Dictionary<string, string> cache)
        {
            if (cache.TryGetValue(entry.CardSheet, out string card)) return card;
            card = IoListLogic.TryParseCardSheetName(entry.CardSheet, out _, out _, out string parsed)
                ? parsed
                : "";
            cache[entry.CardSheet] = card;
            return card;
        }

        private static string DecodeMask(int mask)
        {
            var lines = new List<string>();
            for (int bit = 0; bit < 9; bit++)
                if ((mask & (1 << bit)) != 0) lines.Add("L" + (bit + 1));
            return string.Join(" ", lines);
        }

        private static string Field(CsvRecord record, int index) =>
            record != null && index < record.FieldCount ? (record.Fields[index] ?? "").Trim() : "";

        private static int IntOf(string s) =>
            int.TryParse(s, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out int v) ? v : -1;

        private static bool AllDigits(string s)
        {
            foreach (char c in s)
                if (c < '0' || c > '9') return false;
            return s.Length > 0;
        }
    }
}
