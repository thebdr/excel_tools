using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    public sealed class SwpFilterOptions
    {
        public SwpPlantType PlantType = SwpPlantType.Msc;

        /// <summary>Extra mnemonics accepted beyond the built-in diagnostic patterns (null = none).</summary>
        public Regex IncludeMnemonics;

        /// <summary>Mnemonics dropped even when a built-in pattern accepts them (null = none).</summary>
        public Regex ExcludeMnemonics;

        /// <summary>
        /// The rule table that decides both WHETHER a mnemonic is a diagnostic and WHERE it lands
        /// (null = the shipped default). Making one table answer both questions is what removes
        /// the legacy rewrites: nothing has to be renamed into an existing shape to be recognized.
        /// </summary>
        public SwpMnemonicMap Map;
    }

    /// <summary>
    /// Which IO-list rows are diagnostic signals bound for the SWP. Multi-mnemonic cells split on
    /// "/", the "BUS_" prefix is stripped (notation, not identity — the plant CSVs know the signal
    /// without it), and then the MNEMONIC MAP decides: a shape it claims is a diagnostic, a shape
    /// it claims with no destination is recognized and dropped on purpose, and anything else is
    /// the non-diagnostic mass. Output addresses ("Q.") never reach the SWP — it documents inputs.
    ///
    /// Nothing is renamed here any more. The legacy pattern union, per-plant length rules and
    /// EMERGENCY→CABINETAL / EMERI_→ALLQ_ aliases all existed to force a shape into a branch that
    /// already knew where to write; the table states the destination instead, so a signal keeps
    /// the name its plant CSVs use and a new shape costs one row rather than a new branch.
    /// </summary>
    public static class SwpMnemonicFilter
    {
        public static List<DiagnosticEntry> Run(IEnumerable<IoRowData> rows, SwpFilterOptions options, List<SwpLogEntry> log)
        {
            var entries = new List<DiagnosticEntry>();
            if (rows == null) return entries;
            options = options ?? new SwpFilterOptions();

            var cleanedSeen = new HashSet<string>(System.StringComparer.Ordinal);
            // Resolved ONCE: the shipped table compiles some thirty regexes, and a 30k-row IO list
            // would otherwise rebuild it per mnemonic.
            SwpMnemonicMap map = options.Map ?? SwpMnemonicMapDefaults.Official();
            int candidates = 0;
            foreach (IoRowData row in rows)
            {
                if (row == null) continue;
                foreach (string part in (row.MnemonicCell ?? "").Split('/'))
                {
                    if ((part ?? "").Trim().Length > 0) candidates++;
                    DiagnosticEntry entry = FilterOne(row, part, options, map, cleanedSeen, log);
                    if (entry != null) entries.Add(entry);
                }
            }

            // Most of an IO list is not diagnostics, so an unclaimed mnemonic is normal and gets
            // no line of its own. The RATIO is the signal worth keeping: when the mnemonic map
            // claims almost nothing, the plant type is usually wrong — which the per-row length
            // errors used to shout about before the table replaced them.
            // Needs a real sample: on a handful of rows a low ratio says nothing, and a test or a
            // one-row probe must not be told its plant type is wrong.
            if (candidates >= 20 && entries.Count * 20 < candidates)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the mnemonic map claimed only " + entries.Count + " of " + candidates +
                    " mnemonic(s) for plant type " + options.PlantType +
                    " — check the plant type (1-MSC, 2-ISPC, 3-UPS, 4-SAFETY) and the map"));
            return entries;
        }

        /// <summary>
        /// Strips the KNOWN stray characters from an identity field (FUnit, Location, Device,
        /// Profinet name/IP, address, ID) — a DENYLIST on purpose: it removes only recognized
        /// debris and logs what it cleaned, so an unanticipated character surfaces instead of
        /// being silently deleted (an allowlist corrupts good data the first time it is
        /// incomplete). Current list: the stray apostrophe (Excel text-prefix residue, the one
        /// the field actually shows), its typographic cousins and the backtick, zero-width
        /// characters (deleted), and NBSP/tab (normalized to a space).
        /// </summary>
        internal static string CleanStray(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return raw ?? "";
            var sb = new System.Text.StringBuilder(raw.Length);
            foreach (char c in raw)
            {
                switch (c)
                {
                    case '\'':       // the stray Excel text-prefix apostrophe
                    case '\u2018':   // ' left single quote
                    case '\u2019':   // ' right single quote
                    case '\u201C':   // " left double quote
                    case '\u201D':   // " right double quote
                    case '`':
                    case '\u200B':   // zero-width space
                    case '\u200C':   // zero-width non-joiner
                    case '\u200D':   // zero-width joiner
                    case '\uFEFF':   // zero-width no-break space / BOM
                        break;                      // deleted
                    case '\u00A0':   // no-break space
                    case '\t':
                        sb.Append(' ');
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        private static string Clean(string raw, string field, IoRowData row,
            HashSet<string> cleanedSeen, List<SwpLogEntry> log)
        {
            string cleaned = CleanStray(raw);
            if (!ReferenceEquals(cleaned, raw) && cleaned != (raw ?? "") &&
                cleanedSeen.Add(field + "|" + raw))
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    "stray characters cleaned from " + field + ": \"" + raw + "\" → \"" + cleaned + "\"",
                    row.CardSheet, row.SourceRow));
            return cleaned;
        }

        private static DiagnosticEntry FilterOne(IoRowData row, string rawPart, SwpFilterOptions options,
            SwpMnemonicMap map, HashSet<string> cleanedSeen, List<SwpLogEntry> log)
        {
            string raw = (rawPart ?? "").Trim();
            if (raw.Length == 0) return null;

            string m = raw.Replace("BUS_", "");
            if (m.Length == 0) return null;

            // An include-regex match is an explicit user override: it rescues a shape no rule in
            // the mnemonic map claims. Such a row enters the workspace and fails later AT mapping
            // with "#Invalid Mnemonic", which is what the legacy rescue did too — the override
            // decides what is worth looking at, never where it lands.
            bool rescued = options.IncludeMnemonics != null && options.IncludeMnemonics.IsMatch(m);

            // The mnemonic map IS the filter. A shape no rule claims is the non-diagnostic mass
            // and is skipped without a log line, as the VBA did; a shape claimed by a rule with no
            // destination (the EMERGENCY_ON/RST commands, the AND/OR group members after the
            // first) is recognized ON PURPOSE and skipped just as quietly. What used to be a
            // length gate, a pattern union, a hardcoded exclusion list and a per-plant length
            // check is now one question asked of the table.
            SwpMnemonicRule rule = map.RuleFor(m, options.PlantType);
            if (rule == null && !rescued) return null;
            if (rule != null && rule.MapsNowhere) return null;

            if (options.ExcludeMnemonics != null && options.ExcludeMnemonics.IsMatch(m))
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info, "excluded by the session regex",
                    row.CardSheet, row.SourceRow, m));
                return null;
            }

            var entry = new DiagnosticEntry
            {
                CardSheet = row.CardSheet,
                SourceRow = row.SourceRow,
                RawMnemonic = raw,
                Mnemonic = m,
                TypeCell = (row.TypeCell ?? "").Trim(),
                Desc1A = row.Desc1A ?? "",
                Desc1B = row.Desc1B ?? "",
                Desc2A = row.Desc2A ?? "",
                Desc2B = row.Desc2B ?? "",
                FUnit = Clean(row.FUnit, "FUnit", row, cleanedSeen, log).Trim(),
                Location = Clean(row.Location, "Location", row, cleanedSeen, log).Trim(),
                Device = Clean(row.Device, "Device", row, cleanedSeen, log).Trim(),
                TsRef = (row.TsRef ?? "").Trim(),
                ProfinetName = Clean(row.ProfinetName, "Profinet name", row, cleanedSeen, log).Trim(),
                ProfinetIp = Clean(row.ProfinetIp, "Profinet IP", row, cleanedSeen, log).Trim(),
            };

            string address = Clean(row.AddressCell, "Bit address", row, cleanedSeen, log).Trim();
            if (address.Contains("."))
            {
                if (IoListLogic.TryParseAddress(address, out char direction, out string node, out string offset, out string bit))
                {
                    if (direction == 'O')
                    {
                        // Outputs never reach the SWP (the legacy list builder skipped Q/O rows).
                        return null;
                    }
                    entry.Node = node;
                    entry.Offset = offset;
                    entry.Bit = bit;
                }
                else if (options.PlantType == SwpPlantType.Safety &&
                         IoListLogic.TryParsePlcAddress(address, out char plcDirection, out string byteOffset, out string plcBit))
                {
                    // The safety PLC's list addresses its inputs as Siemens absolute addresses —
                    // "I.0.4" (FVT), "I0.4" (FVZ) — with no node, because a distributed input is
                    // unique PLC-wide. Only SAFETY reads them: on a sorter list a three-part
                    // address is a missing bit, and staying unreadable there keeps it visible.
                    if (plcDirection == 'O') return null;
                    entry.PlcAddress = IoListLogic.BuildPlcAddress(byteOffset, plcBit);
                    entry.Offset = byteOffset;
                    entry.Bit = plcBit;
                }
                else
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "unreadable I/O address \"" + address + "\"", row.CardSheet, row.SourceRow, m));
                    return null;
                }
            }
            else
            {
                // No bit address: the ID column carries the node. Profinet nodes are numeric and
                // the SWP receives them as node.0.0 ("236.0.0", what the legacy wrote); AS-I
                // slaves use nA/nB addressing ("1A", "10B") and the SWP receives their composed
                // line.slave identity ("1M.24A", no ".0.0"). Anything else stays a valid
                // diagnostic WITHOUT an address.
                string nodeRaw = Clean(row.IdCell, "ID", row, cleanedSeen, log).Trim();
                string node = IoListLogic.Norm(nodeRaw);
                if (IsAllDigits(node))
                {
                    entry.IsBusModule = true;
                    entry.Node = node;
                }
                else if (TryAsiNode(nodeRaw, out string asiNode))
                {
                    entry.IsBusModule = true;
                    entry.Node = asiNode;

                    // The slave's full identity: the gateway line from the Slot column (empty =
                    // line 1) + the slave address; the gateway's own node comes from the IO
                    // list's row order (a slave row sits under its gateway row).
                    string slot = Clean(row.SlotCell ?? "", "Slot", row, cleanedSeen, log).Trim();
                    entry.AsiNode = SwpAsiAddressing.Compose(slot, asiNode);
                    if (entry.AsiNode.Length == 0 && slot.Length > 0)
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "AS-I slot \"" + slot + "\" is not a gateway line (1M/2M) — slave " +
                            asiNode + " cannot be bound",
                            row.CardSheet, row.SourceRow, m));
                    entry.AsiMasterNode = IoListLogic.Norm(row.ModuleNode ?? "");
                }
                else if (nodeRaw.Length > 0)
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                        "no I/O address (the ID \"" + nodeRaw + "\" is not a bus node)",
                        row.CardSheet, row.SourceRow, m));
                }
            }

            return entry;
        }


        /// <summary>AS-I node identity: digits + A/B half ("01a" normalizes to "1A").</summary>
        internal static bool TryAsiNode(string id, out string node)
        {
            node = null;
            if (string.IsNullOrEmpty(id) || id.Length < 2) return false;
            char half = char.ToUpperInvariant(id[id.Length - 1]);
            if (half != 'A' && half != 'B') return false;
            string digits = id.Substring(0, id.Length - 1).Trim();
            if (!IsAllDigits(digits)) return false;
            node = IoListLogic.Norm(digits) + half;
            return true;
        }

        private static bool StartsWith(string s, string prefix) =>
            s.StartsWith(prefix, System.StringComparison.Ordinal);

        private static bool IsAllDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char c in s)
                if (c < '0' || c > '9') return false;
            return true;
        }
    }
}
