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
    }

    /// <summary>
    /// The legacy LoadIOList harvesting rules: which IO-list rows are diagnostic signals bound for
    /// the SWP. Multi-mnemonic cells split on "/", the "BUS_" prefix is stripped, the built-in
    /// pattern union accepts alarm/warning shapes, the EMERGENCY→CABINETAL and EMERI_→ALLQ_
    /// aliases are rewritten, and the per-plant-type length rules reject malformed names. Output
    /// addresses ("Q.") never reach the SWP — it documents inputs.
    /// </summary>
    public static class SwpMnemonicFilter
    {
        public static List<DiagnosticEntry> Run(IEnumerable<IoRowData> rows, SwpFilterOptions options, List<SwpLogEntry> log)
        {
            var entries = new List<DiagnosticEntry>();
            if (rows == null) return entries;
            options = options ?? new SwpFilterOptions();

            var cleanedSeen = new HashSet<string>(System.StringComparer.Ordinal);
            foreach (IoRowData row in rows)
            {
                if (row == null) continue;
                foreach (string part in (row.MnemonicCell ?? "").Split('/'))
                {
                    DiagnosticEntry entry = FilterOne(row, part, options, cleanedSeen, log);
                    if (entry != null) entries.Add(entry);
                }
            }
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
            HashSet<string> cleanedSeen, List<SwpLogEntry> log)
        {
            string raw = (rawPart ?? "").Trim();
            if (raw.Length == 0) return null;

            string m = raw.Replace("BUS_", "");
            if (m.Length == 0) return null;

            // An include-regex match is an explicit user override: it bypasses the base length
            // gate AND the per-type length validation below — otherwise the regex could never
            // rescue a shape the built-in rules reject.
            bool rescued = options.IncludeMnemonics != null && options.IncludeMnemonics.IsMatch(m);

            // The legacy base gate: plausible mnemonic lengths only. Rows outside are the
            // non-diagnostic mass and are skipped without a log line, as the VBA did.
            if (!rescued && (m.Length <= 9 || m.Length >= 23)) return null;

            if (!rescued && !MatchesBuiltInPatterns(m)) return null;

            // EMERGENCY_ON / EMERGENCY_RST are commands, not diagnostics (legacy hard exclusion).
            if (m.Contains("EMERGENCY_ON") || m.Contains("EMERGENCY_RST")) return null;

            if (options.ExcludeMnemonics != null && options.ExcludeMnemonics.IsMatch(m))
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info, "excluded by the session regex",
                    row.CardSheet, row.SourceRow, m));
                return null;
            }

            m = ApplyRewrites(m);

            if (!rescued && !LengthValidFor(options.PlantType, m, out string lengthProblem))
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    lengthProblem + " — check the plant type (1-MSC, 2-ISPC, 3-UPS, 4-SAFETY)",
                    row.CardSheet, row.SourceRow, m));
                return null;
            }
            if (!rescued && options.PlantType == SwpPlantType.Ups && (m.Length < 12 || m.Length > 13))
            {
                // The legacy tool only warned here and kept going.
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "length " + m.Length + " is unusual for an upstream mnemonic",
                    row.CardSheet, row.SourceRow, m));
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
                if (!IoListLogic.TryParseAddress(address, out char direction, out string node, out string offset, out string bit))
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "unreadable I/O address \"" + address + "\"", row.CardSheet, row.SourceRow, m));
                    return null;
                }
                if (direction == 'O')
                {
                    // Outputs never reach the SWP (the legacy list builder skipped Q/O rows).
                    return null;
                }
                entry.Node = node;
                entry.Offset = offset;
                entry.Bit = bit;
            }
            else
            {
                // No bit address: the ID column carries the node. Profinet nodes are numeric;
                // AS-I slaves use nA/nB addressing ("1A", "10B") — both are real bus identities
                // and the SWP receives them as node.0.0 ("236.0.0", "1A.0.0" — exactly what the
                // legacy wrote). Anything else stays a valid diagnostic WITHOUT an address.
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

        /// <summary>
        /// The acceptance union, applied regardless of plant type (the per-type length rules
        /// then reject foreign shapes): CABINET/CLASSDIAG and the upstream quads anywhere,
        /// BELT/ZONE alarm-warning, EMERI and EMERGENCY only at the start — command signals
        /// like A_RESET_EMERGENCY carry the word mid-string and are not diagnostics.
        /// </summary>
        internal static bool MatchesBuiltInPatterns(string m)
        {
            return m.Contains("CABINET")
                || StartsWith(m, "BELTAL") || StartsWith(m, "BELTWR")
                || StartsWith(m, "ZONEAL") || StartsWith(m, "ZONEWR")
                || m.Contains("CLASSDIAG")
                || StartsWith(m, "EMERGENCY")
                || m.Contains("ALLQ") || m.Contains("WARQ")
                || m.Contains("ALLB") || m.Contains("WARB")
                || m.Contains("ALLZ") || m.Contains("WARZ")
                || StartsWith(m, "EMERI");
        }

        /// <summary>
        /// The aliases, in the legacy order but ANCHORED AT THE START: EMERI_ folds into ALLQ_,
        /// EMERGENCY into CABINETAL, and ALLQ_AND_ collapses to ALLQ_ with the last two
        /// characters dropped. Mid-string occurrences are left alone — the field showed
        /// A_RESET_EMERGENCY (a command) being mangled into A_RESET_CABINETAL by the old
        /// anywhere-Replace.
        /// </summary>
        internal static string ApplyRewrites(string m)
        {
            if (StartsWith(m, "EMERI_")) m = "ALLQ_" + m.Substring("EMERI_".Length);
            if (StartsWith(m, "EMERGENCY")) m = "CABINETAL" + m.Substring("EMERGENCY".Length);
            if (StartsWith(m, "ALLQ_AND_"))
            {
                m = "ALLQ_" + m.Substring("ALLQ_AND_".Length);
                m = m.Length >= 2 ? m.Substring(0, m.Length - 2) : m;
            }
            return m;
        }

        private static bool LengthValidFor(SwpPlantType plantType, string m, out string problem)
        {
            problem = "";
            switch (plantType)
            {
                case SwpPlantType.Msc:
                case SwpPlantType.Safety:
                    // CABINETAL00100 (14) or CLASSDIAG AL 0062 00001 00 → 22 chars. SAFETY
                    // mnemonics are MSC-shaped (the legacy filter shared its MSC branch too).
                    if (m.Length != 14 && m.Length != 22)
                    {
                        problem = "length " + m.Length + " is not a sorter mnemonic (14 or 22)";
                        return false;
                    }
                    return true;

                case SwpPlantType.Ispc:
                    int expected =
                        m.Contains("CABINET") ? 15 :
                        m.Contains("BELT") ? 11 :
                        m.Contains("ZONE") ? 10 : 0;
                    if (m.Length != expected || expected == 0)
                    {
                        problem = expected == 0
                            ? "not an induction mnemonic shape"
                            : "length " + m.Length + " is not an induction mnemonic (" + expected + ")";
                        return false;
                    }
                    return true;

                default:
                    return true; // UPS: warned by the caller, never rejected (legacy behavior)
            }
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
