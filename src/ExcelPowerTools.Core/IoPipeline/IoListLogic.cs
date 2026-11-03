using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline
{
    /// <summary>
    /// IO Pipeline — pure logic for the IO-list ↔ plant-CSV bridge.
    ///
    /// IO-list layout: one worksheet per card, named "R{rack}.S{slot}.C{card}" with an optional free-text
    /// suffix ("R1.S2.C1 - PC2"); the CARD comes from the sheet NAME, not from a column. Signal rows carry
    /// the full address in the "Bit" column as "I.node.offset.bit" ("Q" = output). Module rows (Type "PA")
    /// have an empty "Bit" and carry the node in "ID".
    ///
    /// Plant CSV: Node / Offset / Bit + Card (a.k.a. "Board number"). Bus-diagnostic signals (Type 11)
    /// have no bit address — they join on the node alone, against the module rows.
    ///
    /// The physical join key is DIRECTION.node.offset.card.bit; the node-only key is card#node.
    /// Remove together with src/ExcelPowerTools/IoPipeline and the tests.
    /// </summary>
    public static class IoListLogic
    {
        private static readonly Regex CardSheetName =
            new Regex(@"^\s*R(\d+)\s*\.\s*S(\d+)\s*\.\s*C(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Rack/slot/card from a worksheet NAME: "R1.S1.C2" or "R1.S2.C1 - PC2" → 1 / 2 / 1. Anything after
        /// the card number is a label and is ignored; components normalize numerically ("C04" → "4").
        /// </summary>
        public static bool TryParseCardSheetName(string sheetName, out string rack, out string slot, out string card)
        {
            rack = slot = card = null;
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Match m = CardSheetName.Match(sheetName);
            if (!m.Success) return false;

            rack = Norm(m.Groups[1].Value);
            slot = Norm(m.Groups[2].Value);
            card = Norm(m.Groups[3].Value);
            return true;
        }

        /// <summary>
        /// Parses the IO-list address "I.2.0.0" / "Q.90.1.4" into direction and numerically-normalized
        /// node/offset/bit. "Q" (Siemens output) is reported as 'O' so it meets the CSV Outputs side.
        /// </summary>
        public static bool TryParseAddress(string s, out char direction, out string node, out string offset, out string bit)
        {
            direction = '\0'; node = offset = bit = null;
            if (string.IsNullOrWhiteSpace(s)) return false;

            string[] parts = s.Trim().Split('.');
            if (parts.Length != 4 || parts[0].Length != 1) return false;

            char d = char.ToUpperInvariant(parts[0][0]);
            if (d == 'Q') d = 'O';                       // the new lists write outputs as Q
            if (d != 'I' && d != 'O') return false;
            if (!AllDigits(parts[1]) || !AllDigits(parts[2]) || !AllDigits(parts[3])) return false;

            direction = d;
            node = Norm(parts[1]);
            offset = Norm(parts[2]);
            bit = Norm(parts[3]);
            return true;
        }

        private static readonly Regex PlcAddress = new Regex(
            @"^\s*%?\s*([IQEA])\s*\.?\s*(\d+)\s*\.\s*(\d+)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        /// <summary>
        /// Parses a Siemens PLC absolute address — "I.0.4" (the FVT safety list), "I0.4" (the
        /// FVZ one), "%I12.7" (TIA), German "E"/"A" — into direction and numerically-normalized
        /// byte and bit. There is NO node: a distributed input's address is unique PLC-wide, which
        /// is why the safety PLC's lists never write one. Q/A outputs are reported as 'O' like
        /// <see cref="TryParseAddress"/> does. A four-part ASC address is not a PLC address — the
        /// two parsers claim disjoint shapes, so a caller may try both.
        /// </summary>
        public static bool TryParsePlcAddress(string s, out char direction, out string byteOffset, out string bit)
        {
            direction = '\0'; byteOffset = bit = null;
            if (string.IsNullOrWhiteSpace(s)) return false;
            Match m = PlcAddress.Match(s);
            if (!m.Success) return false;

            char d = char.ToUpperInvariant(m.Groups[1].Value[0]);
            direction = d == 'I' || d == 'E' ? 'I' : 'O';
            byteOffset = Norm(m.Groups[2].Value);
            bit = Norm(m.Groups[3].Value);
            return true;
        }

        /// <summary>The PLC address as the SWP receives it: "I0.4" — TIA's absolute form without the % sign.</summary>
        public static string BuildPlcAddress(string byteOffset, string bit) => "I" + Norm(byteOffset) + "." + Norm(bit);

        private static bool AllDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (char c in s)
                if (c < '0' || c > '9') return false;
            return true;
        }

        /// <summary>The display address: "I.80.3.0" ('\0' direction → no prefix).</summary>
        public static string BuildAddress(char direction, string node, string offset, string bit)
        {
            string body = $"{Norm(node)}.{Norm(offset)}.{Norm(bit)}";
            return direction == '\0' ? body : direction + "." + body;
        }

        /// <summary>
        /// The physical join key: DIRECTION.node.offset.card.bit ("I.80.3.2.0"), numerically normalized.
        /// Direction is required — I.80.3.0 and O.80.3.0 coexist on the same card in real lists.
        /// </summary>
        public static string JoinKey(char direction, string node, string offset, string card, string bit) =>
            $"{char.ToUpperInvariant(direction)}.{Norm(node)}.{Norm(offset)}.{Norm(card)}.{Norm(bit)}";

        /// <summary>
        /// The node-only key "card#node", used by bus-diagnostic signals (CSV Type 11), which have no bit
        /// address: they join against the IO-list module row whose "ID" is that node.
        /// </summary>
        public static string NodeKey(string card, string node) => $"{Norm(card)}#{Norm(node)}";

        /// <summary>
        /// Identifies one controller: "R{rack}.S{slot}". A slot holds a controller and a controller holds
        /// several cards, so a card number is only unique WITHIN a slot — "R1.S2.C1" and "R1.S3.C1" are
        /// different cards. A plant CSV belongs to one controller and its Card column is scoped to it, so
        /// matching must be confined to a single controller before card numbers mean anything.
        /// </summary>
        public static string ControllerKey(string rack, string slot) => $"R{Norm(rack)}.S{Norm(slot)}";

        /// <summary>Numeric-normalizes a component ("07"→"7"); non-numeric values pass through trimmed.</summary>
        public static string Norm(string s)
        {
            s = (s ?? "").Trim();
            return long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out long v)
                ? v.ToString(CultureInfo.InvariantCulture)
                : s;
        }

        /// <summary>
        /// True for virtual signals — not read from hardware I/O, their node/offset/bit/card are
        /// frequently stale copy-paste values, so they are excluded from address matching.
        /// The taxonomy is <see cref="CsvSignalType"/> (the controller's own defines): virtuals are
        /// 5 VI · 6 VO · 9 AIS · 10 AOS · 13 VAI · 14 VAO.
        /// </summary>
        public static bool IsVirtualType(string typeValue)
        {
            CsvSignalType.Entry e = CsvSignalType.Find(typeValue);
            return e != null && e.Virtual;
        }

        /// <summary>
        /// True for bus-diagnostic signals (CSV Type 11, e.g. "CABINETAL00162"). They address a DEVICE, not
        /// a bit: offset/bit are 0 and the join runs on the node alone.
        /// </summary>
        public static bool IsBusDiagnosticType(string typeValue) => Norm(typeValue) == "11";

        /// <summary>
        /// Drops the "BUS_" prefix the IO list puts on module-row mnemonics ("BUS_CABINETAL00161" →
        /// "CABINETAL00161"); applied per "/"-joined part, since a module row can carry several.
        /// </summary>
        public static string StripBusPrefix(string mnemonic)
        {
            if (string.IsNullOrEmpty(mnemonic)) return mnemonic ?? "";

            string[] parts = mnemonic.Split('/');
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i].Trim();
                parts[i] = p.StartsWith("BUS_", StringComparison.OrdinalIgnoreCase) ? p.Substring(4) : p;
            }
            return string.Join("/", parts);
        }

        /// <summary>
        /// The comparison/index form of ONE mnemonic: trimmed, "BUS_" dropped, upper-cased. Two spellings
        /// of the same signal ("BUS_CABINETAL00161" and "CABINETAL00161") share a key.
        /// </summary>
        public static string MnemonicKey(string mnemonic)
        {
            string p = (mnemonic ?? "").Trim();
            if (p.StartsWith("BUS_", StringComparison.OrdinalIgnoreCase)) p = p.Substring(4);
            return p.ToUpperInvariant();
        }

        /// <summary>The distinct keys of a "/"-joined mnemonic list; empty parts are dropped.</summary>
        public static List<string> MnemonicKeys(string list)
        {
            var keys = new List<string>();
            if (string.IsNullOrEmpty(list)) return keys;

            foreach (string part in list.Split('/'))
            {
                string k = MnemonicKey(part);
                if (k.Length > 0 && !keys.Contains(k)) keys.Add(k);
            }
            return keys;
        }

        /// <summary>
        /// True when two "/"-joined mnemonic lists name the same SET of signals — order-insensitive,
        /// "BUS_"-insensitive, duplicates and empty parts ignored. Two empty lists match.
        /// </summary>
        public static bool MnemonicSetsMatch(string a, string b)
        {
            var ka = MnemonicKeys(a);
            var kb = MnemonicKeys(b);
            if (ka.Count != kb.Count) return false;

            foreach (string k in ka)
                if (!kb.Contains(k)) return false;
            return true;
        }

        /// <summary>
        /// True when every mnemonic of <paramref name="expected"/> appears in <paramref name="pulled"/>.
        /// Separates the two kinds of mismatch: a row whose own name is simply ACCOMPANIED by others
        /// (one IO-list row carrying several mnemonics — usually fine) still covers, while a row whose
        /// name is missing entirely means the two sides genuinely disagree about the signal.
        /// </summary>
        public static bool MnemonicSetCovers(string pulled, string expected)
        {
            var have = MnemonicKeys(pulled);
            foreach (string k in MnemonicKeys(expected))
                if (!have.Contains(k)) return false;
            return true;
        }

        /// <summary>
        /// The agreement verdict for one row, as "matched/total" — or EMPTY when the two sides name the
        /// same set of signals, so only disagreements draw the eye.
        ///
        /// Only the address-matched mnemonics are judged; a mnemonic found elsewhere by name is reported
        /// separately and never counts here. <paramref name="matchedByAddress"/> is what sits at the row's
        /// address, <paramref name="rowMnemonics"/> is what the row calls itself, and
        /// <paramref name="ioListMnemonics"/> is whichever of the two belongs to the IO LIST — it sets the
        /// total, since the IO list is the reference. An empty address match yields no verdict at all.
        /// </summary>
        public static string MatchRatio(string rowMnemonics, string matchedByAddress, string ioListMnemonics)
        {
            var pulled = MnemonicKeys(matchedByAddress);
            if (pulled.Count == 0) return "";                                  // nothing to compare against

            var own = MnemonicKeys(rowMnemonics);
            if (MnemonicSetsMatch(rowMnemonics, matchedByAddress)) return "";   // full agreement

            int matched = 0;
            foreach (string k in own)
                if (pulled.Contains(k)) matched++;

            int total = MnemonicKeys(ioListMnemonics).Count;
            return matched.ToString(CultureInfo.InvariantCulture) + "/" +
                   total.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// How a module-row match is shown in the address column: "#236" (node 236). Module rows have no
        /// bit address, so this is deliberately unlike the "I.236.0.0" of a signal row.
        /// </summary>
        public static string ModuleAddress(string node) => "#" + Norm(node);

        /// <summary>A mnemonic tagged with where it lives: "LP100301@O.90.0.1".</summary>
        public static string AtAddress(string mnemonic, string address) => mnemonic + "@" + address;

        /// <summary>Direction from a file/sheet name ("Plant_Outputs.csv" → 'O'); '\0' when undecidable.</summary>
        public static char DirectionFromName(string name)
        {
            string n = (name ?? "").ToLowerInvariant();
            if (n.Contains("output")) return 'O';
            if (n.Contains("input")) return 'I';
            return '\0';
        }

        /// <summary>
        /// Finds a column by header text. Headers are normalized (lowercase, whitespace/underscores
        /// stripped, dots kept); exact match wins, else the first header CONTAINING the fragment. -1 if none.
        /// </summary>
        public static int FindColumn(IReadOnlyList<string> headers, params string[] fragments)
        {
            var norm = new string[headers.Count];
            for (int i = 0; i < headers.Count; i++) norm[i] = NormalizeHeader(headers[i]);

            foreach (string f in fragments)
            {
                string nf = NormalizeHeader(f);
                for (int i = 0; i < norm.Length; i++)
                    if (norm[i] == nf) return i;
            }
            foreach (string f in fragments)
            {
                string nf = NormalizeHeader(f);
                for (int i = 0; i < norm.Length; i++)
                    if (nf.Length > 0 && norm[i].Contains(nf)) return i;
            }
            return -1;
        }

        /// <summary>
        /// Lowercase letters/digits/dots only. Degree signs, spaces and brackets vanish, so
        /// "Description language 1 (1° part) (permanent part)" → "descriptionlanguage11partpermanentpart".
        /// </summary>
        public static string NormalizeHeader(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c) || c == '.') sb.Append(char.ToLowerInvariant(c));
                // whitespace, newlines, underscores, punctuation, degree signs → dropped
            }
            return sb.ToString();
        }

        /// <summary>Adds a value to a key's list (used to join multiple mnemonics per address with "/").</summary>
        public static void AddToMap(Dictionary<string, List<string>> map, string key, string value)
        {
            if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(value)) return;
            if (!map.TryGetValue(key, out var list)) map[key] = list = new List<string>();
            if (!list.Contains(value)) list.Add(value);
        }

        public static string Joined(Dictionary<string, List<string>> map, string key) =>
            map.TryGetValue(key, out var list) ? string.Join("/", list) : "";
    }
}
