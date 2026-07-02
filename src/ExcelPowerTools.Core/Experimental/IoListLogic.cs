using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace ExcelPowerTools.Core.Experimental
{
    /// <summary>
    /// EXPERIMENTAL — pure logic for the IO-list ↔ plant-CSV bridge. Legacy IO-list format:
    /// "PC Board Rack.Slot.Card" column holds node.offset.card; the CSV holds Node/Offset/Bit and
    /// BOARD NUMBER (= card). The physical join key is node.offset.card.bit; the displayed address
    /// is I|O.node.offset.bit. Remove together with src/ExcelPowerTools/Experimental and the tests.
    /// </summary>
    public static class IoListLogic
    {
        /// <summary>
        /// Splits "R1.S1.C2" (or bare "80.3.2") into rack "1", slot "1", card "2" — the card is the LAST
        /// component; letter prefixes (R/S/C) are stripped, digits compared numerically.
        /// </summary>
        public static bool TryParseRackSlotCard(string s, out string rack, out string slot, out string card)
        {
            rack = slot = card = null;
            if (string.IsNullOrWhiteSpace(s)) return false;

            string[] parts = s.Trim().Split('.');
            if (parts.Length < 3) return false;

            card = DigitsOf(parts[parts.Length - 1]);
            slot = DigitsOf(parts[parts.Length - 2]);
            rack = DigitsOf(string.Join(".", parts, 0, parts.Length - 2));
            return rack.Length > 0 && slot.Length > 0 && card.Length > 0;
        }

        /// <summary>
        /// Parses a full IO address "I.0080.3.0" into direction 'I'/'O' and numerically-normalized
        /// node/offset/bit ("0080" → "80"). This is the format the legacy IO list stores per row.
        /// </summary>
        public static bool TryParseAddress(string s, out char direction, out string node, out string offset, out string bit)
        {
            direction = '\0'; node = offset = bit = null;
            if (string.IsNullOrWhiteSpace(s)) return false;

            string[] parts = s.Trim().Split('.');
            if (parts.Length != 4 || parts[0].Length != 1) return false;

            char d = char.ToUpperInvariant(parts[0][0]);
            if (d != 'I' && d != 'O') return false;
            if (!AllDigits(parts[1]) || !AllDigits(parts[2]) || !AllDigits(parts[3])) return false;

            direction = d;
            node = Norm(parts[1]);
            offset = Norm(parts[2]);
            bit = Norm(parts[3]);
            return true;
        }

        /// <summary>The digits of a component ("C2" → "2"), numerically normalized ("02" → "2").</summary>
        public static string DigitsOf(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
                if (c >= '0' && c <= '9') sb.Append(c);
            return Norm(sb.ToString());
        }

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
        /// Direction is required — I.0080.3.0 and O.0080.3.0 coexist on the same card in real lists.
        /// </summary>
        public static string JoinKey(char direction, string node, string offset, string card, string bit) =>
            $"{char.ToUpperInvariant(direction)}.{Norm(node)}.{Norm(offset)}.{Norm(card)}.{Norm(bit)}";

        /// <summary>Numeric-normalizes a component ("07"→"7"); non-numeric values pass through trimmed.</summary>
        public static string Norm(string s)
        {
            s = (s ?? "").Trim();
            return long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out long v)
                ? v.ToString(CultureInfo.InvariantCulture)
                : s;
        }

        /// <summary>
        /// True for virtual signals (Input/Output Type 5 or 6): not read from hardware I/O, and their
        /// node/offset/bit/board are frequently stale copy-paste values — excluded from address matching.
        /// </summary>
        public static bool IsVirtualType(string typeValue)
        {
            string t = Norm(typeValue);
            return t == "5" || t == "6";
        }

        /// <summary>'I' or 'O' from a file/sheet name ("Plant_Outputs.csv" → 'O'); '\0' when undecidable.</summary>
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

        public static string NormalizeHeader(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c) || c == '.') sb.Append(char.ToLowerInvariant(c));
                // whitespace, newlines, underscores, punctuation → dropped
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
