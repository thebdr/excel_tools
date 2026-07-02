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
        /// <summary>Splits "80.3.2" into node "80", offset "3", card "2" (card = LAST component).</summary>
        public static bool TryParseRackSlotCard(string s, out string node, out string offset, out string card)
        {
            node = offset = card = null;
            if (string.IsNullOrWhiteSpace(s)) return false;

            string[] parts = s.Trim().Split('.');
            if (parts.Length < 3) return false;

            card = Norm(parts[parts.Length - 1]);
            offset = Norm(parts[parts.Length - 2]);
            node = Norm(string.Join(".", parts, 0, parts.Length - 2));   // tolerate a dotted node
            return node.Length > 0 && offset.Length > 0 && card.Length > 0;
        }

        /// <summary>The display address: "I.80.3.0" ('\0' direction → no prefix).</summary>
        public static string BuildAddress(char direction, string node, string offset, string bit)
        {
            string body = $"{Norm(node)}.{Norm(offset)}.{Norm(bit)}";
            return direction == '\0' ? body : direction + "." + body;
        }

        /// <summary>The physical join key: node.offset.card.bit, numerically normalized ("07" == "7").</summary>
        public static string JoinKey(string node, string offset, string card, string bit) =>
            $"{Norm(node)}.{Norm(offset)}.{Norm(card)}.{Norm(bit)}";

        /// <summary>Numeric-normalizes a component ("07"→"7"); non-numeric values pass through trimmed.</summary>
        public static string Norm(string s)
        {
            s = (s ?? "").Trim();
            return long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out long v)
                ? v.ToString(CultureInfo.InvariantCulture)
                : s;
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
