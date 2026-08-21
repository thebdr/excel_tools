using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// AS-I identities as the IO list writes them and as the realtime layer numbers them. A slave
    /// row carries Slot = the gateway line ("1M" / "2M" — one AS-I gateway drives two lines) and
    /// ID = the slave address ("1A" … "31A", "1B" … "31B"); the bridge composes them into
    /// "1M.24A" (an empty slot is line 1). The realtime sub-device number — column 5 of
    /// BUS_ExtendedDiagnostic.csv, "IODriver ID" on the IO list's ASI Addressing sheet — is
    /// fixed by the gateway's image layout: line 1 A-slaves 1..31, line 1 B-slaves 33..63,
    /// line 2 A-slaves 65..95, line 2 B-slaves 97..127 (0/32/64/96 are the master flags).
    /// Kept in code because not every IO list ships the addressing sheet.
    /// </summary>
    public static class SwpAsiAddressing
    {
        public const string DefaultSlot = "1M";

        /// <summary>"1M" + "24A" → "1M.24A"; an empty slot defaults to line 1. "" when the id
        /// is not a slave address or the slot is not a gateway line.</summary>
        public static string Compose(string slot, string slaveId)
        {
            string id = (slaveId ?? "").Trim().ToUpperInvariant();
            if (!TryParseSlave(id, out _, out _)) return "";
            string line = NormalizeSlot(slot);
            return line.Length == 0 ? "" : line + "." + id;
        }

        /// <summary>"1M" / "2M" (case-insensitive; empty → 1M); "" for anything else.</summary>
        public static string NormalizeSlot(string slot)
        {
            string s = (slot ?? "").Trim().ToUpperInvariant();
            if (s.Length == 0) return DefaultSlot;
            return s == "1M" || s == "2M" ? s : "";
        }

        /// <summary>The realtime sub-device number of a composed node ("1M.24A" → 24,
        /// "2M.1B" → 97); -1 when it does not parse.</summary>
        public static int SubDevice(string asiNode)
        {
            if (!TryParse(asiNode, out int line, out int slave, out bool lineB)) return -1;
            return (line - 1) * 64 + (lineB ? 32 : 0) + slave;
        }

        public static bool TryParse(string asiNode, out int line, out int slave, out bool lineB)
        {
            line = 0;
            slave = 0;
            lineB = false;
            string s = (asiNode ?? "").Trim().ToUpperInvariant();
            int dot = s.IndexOf('.');
            if (dot <= 0) return false;
            string slot = s.Substring(0, dot);
            if (slot != "1M" && slot != "2M") return false;
            line = slot[0] - '0';
            return TryParseSlave(s.Substring(dot + 1), out slave, out lineB);
        }

        /// <summary>"24A" → 24 on line A; valid addresses are 1..31 with an A or B suffix.</summary>
        public static bool TryParseSlave(string id, out int slave, out bool lineB)
        {
            slave = 0;
            lineB = false;
            string s = (id ?? "").Trim().ToUpperInvariant();
            if (s.Length < 2 || s.Length > 3) return false;
            char suffix = s[s.Length - 1];
            if (suffix != 'A' && suffix != 'B') return false;
            string digits = s.Substring(0, s.Length - 1);
            foreach (char c in digits)
                if (c < '0' || c > '9') return false;
            slave = int.Parse(digits, CultureInfo.InvariantCulture);
            if (slave < 1 || slave > 31) return false;
            lineB = suffix == 'B';
            return true;
        }
    }
}
