using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The upstream-only alarm synthesizers. Belt-photocell alarms come from Sensor.csv
    /// (PREV_BELT_ID assigns each photocell to a belt, OFFSET_ON_BELT_Y positions it; photocells
    /// sort by absolute position and claim bits 50, 51, … of their belt's Alarm2 word), matched
    /// back to the IO list's PHT_* rows for address and electrical identity. Strobe-device alarms
    /// come from StrobeDevice.csv (TYPE_ID / REFERENCE_BELT_ID / REFERENCE_CABINET_ID) expanded
    /// through the editable templates. Generated entries carry the legacy markers ("AP", "ASD")
    /// and implicitly skip the electrical-label checks.
    /// </summary>
    public static class SwpUpsGenerators
    {
        // Sensor.csv / StrobeDevice.csv column layout (semicolon):
        private const int ColKey = 1, ColTag = 2, ColSystem = 3, ColDevice = 5, ColValue = 7;

        private sealed class Belt
        {
            public int System;
            public int Index;
            public List<int> Photocells = new List<int>();       // in discovery order, then sorted
            public Dictionary<int, long> Positions = new Dictionary<int, long>();
        }

        // ---- photocells -------------------------------------------------------------------

        public static List<DiagnosticEntry> GeneratePhotocellAlarms(IEnumerable<string> sensorCsvLines,
            IReadOnlyList<IoRowData> ioRows, List<SwpLogEntry> log)
        {
            var entries = new List<DiagnosticEntry>();
            if (sensorCsvLines == null) return entries;

            var belts = new List<Belt>();
            var beltByKey = new Dictionary<string, Belt>();
            var pendingPositions = new List<string[]>();

            foreach (string line in sensorCsvLines)
            {
                string[] fields = (line ?? "").Split(';');
                if (!IsActive(fields)) continue;

                string tag = Field(fields, ColTag);
                if (tag == "PREV_BELT_ID")
                {
                    int system = IntOf(Field(fields, ColSystem));
                    int beltIndex = IntOf(Field(fields, ColValue));
                    int photocell = IntOf(Field(fields, ColDevice));
                    if (system <= 0 || beltIndex < 0 || photocell <= 0) continue;

                    string key = system + "_" + beltIndex;
                    if (!beltByKey.TryGetValue(key, out Belt belt))
                    {
                        beltByKey[key] = belt = new Belt { System = system, Index = beltIndex };
                        belts.Add(belt);
                    }
                    if (!belt.Photocells.Contains(photocell)) belt.Photocells.Add(photocell);
                }
                else if (tag == "OFFSET_ON_BELT_Y")
                {
                    if (!StorePosition(belts, fields)) pendingPositions.Add(fields);
                }
            }

            // Photocells whose belt line came after their position line get a second chance.
            foreach (string[] fields in pendingPositions)
                if (!StorePosition(belts, fields))
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "no belt found for sensor " + Field(fields, ColSystem) + "_" + Field(fields, ColDevice)));

            // The IO list's PHT_* rows, findable by exact mnemonic part.
            var photocellRows = new Dictionary<string, IoRowData>();
            foreach (IoRowData row in ioRows ?? new List<IoRowData>())
            {
                if (row == null || (row.MnemonicCell ?? "").IndexOf("PHT", StringComparison.Ordinal) < 0) continue;
                foreach (string part in row.MnemonicCell.Split('/'))
                {
                    string name = part.Trim();
                    if (name.Length > 0 && !photocellRows.ContainsKey(name)) photocellRows[name] = row;
                }
            }

            foreach (Belt belt in belts)
            {
                var ordered = belt.Photocells
                    .OrderBy(p => belt.Positions.TryGetValue(p, out long pos) ? pos : 0L)
                    .ToList();

                int bit = 50;
                foreach (int photocell in ordered)
                {
                    string mnemonic = "ALLB_" + belt.System + "_" +
                        belt.Index.ToString("000", CultureInfo.InvariantCulture) + "_" +
                        bit.ToString("00", CultureInfo.InvariantCulture);
                    string photocellName = "PHT_" + belt.System + "_" +
                        photocell.ToString("000", CultureInfo.InvariantCulture) + "_000";

                    var entry = new DiagnosticEntry
                    {
                        Mnemonic = mnemonic,
                        RawMnemonic = mnemonic,
                        GeneratedBy = "AP",
                        TsRef = "TS_031",
                        CardSheet = "(generated)",
                    };

                    int sensorNumber = bit - 49;
                    if (photocellRows.TryGetValue(photocellName, out IoRowData source))
                    {
                        if (IoListLogic.TryParseAddress((source.AddressCell ?? "").Trim(),
                            out char direction, out string node, out string offset, out string addressBit) &&
                            direction == 'I')
                        {
                            entry.Node = node;
                            entry.Offset = offset;
                            entry.Bit = addressBit;
                        }
                        entry.CardSheet = source.CardSheet;
                        entry.SourceRow = source.SourceRow;
                        entry.FUnit = (source.FUnit ?? "").Trim();
                        entry.Location = (source.Location ?? "").Trim();
                        entry.Device = (source.Device ?? "").Trim();
                        // The legacy synthesized J text — the description pipeline runs over it.
                        entry.Desc1A = "Sensor Fault #" + sensorNumber + " " + entry.Location + entry.Device;
                    }
                    else
                    {
                        entry.Desc1A = "Sensor Fault #" + sensorNumber;
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            photocellName + " not found in the IO list — the belt alarm has no address",
                            mnemonic: mnemonic));
                    }

                    entries.Add(entry);
                    bit++;
                }
            }
            return entries;
        }

        private static bool StorePosition(List<Belt> belts, string[] fields)
        {
            int system = IntOf(Field(fields, ColSystem));
            int photocell = IntOf(Field(fields, ColDevice));
            foreach (Belt belt in belts)
            {
                if (belt.System != system || !belt.Photocells.Contains(photocell)) continue;
                belt.Positions[photocell] = AbsolutePosition(Field(fields, ColValue));
                return true;
            }
            return false;
        }

        /// <summary>
        /// PB/PE/NB/NE offsets to one comparable axis (the legacy half-belt arithmetic): the
        /// previous belt's begin, its end (= next begin), and the next belt's end.
        /// </summary>
        internal static long AbsolutePosition(string offsetParameter)
        {
            string s = (offsetParameter ?? "").Trim();
            if (s.Length < 3) return 0;
            string reference = s.Substring(0, 2).ToUpperInvariant();
            if (!long.TryParse(s.Substring(2), NumberStyles.Integer, CultureInfo.InvariantCulture, out long offset))
                return 0;
            const long half = 10000;
            switch (reference)
            {
                case "PB": return half + offset;
                case "PE":
                case "NB": return half * 2 + offset;
                case "NE": return half * 4 + offset;
                default: return 0;
            }
        }

        // ---- strobe devices ---------------------------------------------------------------

        public static List<DiagnosticEntry> GenerateStrobeAlarms(IEnumerable<string> strobeCsvLines,
            IReadOnlyList<SwpStrobeTemplate> templates, List<SwpLogEntry> log)
        {
            var entries = new List<DiagnosticEntry>();
            if (strobeCsvLines == null || templates == null || templates.Count == 0) return entries;

            var devices = new List<(int System, int Index, int TypeId, int Belt, int Cabinet)>();
            var deviceState = new Dictionary<string, int[]>();   // key → [typeId, belt, cabinet]
            var deviceOrder = new List<(string Key, int System, int Index)>();

            foreach (string line in strobeCsvLines)
            {
                string[] fields = (line ?? "").Split(';');
                if (!IsActive(fields)) continue;
                int system = IntOf(Field(fields, ColSystem));
                int device = IntOf(Field(fields, ColDevice));
                if (system <= 0) continue;

                string key = system + "_" + device;
                if (!deviceState.TryGetValue(key, out int[] state))
                {
                    deviceState[key] = state = new int[] { 0, 0, 0 };
                    deviceOrder.Add((key, system, device));
                }
                switch (Field(fields, ColTag))
                {
                    case "TYPE_ID": state[0] = IntOf(Field(fields, ColValue)); break;
                    case "REFERENCE_BELT_ID": state[1] = IntOf(Field(fields, ColValue)); break;
                    case "REFERENCE_CABINET_ID": state[2] = IntOf(Field(fields, ColValue)); break;
                }
            }
            foreach (var (key, system, index) in deviceOrder)
            {
                int[] state = deviceState[key];
                devices.Add((system, index, state[0], state[1], state[2]));
            }

            var templateByType = new Dictionary<string, SwpStrobeTemplate>();
            foreach (SwpStrobeTemplate template in templates)
                if (!templateByType.ContainsKey(template.TypeId)) templateByType[template.TypeId] = template;

            foreach (var device in devices)
            {
                if (!templateByType.TryGetValue(
                        device.TypeId.ToString(CultureInfo.InvariantCulture), out SwpStrobeTemplate template))
                {
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "strobe device type " + device.TypeId + " (system " + device.System +
                        ", device " + device.Index + ") has no template"));
                    continue;
                }

                for (int i = 0; i < template.Mnemonics.Count; i++)
                {
                    string mnemonic = template.Mnemonics[i]
                        .Replace("%sys%", device.System.ToString(CultureInfo.InvariantCulture))
                        .Replace("%belt%", device.Belt.ToString("000", CultureInfo.InvariantCulture))
                        .Replace("%cabinet%", device.Cabinet.ToString("00", CultureInfo.InvariantCulture));
                    entries.Add(new DiagnosticEntry
                    {
                        Mnemonic = mnemonic,
                        RawMnemonic = mnemonic,
                        GeneratedBy = "ASD",
                        CardSheet = "(generated)",
                        Desc1A = i < template.Descriptions.Count ? template.Descriptions[i] : "",
                    });
                }
            }
            return entries;
        }

        private static bool IsActive(string[] fields)
        {
            string key = Field(fields, ColKey);
            return (key == "#" || key == "@") && IntOf(Field(fields, ColSystem)) > 0;
        }

        private static string Field(string[] fields, int index) =>
            fields != null && index < fields.Length ? (fields[index] ?? "").Trim() : "";

        private static int IntOf(string s) =>
            int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : -1;
    }
}
