using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// One row of the index-rules table: how a diagnosable object's indexes shift between the
    /// three worlds a mnemonic lives in. The columns mirror the official "SWP and diagnostic
    /// configuration rules" tables (§4.9 sorter, §5.2 induct line, §7.10 upstream) verbatim, so
    /// the sheet can be compared against the document side by side.
    /// </summary>
    public sealed class SwpIndexRule
    {
        public SwpPlantType PlantType = SwpPlantType.Msc;
        public string ObjectKey = "";      // CABINET / GROUP_CABINET / BELT / ZONE / CHUTE / …
        public int MnemonicStart;          // "Mnemonic start index (realtime index)"
        public int SwpEntryStart;          // "SWP start index (Table entry)"
        public int OpcTagStart;            // "OPC TAG start index" — the alarm-strings export
        public bool AlarmCheckEnabled;     // whether the alarm-strings leg checks this object

        /// <summary>Added to the mnemonic's raw index to reach the SWP table entry.</summary>
        public int EntryOffset => SwpEntryStart - MnemonicStart;

        /// <summary>Added to the mnemonic's raw index to reach the exported OPC-tag index.</summary>
        public int AlarmOffset => OpcTagStart - MnemonicStart;
    }

    /// <summary>
    /// The living index-rules table, seeded from the official document and editable on the
    /// _Offsets workspace sheet (persisted globally, like the language table). Lookups fall back
    /// to a neutral per-plant rule — mnemonics start at 0 for MSC/SAFETY and 1 for ISPC/UPS,
    /// table entries at 0, no alarm export shift, alarm check off — which reproduces the legacy
    /// hardcoded behavior for any object the table does not name.
    /// </summary>
    public sealed class SwpIndexRules
    {
        public const string GroupCabinetKey = "GROUP_CABINET";

        private readonly Dictionary<string, SwpIndexRule> _byKey =
            new Dictionary<string, SwpIndexRule>(StringComparer.OrdinalIgnoreCase);

        public List<SwpIndexRule> Rules { get; } = new List<SwpIndexRule>();

        public void Add(SwpIndexRule rule)
        {
            if (rule == null || rule.ObjectKey.Length == 0) return;
            string key = KeyOf(rule.PlantType, rule.ObjectKey);
            if (_byKey.ContainsKey(key)) return;   // first row wins, like every workspace table
            _byKey[key] = rule;
            Rules.Add(rule);
        }

        public SwpIndexRule RuleFor(SwpPlantType plantType, string objectKey)
        {
            if (_byKey.TryGetValue(KeyOf(plantType, objectKey), out SwpIndexRule rule)) return rule;
            int mnemonicStart = plantType == SwpPlantType.Ispc || plantType == SwpPlantType.Ups ? 1 : 0;
            return new SwpIndexRule
            {
                PlantType = plantType,
                ObjectKey = objectKey ?? "",
                MnemonicStart = mnemonicStart,
                SwpEntryStart = 0,
                OpcTagStart = mnemonicStart,
                AlarmCheckEnabled = false,
            };
        }

        public int EntryOffset(SwpPlantType plantType, string objectKey) =>
            RuleFor(plantType, objectKey).EntryOffset;

        public int AlarmOffset(SwpPlantType plantType, string objectKey) =>
            RuleFor(plantType, objectKey).AlarmOffset;

        public bool AlarmCheckEnabled(SwpPlantType plantType, string objectKey) =>
            RuleFor(plantType, objectKey).AlarmCheckEnabled;

        /// <summary>
        /// The rules key of a mapped mnemonic. ISPC group cabinets (line 0) are a distinct
        /// diagnosable object from line cabinets — their exports shift differently.
        /// </summary>
        public static string ObjectKeyFor(SwpPlantType plantType, SwpMappedMnemonic mapped)
        {
            if (mapped == null) return "";
            if (plantType == SwpPlantType.Ispc && mapped.ObjectType == "CABINET" && mapped.ReferenceIndex == 0)
                return GroupCabinetKey;
            return mapped.ObjectType;
        }

        /// <summary>
        /// The official tables of "SWP and diagnostic configuration rules" v93: §4.9 (sorter),
        /// §5.2 (induct line), §7.10 (upstream); SAFETY per §9.2 with the 1:1 zero-based alarm
        /// matching of the direct-OPC-UA safety PLC. The sorter CABINET/"all other devices" OPC
        /// shift is the DDSNG ≥ 9.4.2 numbering. Alarm checks are enabled only where the
        /// exported-alarms parser knows the tag family (cabinet/belt/zone).
        /// </summary>
        public static SwpIndexRules Official()
        {
            var rules = new SwpIndexRules();
            void Row(SwpPlantType plant, string objectKey, int mnemonicStart, int entryStart,
                int opcStart, bool check) =>
                rules.Add(new SwpIndexRule
                {
                    PlantType = plant,
                    ObjectKey = objectKey,
                    MnemonicStart = mnemonicStart,
                    SwpEntryStart = entryStart,
                    OpcTagStart = opcStart,
                    AlarmCheckEnabled = check,
                });

            Row(SwpPlantType.Msc, "CABINET", 0, 0, 1, true);
            Row(SwpPlantType.Msc, "CHUTE", 0, 0, 0, false);          // mnemonic = chute id, remapped
            Row(SwpPlantType.Msc, "SCANNER", 0, 0, 1, false);
            Row(SwpPlantType.Msc, "CONTROL_ZONE", 0, 0, 1, false);
            Row(SwpPlantType.Msc, "CHECK_POSITION", 0, 0, 1, false);
            Row(SwpPlantType.Msc, "CELL_TEST", 0, 0, 1, false);
            Row(SwpPlantType.Msc, "AIR_BLOWER", 0, 0, 0, false);
            Row(SwpPlantType.Msc, "ITEM_LOCATOR", 0, 0, 1, false);

            Row(SwpPlantType.Ispc, "CABINET", 1, 0, 1, true);
            Row(SwpPlantType.Ispc, GroupCabinetKey, 1, 0, 0, true);
            Row(SwpPlantType.Ispc, "ZONE", 1, 0, 1, false);          // no induct zone tags parsed
            Row(SwpPlantType.Ispc, "BELT", 1, 0, 1, false);          // no induct belt tags parsed

            Row(SwpPlantType.Ups, "CABINET", 1, 0, 0, true);
            Row(SwpPlantType.Ups, "BELT", 1, 0, 1, true);
            Row(SwpPlantType.Ups, "ZONE", 1, 0, 1, true);
            Row(SwpPlantType.Ups, "STROBE_DEVICE", 1, 0, 1, false);

            Row(SwpPlantType.Safety, "CABINET", 0, 0, 0, true);
            return rules;
        }

        // ---- persistence ("Msc|CABINET|0|0|1|x") --------------------------------------------

        public List<string> ToLines()
        {
            var lines = new List<string>();
            foreach (SwpIndexRule rule in Rules)
                lines.Add(rule.PlantType + "|" + rule.ObjectKey + "|" +
                    rule.MnemonicStart.ToString(CultureInfo.InvariantCulture) + "|" +
                    rule.SwpEntryStart.ToString(CultureInfo.InvariantCulture) + "|" +
                    rule.OpcTagStart.ToString(CultureInfo.InvariantCulture) + "|" +
                    (rule.AlarmCheckEnabled ? "x" : "-"));
            return lines;
        }

        /// <summary>Tolerant reader: malformed lines are skipped, an empty result yields null.</summary>
        public static SwpIndexRules FromLines(IEnumerable<string> lines)
        {
            if (lines == null) return null;
            var rules = new SwpIndexRules();
            foreach (string line in lines)
            {
                string[] parts = (line ?? "").Split('|');
                if (parts.Length < 6) continue;
                if (!Enum.TryParse(parts[0].Trim(), true, out SwpPlantType plant)) continue;
                string objectKey = parts[1].Trim();
                if (objectKey.Length == 0) continue;
                if (!TryInt(parts[2], out int mnemonicStart) ||
                    !TryInt(parts[3], out int entryStart) ||
                    !TryInt(parts[4], out int opcStart)) continue;
                rules.Add(new SwpIndexRule
                {
                    PlantType = plant,
                    ObjectKey = objectKey,
                    MnemonicStart = mnemonicStart,
                    SwpEntryStart = entryStart,
                    OpcTagStart = opcStart,
                    AlarmCheckEnabled = parts[5].Trim().Equals("x", StringComparison.OrdinalIgnoreCase),
                });
            }
            return rules.Rules.Count > 0 ? rules : null;
        }

        private static string KeyOf(SwpPlantType plantType, string objectKey) =>
            plantType + "|" + (objectKey ?? "");

        private static bool TryInt(string s, out int value) =>
            int.TryParse((s ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }
}
