using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The shipped mnemonic map: every shape the hardcoded <see cref="SwpAddressMap"/> understood,
    /// restated as rules. Two things it is NOT allowed to be — a rewrite of the behaviour, or a
    /// second source of truth. It is pinned to the old mapper by a parity test over the real IO
    /// lists, and it is the file a user exports, edits and shares.
    ///
    /// Where the legacy code MUTATED a mnemonic to reach an existing branch, the rule states the
    /// destination outright: EMERGENCY is its own row targeting DB_CABINETS_DIAG rather than a
    /// rename into CABINETAL, so the signal keeps the name the plant CSVs know it by.
    /// </summary>
    public static class SwpMnemonicMapDefaults
    {
        /// <summary>MSC CLASSDIAG type digits → (sheet, tag). Chutes alone remap their index
        /// through Destination_Parameters.</summary>
        private static readonly (int Type, string Sheet, string Tag)[] ClassDiag =
        {
            (62, "CHUTE", "DB_CHUTE"),
            (64, "SCANNER", "DB_SCANNER"),
            (66, "CONTROL_ZONE", "DB_CONTROLZONE"),
            (69, "CHECK_POSITION", "DB_CHECKPOSITION"),
            (73, "CELL_TEST", "DB_CELLTEST"),
            (112, "AIR_BLOWER", "DB_AIR_BLOWER"),
            (113, "ITEM_LOCATOR", "DB_ITEM_LOCATOR"),
        };

        /// <summary>
        /// Built once and then ROUND-TRIPPED through the text format, so the shipped table is
        /// parsed by exactly the path a user's edited file takes — a built-in that the format
        /// could not express, or that the parser read back differently, fails here rather than
        /// in the field.
        /// </summary>
        public static SwpMnemonicMap Official() =>
            SwpMnemonicMap.Parse(OfficialLines(), null);

        public static List<string> OfficialLines()
        {
            var draft = new SwpMnemonicMap();
            AddMscFamily(draft, SwpPlantType.Msc, cabinetsOnly: false);
            AddMscFamily(draft, SwpPlantType.Safety, cabinetsOnly: true);
            AddIspcFamily(draft);
            AddUpsFamily(draft);
            return draft.ToLines();
        }

        // ---- MSC / SAFETY -------------------------------------------------------------------
        // SAFETY rides the sorter's CABINET shapes by contract (rules §9.2): a safety SWP is the
        // standard configuration with the cabinet sheet only, and the safety tooling writes its
        // mnemonic column in sorter syntax on purpose. It takes the cabinet rows and NOTHING else:
        // the CLASSDIAG classes are sorter objects (chutes, scanners, …) a safety SWP has no sheet
        // for, so a CLASSDIAG mnemonic in a safety list stays unclaimed instead of being sent to a
        // sheet that does not exist. _Offsets is keyed by plant already (SAFETY: CABINET, 1:1
        // zero-based — the FVT safety SCADA's export starts at S1.CABINETAL00000).

        private static void AddMscFamily(SwpMnemonicMap map, SwpPlantType plant, bool cabinetsOnly)
        {
            // CABINET AL 001 00 — the legacy prefix(7) + AL/WR + 3 device + 2 bit = 14.
            Add(map, plant, 10, "CABINET",
                @"^CABINET(?<aw>AL|WR)(?<i>\d{3})(?<b>\d{2})$",
                "CABINET", "DB_CABINETS_DIAG", "auto", "{i}", "{b}",
                "the sorter cabinet alarms and warnings");

            // Was a RENAME into CABINETAL — the prefixes are both nine characters, which is the
            // only reason the trick worked. Now it is a destination, and the mnemonic stays the
            // name Plant_Inputs knows: EMERGENCY02000, not CABINETAL02000.
            Add(map, plant, 11, "CABINET",
                @"^EMERGENCY(?<i>\d{3})(?<b>\d{2})$",
                "CABINET", "DB_CABINETS_DIAG", "Alarm", "{i}", "{b}",
                "emergency-circuit alarms; the legacy tool renamed these to CABINETAL");

            // Commands, not diagnostics — recognized on purpose so they are not reported as
            // unknown shapes (was a hardcoded exclusion in the filter).
            Add(map, plant, 12, "", @"^EMERGENCY_(ON|RST)$", "", "", "", "", "",
                "commands, deliberately not written to the SWP");

            if (cabinetsOnly) return;   // SAFETY: no CLASSDIAG objects on a cabinet-only SWP

            int order = 20;
            foreach ((int type, string sheet, string tag) in ClassDiag)
            {
                // CLASSDIAG AL tttt ddddd bb — prefix(9) + AL/WR + 4 type + 5 device + 2 bit = 22.
                string index = type == 62 ? "map(Destination_Parameters,{i})" : "{i}";
                Add(map, plant, order, sheet,
                    @"^CLASSDIAG(?<aw>AL|WR)" + type.ToString("0000") + @"(?<i>\d{5})(?<b>\d{2})$",
                    sheet, tag, "auto", index, "{b}",
                    type == 62 ? "chute index remapped through Destination_Parameters" : "");
                order++;
            }
        }

        // ---- ISPC ---------------------------------------------------------------------------
        // The line digit selects the sheet: 1..n are the induct lines IL01…, 0 is the group
        // cabinet ILGRP. In the old code that was an if inside the branch AND a second _Offsets
        // key; here it is simply two rules.

        private static void AddIspcFamily(SwpMnemonicMap map)
        {
            Add(map, SwpPlantType.Ispc, 10, "CABINET",
                @"^CABINET(?<aw>AL|WR)(?<line>[1-9])(?<i>\d{3})(?<b>\d{2})$",
                "IL{firstLineId+line-1:00}_CABINET", "", "auto", "{i}", "{b}",
                "induct-line cabinets");

            Add(map, SwpPlantType.Ispc, 11, SwpIndexRules.GroupCabinetKey,
                // The 0 is CAPTURED, not just matched: the alarm-strings leg reports the line a
                // row belongs to, and the group cabinet's line is 0 — not "absent".
                @"^CABINET(?<aw>AL|WR)(?<line>0)(?<i>\d{3})(?<b>\d{2})$",
                "ILGRP{groupIndex:00}_CABINET", "", "auto", "{i}", "{b}",
                "line 0 is the group cabinet — its own _Offsets key");

            Add(map, SwpPlantType.Ispc, 20, "BELT",
                @"^BELT(?<aw>AL|WR)(?<line>\d)(?<i>\d{2})(?<b>\d{2})$",
                "IL{firstLineId+line-1:00}_BELT", "", "auto", "{i}", "{b}", "");

            Add(map, SwpPlantType.Ispc, 30, "ZONE",
                @"^ZONE(?<aw>AL|WR)(?<line>\d)(?<i>\d)(?<b>\d{2})$",
                "IL{firstLineId+line-1:00}_ZONE", "", "auto", "{i}", "{b}", "");
        }

        // ---- UPS ----------------------------------------------------------------------------
        // Prefix shape A…/W… chooses alarm or warning, and its LAST letter the object: Q cabinet,
        // B belt, Z zone. The legacy mapper split on '_' and read parts 1..3.

        private static void AddUpsFamily(SwpMnemonicMap map)
        {
            var objects = new[] { ("Q", "CABINET"), ("B", "BELT"), ("Z", "ZONE") };
            int order = 10;
            foreach ((string letter, string objectKey) in objects)
            {
                string sheet = "UP{firstLineId+sys-1:00}_" + objectKey;

                Add(map, SwpPlantType.Ups, order, objectKey,
                    @"^(?<aw>[AW])[A-Z]*" + letter + @"_(?<sys>\d+)_(?<i>\d+)_(?<b>\d+)$",
                    sheet, "", "auto", "{i}", "{b}", "");

                // AND / OR groups: many digital inputs drive ONE diagnostic bit. The string is
                // built from the _0 member; the rest are recognized and go nowhere. The legacy
                // tool got here by rewriting ALLQ_AND_ to ALLQ_ and dropping the last two
                // characters — the same outcome, spelled out.
                Add(map, SwpPlantType.Ups, order + 1, objectKey,
                    @"^(?<aw>[AW])[A-Z]*" + letter + @"_(?:AND|OR)_(?<sys>\d+)_(?<i>\d+)_(?<b>\d+)_0$",
                    sheet, "", "auto", "{i}", "{b}",
                    "first member of an AND/OR group — carries the whole group");

                Add(map, SwpPlantType.Ups, order + 2, "",
                    @"^(?<aw>[AW])[A-Z]*" + letter + @"_(?:AND|OR)_\d+_\d+_\d+_[1-9]\d*$",
                    "", "", "", "", "",
                    "further members of an AND/OR group — the _0 member already carries it");
                order += 10;
            }

            // EMERI_ folded into ALLQ_ in the legacy rewrites; it is a cabinet shape of its own.
            Add(map, SwpPlantType.Ups, 90, "CABINET",
                @"^EMERI_(?<sys>\d+)_(?<i>\d+)_(?<b>\d+)$",
                "UP{firstLineId+sys-1:00}_CABINET", "", "Alarm", "{i}", "{b}",
                "the legacy tool renamed these to ALLQ_");

            // ---- producers ------------------------------------------------------------------
            // These do not claim an incoming mnemonic — they CREATE mnemonics, which then flow
            // through the Map rules above like any other. They live here so the table is the one
            // place that answers "what happens to mnemonics", and so a build can be told to stop
            // running one by clearing its tick.

            // Sensor.csv assigns each photocell to a belt (PREV_BELT_ID) and positions it along it
            // (OFFSET_ON_BELT_Y). Photocells sort by absolute position and claim consecutive bits
            // from firstBit; each emitted belt alarm then seeks its PHT_* row in the IO list for
            // the address and the electrical identity. The belt assignment and the ordering are
            // bespoke and stay in code; everything a person would want to turn is here.
            Generate(map, SwpPlantType.Ups, 200, "Photocells", "BELT",
                seeks: "PHT_{sys}_{n:000}_000",
                emits: "ALLB_{sys}_{belt:000}_{bit:00}",
                parameters: "source=Sensor;firstBit=50;mark=AP;tsRef=TS_031;desc=Sensor Fault #{sensor}",
                note: "belt-photocell alarms, ordered by position along the belt");

            // StrobeDevice.csv (TYPE_ID / REFERENCE_BELT_ID / REFERENCE_CABINET_ID) expanded
            // through the editable strobe templates, which keep their own sheet.
            Generate(map, SwpPlantType.Ups, 210, "Strobes", "CABINET",
                seeks: "",
                emits: "{template}",
                parameters: "source=StrobeDevice;mark=ASD",
                note: "strobe-device alarms from the _Strobes templates");
        }

        private static void Generate(SwpMnemonicMap map, SwpPlantType plant, int order,
            string generator, string objectKey, string seeks, string emits, string parameters,
            string note)
        {
            map.Rules.Add(new SwpMnemonicRule
            {
                PlantType = plant,
                Order = order,
                Enabled = true,
                Kind = SwpRuleKind.Generate,
                Generator = generator,
                ObjectKey = objectKey,
                Match = seeks,
                Emit = emits,
                Params = parameters,
                Note = note,
            });
        }

        private static void Add(SwpMnemonicMap map, SwpPlantType plant, int order, string objectKey,
            string match, string sheet, string tag, string word, string index, string bit, string note)
        {
            var rule = new SwpMnemonicRule
            {
                PlantType = plant,
                Order = order,
                Enabled = true,
                ObjectKey = objectKey,
                Match = match,
                Sheet = sheet,
                Tag = tag,
                Word = word.Length == 0 ? "auto" : word,
                Index = index,
                Bit = bit,
                Note = note,
            };
            map.Rules.Add(rule);
        }
    }
}
