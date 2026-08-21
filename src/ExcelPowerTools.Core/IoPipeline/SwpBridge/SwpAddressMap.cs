using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The mnemonic → SWP coordinate mapping, ported from the legacy MakeSwpAddressFromMnemonic.
    /// A mnemonic encodes object type, device index and bit; the SWP splits each device's bits
    /// into 32-bit words, so the tag type carries "Alarm1"/"Alarm2"/"Warning1"/"Warning2" with
    /// word = ceil((bit+1)/32) and the raw bit number continues across words (Alarm2 rows hold
    /// D = 32…63). Error strings keep the legacy "#Invalid …" codes for familiarity.
    /// </summary>
    public static class SwpAddressMap
    {
        /// <summary>MSC CLASSDIAG type digits → object type / DDE tag. Extend here for new types.
        /// 62–73 are the legacy set; 112/113 come from rules §3.5 (tag prefixes verified against
        /// the sheet parameters of a real SWP — DB_AIR_BLOWER and DB_ITEM_LOCATOR keep their
        /// underscores, unlike DB_CONTROLZONE and friends).</summary>
        private static readonly Dictionary<int, (string ObjectType, string TagType)> ClassDiagTypes =
            new Dictionary<int, (string, string)>
            {
                { 62, ("CHUTE", "DB_CHUTE") },
                { 64, ("SCANNER", "DB_SCANNER") },
                { 66, ("CONTROL_ZONE", "DB_CONTROLZONE") },
                { 69, ("CHECK_POSITION", "DB_CHECKPOSITION") },
                { 73, ("CELL_TEST", "DB_CELLTEST") },
                { 112, ("AIR_BLOWER", "DB_AIR_BLOWER") },
                { 113, ("ITEM_LOCATOR", "DB_ITEM_LOCATOR") },
            };

        private static readonly SwpIndexRules DefaultRules = SwpIndexRules.Official();

        /// <summary>
        /// Maps one (already filtered and rewritten) mnemonic. <paramref name="deviceIndexOverride"/>
        /// is the per-row "SWP DB Index" steering value and wins over everything;
        /// <paramref name="chuteMap"/> (MSC index → SWP device ordinal, from
        /// Destination_Parameters.csv) applies to CLASSDIAG chutes (type 0062) only.
        /// <paramref name="rules"/> supplies the per-object mnemonic→entry shifts (null = the
        /// official table).
        /// </summary>
        public static SwpMappedMnemonic Map(string mnemonic, SwpPlantType plantType,
            int firstLineId = 1, int groupIndex = 1, int? deviceIndexOverride = null,
            IReadOnlyDictionary<string, int> chuteMap = null, SwpIndexRules rules = null)
        {
            string m = (mnemonic ?? "").Trim();
            if (m.Length == 0) return Fail("#Empty Mnemonic");
            rules = rules ?? DefaultRules;

            switch (plantType)
            {
                case SwpPlantType.Msc:
                case SwpPlantType.Safety:
                    // SAFETY rides the MSC branch by contract: safety SWPs only have cabinets
                    // (§9.2) and the safety-side tooling generates its mnemonic column in the
                    // sorter (MSC) syntax on purpose — only the index table differs here
                    // (everything 1:1 zero-based).
                    return MapMsc(m, plantType, deviceIndexOverride, chuteMap, rules);
                case SwpPlantType.Ispc: return MapIspc(m, firstLineId, groupIndex, deviceIndexOverride, rules);
                case SwpPlantType.Ups: return MapUps(m, firstLineId, deviceIndexOverride, rules);
                default: return Fail("#Invalid System Type");
            }
        }

        private static SwpMappedMnemonic MapMsc(string m, SwpPlantType plantType, int? deviceIndexOverride,
            IReadOnlyDictionary<string, int> chuteMap, SwpIndexRules rules)
        {
            string objectType, tagType;
            int rawDevice, deviceIdx;

            if (m.Length == 14 && StartsWith(m, "CABINET") && AllDigits(m, 9, 5))
            {
                // CABINET AL 001 00 — prefix = first 9, device = chars 10-12, bit = last 2.
                objectType = "CABINET";
                tagType = "DB_CABINETS_DIAG";
                if (!TryInt(m.Substring(9, 3), out rawDevice)) return Fail("#Invalid Mnemonic MSC-2");
                deviceIdx = deviceIndexOverride ?? rawDevice + rules.EntryOffset(plantType, objectType);
            }
            else if (m.Length == 22 && StartsWith(m, "CLASSDIAG") && AllDigits(m, 11, 11))
            {
                // CLASSDIAG AL tttt ddddd bb — type = chars 12-15, device = 16-20, bit = last 2.
                if (!TryInt(m.Substring(11, 4), out int classType)) return Fail("#Invalid Mnemonic MSC-2");
                if (!ClassDiagTypes.TryGetValue(classType, out var entry))
                    return Fail("#Invalid or Not Yet Managed CLASSDIAG Type MSC-1");
                objectType = entry.ObjectType;
                tagType = entry.TagType;

                if (!TryInt(m.Substring(15, 5), out rawDevice)) return Fail("#Invalid Mnemonic MSC-2");
                if (deviceIndexOverride.HasValue)
                {
                    deviceIdx = deviceIndexOverride.Value;
                }
                else if (classType == 62 && chuteMap != null)
                {
                    string key = IoListLogic.Norm(rawDevice.ToString(CultureInfo.InvariantCulture));
                    if (!chuteMap.TryGetValue(key, out deviceIdx))
                        return Fail("#Chute " + rawDevice + " not in Destination_Parameters");
                    deviceIdx += rules.EntryOffset(plantType, objectType);
                }
                else
                {
                    deviceIdx = rawDevice + rules.EntryOffset(plantType, objectType);
                }
            }
            else
            {
                return Fail("#Invalid Mnemonic MSC-2");
            }

            string alWr = m.Length == 14 ? m.Substring(7, 2) : m.Substring(9, 2);
            if (alWr != "AL" && alWr != "WR") return Fail("#Invalid Mnemonic MSC-3");
            if (!TryInt(m.Substring(m.Length - 2, 2), out int bit)) return Fail("#Invalid Mnemonic MSC-2");

            return Success(objectType, tagType, objectType, alWr == "WR", bit, deviceIdx, rawDevice, -1);
        }

        private static SwpMappedMnemonic MapIspc(string m, int firstLineId, int groupIndex,
            int? deviceIndexOverride, SwpIndexRules rules)
        {
            string objectType;
            int prefixLength, rawDevice;

            if (StartsWith(m, "CABI"))
            {
                // CABINETAL 0 001 62 — prefix 9, line = char 10, device = 11-13, bit = last 2.
                objectType = "CABINET";
                prefixLength = m.Length - 6;
                if (m.Length < 15 || !TryInt(m.Substring(10, 3), out rawDevice)) return Fail("#Invalid Mnemonic IND-1");
            }
            else if (StartsWith(m, "BELT"))
            {
                // BELTAL 1 01 62 — prefix 6, line = char 7, device = 8-9, bit = last 2.
                objectType = "BELT";
                prefixLength = m.Length - 5;
                if (m.Length < 11 || !TryInt(m.Substring(7, 2), out rawDevice)) return Fail("#Invalid Mnemonic IND-1");
            }
            else if (StartsWith(m, "ZONE"))
            {
                // ZONEAL 1 1 62 — prefix 6, line = char 7, device = char 8, bit = last 2.
                objectType = "ZONE";
                prefixLength = m.Length - 4;
                if (m.Length < 10 || !TryInt(m.Substring(7, 1), out rawDevice)) return Fail("#Invalid Mnemonic IND-1");
            }
            else
            {
                return Fail("#Invalid Mnemonic IND-1");
            }

            if (prefixLength < 2 || prefixLength > m.Length) return Fail("#Invalid Mnemonic IND-2");
            string alWr = m.Substring(prefixLength - 2, 2);
            if (alWr != "AL" && alWr != "WR") return Fail("#Invalid Mnemonic IND-2");
            if (!TryInt(m.Substring(m.Length - 2, 2), out int bit)) return Fail("#Invalid Mnemonic IND-2");

            if (prefixLength >= m.Length || !TryInt(m.Substring(prefixLength, 1), out int lineIdx))
                return Fail("#Invalid Mnemonic IND-1");

            string objectKey = lineIdx == 0 && objectType == "CABINET"
                ? SwpIndexRules.GroupCabinetKey
                : objectType;
            int deviceIdx = deviceIndexOverride
                ?? rawDevice + rules.EntryOffset(SwpPlantType.Ispc, objectKey);
            string groupOrLine = lineIdx > 0
                ? "IL" + (firstLineId + lineIdx - 1).ToString("00", CultureInfo.InvariantCulture)
                : "ILGRP" + groupIndex.ToString("00", CultureInfo.InvariantCulture);

            string sheet = groupOrLine + "_" + objectType;
            return Success(sheet, "DB_" + sheet, objectType, alWr == "WR", bit, deviceIdx, rawDevice, lineIdx);
        }

        private static SwpMappedMnemonic MapUps(string m, int firstLineId,
            int? deviceIndexOverride, SwpIndexRules rules)
        {
            string[] parts = m.Split('_');
            if (parts.Length < 4) return Fail("#Invalid Mnemonic UP-1");

            string prefix = parts[0];
            if (prefix.Length == 0 || (prefix[0] != 'A' && prefix[0] != 'W')) return Fail("#Invalid Mnemonic UP-1");

            string objectType;
            switch (prefix[prefix.Length - 1])
            {
                case 'B': objectType = "BELT"; break;
                case 'Q': objectType = "CABINET"; break;
                case 'Z': objectType = "ZONE"; break;
                default: return Fail("#Invalid Mnemonic UP-2");
            }

            if (!TryInt(parts[3], out int bit)) return Fail("#Invalid Mnemonic UP-1");
            if (!TryInt(parts[1], out int systemIdx)) return Fail("#Invalid Mnemonic UP-1");
            if (!TryInt(parts[2], out int rawDevice)) return Fail("#Invalid Mnemonic UP-1");

            int deviceIdx = deviceIndexOverride
                ?? rawDevice + rules.EntryOffset(SwpPlantType.Ups, objectType); // official: mnemonics 1-based, entries 0-based
            int swpSystem = firstLineId + systemIdx - 1;

            string sheet = "UP" + swpSystem.ToString("00", CultureInfo.InvariantCulture) + "_" + objectType;
            return Success(sheet, "DB_" + sheet, objectType, prefix[0] == 'W', bit, deviceIdx, rawDevice, systemIdx);
        }

        private static SwpMappedMnemonic Success(string sheet, string dbTag, string objectType,
            bool isWarning, int bit, int deviceIdx, int rawDevice, int referenceIndex)
        {
            string word = (isWarning ? "Warning" : "Alarm") + (int)Math.Ceiling((bit + 1) / 32.0);
            return new SwpMappedMnemonic
            {
                Coordinate = new SwpCoordinate
                {
                    SheetName = sheet,
                    TagType = dbTag + "." + word,
                    TableEntry = deviceIdx,
                    Bit = bit,
                },
                IsWarning = isWarning,
                ObjectType = objectType,
                RawDeviceNumber = rawDevice,
                ReferenceIndex = referenceIndex,
            };
        }

        private static SwpMappedMnemonic Fail(string error) => new SwpMappedMnemonic { Error = error };

        private static bool StartsWith(string s, string prefix) => s.StartsWith(prefix, StringComparison.Ordinal);

        private static bool TryInt(string s, out int value) =>
            int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool AllDigits(string s, int start, int count)
        {
            if (start + count > s.Length) return false;
            for (int i = start; i < start + count; i++)
                if (s[i] < '0' || s[i] > '9') return false;
            return true;
        }
    }
}
