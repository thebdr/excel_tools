using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>One exported alarm string, decoded per plant type (the legacy AS_* fields).</summary>
    public sealed class SwpAlarmString
    {
        public string Mnemonic = "";
        public int CabinetNum = -1;
        public int DiagBit = -1;
        public bool IsWarning;
        public int ReferenceIndex = -1;
        public int GroupIndex = -1;
        public bool IsGroupAlarm;
        public bool IsBeltAlarm;
        public bool IsZoneAlarm;
        public string Description1 = "";
        public string Description2 = "";
        public string SourceLabel = "";   // the Alarms_*.txt file name
        public int Line;                  // 1-based line in that file
    }

    /// <summary>
    /// The alarm-strings coherence leg — the exported Alarms_*.txt files checked against the IO
    /// list. Parsing keeps the legacy field surgery verbatim (the Mid() indexes are the law);
    /// MATCHING carries the official v49 upgrades ("ddsng 9.4.2"): MSC/SAFETY matches
    /// structurally (cabinet −1, bit, warning) instead of by mnemonic text, and ISPC line
    /// cabinets dropped their +1 offset (group cabinets keep it).
    /// </summary>
    public static class SwpAlarmStrings
    {
        // ---- parsing ------------------------------------------------------------------------

        public static List<SwpAlarmString> Parse(IEnumerable<string> lines, SwpPlantType plantType,
            string sourceLabel, List<SwpLogEntry> log)
        {
            var alarms = new List<SwpAlarmString>();
            if (lines == null) return alarms;

            var seenTags = new HashSet<string>(StringComparer.Ordinal);
            int duplicateTags = 0;
            int lineNumber = 0;
            foreach (string line in lines)
            {
                lineNumber++;
                string[] fields = (line ?? "").Split(',');
                if (fields.Length < 11)
                {
                    // The official v49 upgrade: name the offending line, not just "wrong format".
                    if (!Field(fields, 0).Contains("#") && (line ?? "").Trim().Length > 0)
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "alarm strings " + sourceLabel + " line " + lineNumber +
                            " has a wrong format: [" + line + "]"));
                    continue;
                }

                string tag = Field(fields, 1);
                string description1 = Field(fields, 9);
                if (!IsNumeric(Field(fields, 2))) continue;
                if (tag.Length > 0 && !seenTags.Add(tag)) duplicateTags++;
                if (description1.Contains(" -  AL") || description1.Contains(" -  WR") ||
                    description1.Contains(" - Spare")) continue;
                if (!tag.Contains("CABINET") && !tag.Contains("BELT") && !tag.Contains("ZONE")) continue;

                SwpAlarmString alarm = Decode(tag, plantType, sourceLabel, lineNumber, log);
                if (alarm == null) continue;
                alarm.IsWarning = alarm.Mnemonic.Contains("WR");
                alarm.Description1 = description1;
                alarm.Description2 = Field(fields, 10);
                alarm.SourceLabel = sourceLabel ?? "";
                alarm.Line = lineNumber;
                alarms.Add(alarm);
            }

            // ExcelLoader APPENDS to existing Alarms_*.txt files instead of overwriting (rules
            // §11 tells operators to delete the old ones first) — duplicated tags are the smell.
            if (duplicateTags > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "alarm strings " + sourceLabel + ": " + duplicateTags + " duplicated tag(s) — " +
                    "the export looks appended-to; delete the old Alarms_*.txt and regenerate"));
            return alarms;
        }

        private static SwpAlarmString Decode(string tag, SwpPlantType plantType,
            string sourceLabel, int lineNumber, List<SwpLogEntry> log)
        {
            switch (plantType)
            {
                case SwpPlantType.Msc:
                case SwpPlantType.Safety:
                {
                    // "S1.CABINETAL00000" → mnemonic Mid(4,15); cab Mid(10,3); bit Mid(13,2).
                    // SAFETY exports carry the same cabinet-tag shape (a cabinet-only SWP).
                    string m = Mid(tag, 4, 15);
                    if (m.Length != 14)
                    {
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "alarm strings " + sourceLabel + " line " + lineNumber +
                            ": mnemonic \"" + m + "\" has the wrong length"));
                        return null;
                    }
                    if (!TryInt(Mid(m, 10, 3), out int cab) || !TryInt(Mid(m, 13, 2), out int bit))
                        return null;
                    return new SwpAlarmString { Mnemonic = m, CabinetNum = cab, DiagBit = bit };
                }

                case SwpPlantType.Ispc:
                {
                    if (tag.Contains("GRP"))
                    {
                        // "S1.ILGRP01.CABINETAL…" → Mid(4,23); group Mid(6,2); cab Mid(18,3); bit Mid(21,2).
                        string m = Mid(tag, 4, 23);
                        if (!TryInt(Mid(m, 6, 2), out int group) ||
                            !TryInt(Mid(m, 18, 3), out int cab) ||
                            !TryInt(Mid(m, 21, 2), out int bit)) return null;
                        return new SwpAlarmString
                        {
                            Mnemonic = m,
                            GroupIndex = group,
                            CabinetNum = cab,
                            DiagBit = bit,
                            IsGroupAlarm = true,
                        };
                    }
                    if (tag.Contains("CABINET"))
                    {
                        // "S1.IL01.CABINETAL…" → Mid(4,20); ref Mid(3,2); cab Mid(15,3); bit Mid(18,2).
                        string m = Mid(tag, 4, 20);
                        if (!TryInt(Mid(m, 3, 2), out int reference) ||
                            !TryInt(Mid(m, 15, 3), out int cab) ||
                            !TryInt(Mid(m, 18, 2), out int bit)) return null;
                        return new SwpAlarmString
                        {
                            Mnemonic = m,
                            ReferenceIndex = reference,
                            CabinetNum = cab,
                            DiagBit = bit,
                        };
                    }
                    return null;
                }

                case SwpPlantType.Ups:
                {
                    string m = Mid(tag, 4, 20);
                    if (tag.Contains("CABINET"))
                    {
                        if (!TryInt(Mid(m, 3, 2), out int reference) ||
                            !TryInt(Mid(m, 15, 3), out int cab) ||
                            !TryInt(Mid(m, 18, 2), out int bit))
                        {
                            log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                                "alarm strings " + sourceLabel + " line " + lineNumber +
                                ": unreadable cabinet tag \"" + tag + "\""));
                            return null;
                        }
                        return new SwpAlarmString
                        {
                            Mnemonic = m,
                            ReferenceIndex = reference,
                            CabinetNum = cab,
                            DiagBit = bit,
                        };
                    }
                    if (tag.Contains("BELT"))
                    {
                        // "S1.UP05.BELTAL00100" → cab Mid(12,3); bit Mid(15,2).
                        if (!TryInt(Mid(m, 3, 2), out int reference) ||
                            !TryInt(Mid(m, 12, 3), out int cab) ||
                            !TryInt(Mid(m, 15, 2), out int bit)) return null;
                        return new SwpAlarmString
                        {
                            Mnemonic = m,
                            ReferenceIndex = reference,
                            CabinetNum = cab,
                            DiagBit = bit,
                            IsBeltAlarm = true,
                        };
                    }
                    if (tag.Contains("ZONE"))
                    {
                        // "S1.UP05.ZONEAL0100" → cab Mid(12,2); bit Mid(14,2).
                        if (!TryInt(Mid(m, 3, 2), out int reference) ||
                            !TryInt(Mid(m, 12, 2), out int cab) ||
                            !TryInt(Mid(m, 14, 2), out int bit)) return null;
                        return new SwpAlarmString
                        {
                            Mnemonic = m,
                            ReferenceIndex = reference,
                            CabinetNum = cab,
                            DiagBit = bit,
                            IsZoneAlarm = true,
                        };
                    }
                    return null;
                }

                default:
                    return null;
            }
        }

        // ---- matching (the v49 semantics) ---------------------------------------------------

        public static SwpAlarmString Match(SwpMappedMnemonic mapped, SwpPlantType plantType,
            int firstLineId, int groupIndex, IReadOnlyList<SwpAlarmString> alarms,
            SwpIndexRules rules = null)
        {
            if (mapped == null || !mapped.Ok || alarms == null) return null;
            rules = rules ?? DefaultRules;

            // The export's expected index for this mnemonic: raw digits + the object's OPC-tag
            // shift from the rules table (the official v93 values reproduce the v49 semantics —
            // MSC cabinets +1, ISPC groups and UPS cabinets −1, everything else 1:1).
            int expected = mapped.RawDeviceNumber +
                rules.AlarmOffset(plantType, SwpIndexRules.ObjectKeyFor(plantType, mapped));
            int bit = mapped.Coordinate.Bit;

            foreach (SwpAlarmString alarm in alarms)
            {
                if (bit != alarm.DiagBit || mapped.IsWarning != alarm.IsWarning) continue;

                switch (plantType)
                {
                    case SwpPlantType.Msc:
                    case SwpPlantType.Safety:
                        // The parser only yields cabinet tags here — never let a CLASSDIAG
                        // object's numbers collide into them.
                        if (mapped.ObjectType == "CABINET" && expected == alarm.CabinetNum)
                            return alarm;
                        break;

                    case SwpPlantType.Ispc:
                        if (mapped.ObjectType != "CABINET") break;
                        if (mapped.ReferenceIndex > 0)
                        {
                            if (!alarm.IsGroupAlarm &&
                                mapped.ReferenceIndex + firstLineId - 1 == alarm.ReferenceIndex &&
                                expected == alarm.CabinetNum)
                                return alarm;
                        }
                        else if (alarm.IsGroupAlarm &&
                                 groupIndex == alarm.GroupIndex &&
                                 expected == alarm.CabinetNum)
                        {
                            return alarm;
                        }
                        break;

                    case SwpPlantType.Ups:
                        if (mapped.ReferenceIndex > 0 &&
                            mapped.ReferenceIndex + firstLineId - 1 == alarm.ReferenceIndex &&
                            expected == alarm.CabinetNum &&
                            (mapped.ObjectType == "BELT") == alarm.IsBeltAlarm &&
                            (mapped.ObjectType == "ZONE") == alarm.IsZoneAlarm)
                            return alarm;
                        break;
                }
            }
            return null;
        }

        private static readonly SwpIndexRules DefaultRules = SwpIndexRules.Official();

        // ---- helpers ------------------------------------------------------------------------

        /// <summary>VBA Mid(): 1-based, tolerant of running past the end.</summary>
        internal static string Mid(string s, int start, int length)
        {
            s = s ?? "";
            if (start < 1 || start > s.Length) return "";
            return s.Substring(start - 1, Math.Min(length, s.Length - start + 1));
        }

        private static string Field(string[] fields, int index) =>
            fields != null && index < fields.Length ? (fields[index] ?? "").Trim() : "";

        private static bool TryInt(string s, out int value) =>
            int.TryParse((s ?? "").Trim(), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out value);

        private static bool IsNumeric(string s) =>
            double.TryParse((s ?? "").Trim(), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out _);
    }
}
