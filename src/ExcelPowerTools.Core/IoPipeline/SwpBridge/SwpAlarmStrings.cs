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
        // ---- ownership ----------------------------------------------------------------------

        /// <summary>
        /// Which of the offered exported alarm lists belong to THIS SWP, by file name.
        ///
        /// A production folder commonly holds the lists for a whole plant — or for every system of
        /// one type — so they can be collected from one place; each controller then picks its own,
        /// because the indexes are global. A foreign list is not damage and must not be parsed:
        /// its mnemonics are shaped for another controller and every line of it would be reported
        /// as malformed (the field case: "Alarms_S1_IL03.txt line 414: mnemonic
        /// \"IL03.BELTWR0630\" has the wrong length", an induct list read by a sorter SWP).
        ///
        /// The authority is the SWP itself: every diagnostic sheet names the list it exports in
        /// row 5, "Output filename". An SWP whose sheets declare NOTHING — an older file, or one
        /// whose parameter was cleared — keeps every offered file: parsing a foreign list costs
        /// warnings, while silently dropping the real one would leave the coherence leg with no
        /// alarms at all and nothing to say about it.
        /// </summary>
        public static List<string> OwnedFiles(SwpWorkbookIndex index, IReadOnlyList<string> fileNames,
            List<SwpLogEntry> log)
        {
            var kept = new List<string>();
            if (fileNames == null) return kept;

            var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (index != null)
                foreach (SwpSheetIndex sheet in index.Sheets.Values)
                {
                    if (!sheet.IsDiagnosticSheet) continue;
                    string name = (sheet.OutputFileName ?? "").Trim();
                    if (name.Length > 0) declared.Add(name);
                }

            if (declared.Count == 0)
            {
                kept.AddRange(fileNames);
                if (fileNames.Count > 0)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                        "no SWP sheet names its exported alarm list (row 5, \"Output filename\") — " +
                        "every offered file is read, so a list belonging to another controller " +
                        "will report malformed mnemonics"));
                return kept;
            }

            var skipped = new List<string>();
            foreach (string name in fileNames)
            {
                if (declared.Contains((name ?? "").Trim())) kept.Add(name);
                else skipped.Add(name);
            }

            foreach (string name in skipped)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    "alarm strings " + name + " skipped — not part of the current SWP file, " +
                    "whose sheets export " + Sample(declared)));

            if (kept.Count == 0 && skipped.Count > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "none of the " + skipped.Count + " offered alarm-strings file(s) belong to this " +
                    "SWP, which exports " + Sample(declared) +
                    " — the alarm-strings check has nothing to work with"));

            return kept;
        }

        /// <summary>A few declared names, so the log line stays readable on a 20-sheet SWP.</summary>
        private static string Sample(HashSet<string> declared)
        {
            var names = new List<string>(declared);
            names.Sort(StringComparer.OrdinalIgnoreCase);
            if (names.Count <= 3) return string.Join(", ", names.ToArray());
            return string.Join(", ", names.GetRange(0, 3).ToArray()) +
                " and " + (names.Count - 3) + " more";
        }

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
