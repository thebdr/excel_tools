using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// SWP-content audits from the official configuration rules, run over every indexed
    /// diagnostic sheet (not just the rows the IO list touches): duplicated strings within a
    /// device block AND category (§3.3), rows where a MNEMONIC lands on the forbidden bits
    /// 31/63 (§3.3), the priority column E of exported rows (§3.1), and '-' inside location
    /// names (§3.8). Findings are log lines only — nothing here queues an edit.
    ///
    /// Two of these are narrower than the rule text reads, because the wider reading produced
    /// findings nobody could act on. The same string as an alarm AND as a warning is the normal
    /// shape of a declassed signal, not a duplicate; and an export tick on bit 31/63 is ignored
    /// by the plant, so only a mnemonic actually landing there is worth reporting.
    /// </summary>
    public static class SwpContentChecks
    {
        private const int MaxDetailedPerCheck = 30;

        /// <summary>
        /// <paramref name="targetedRows"/> — "sheet|row" for every SWP row a mnemonic maps to.
        /// Only the forbidden-bit check consults it, and null means it does not run at all: a
        /// finding that says "a mnemonic lands on bit 63" cannot be made without knowing where the
        /// mnemonics land, and guessing from the export tick alone is what made it noise.
        /// </summary>
        public static void Run(SwpWorkbookIndex index, List<SwpLogEntry> log,
            HashSet<string> targetedRows = null)
        {
            if (index == null || log == null) return;
            HashSet<string> targeted = targetedRows;

            var counters = new Dictionary<string, int>();
            void Report(string check, SwpLogSeverity severity, string message)
            {
                counters.TryGetValue(check, out int n);
                counters[check] = n + 1;
                if (n < MaxDetailedPerCheck) log.Add(SwpLogEntry.For(severity, message));
            }

            foreach (SwpSheetIndex sheet in index.Sheets.Values)
            {
                if (!sheet.IsDiagnosticSheet) continue;

                CheckHeaderLocations(sheet, Report);
                CheckBlocks(sheet, Report, targeted);
            }

            foreach (KeyValuePair<string, int> pair in counters)
                if (pair.Value > MaxDetailedPerCheck)
                    log.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        pair.Key + ": " + pair.Value + " finding(s) in total — only the first " +
                        MaxDetailedPerCheck + " are listed above"));
        }

        private static void CheckHeaderLocations(SwpSheetIndex sheet,
            Action<string, SwpLogSeverity, string> report)
        {
            foreach (SwpBlockHeader header in sheet.BlockHeaders)
            {
                if (header.LocationI.IndexOf('-') < 0) continue;
                report("locations with '-'", SwpLogSeverity.Warning,
                    "sheet " + sheet.SheetName + " row " + header.RowNumber + ": location \"" +
                    header.LocationI + "\" contains '-' — forbidden in location names (rules §3.8)");
            }
        }

        private static void CheckBlocks(SwpSheetIndex sheet,
            Action<string, SwpLogSeverity, string> report, HashSet<string> targeted)
        {
            // Blocks keyed by their header row; rows arrive in file order.
            var descriptionsSeen = new Dictionary<string, SwpRowSnapshot>(StringComparer.OrdinalIgnoreCase);
            int currentBlock = -1;

            foreach (SwpRowSnapshot row in sheet.BitRows)
            {
                if (row.BlockHeaderRow != currentBlock)
                {
                    currentBlock = row.BlockHeaderRow;
                    descriptionsSeen.Clear();
                }

                if (!IsExported(row.ExportB)) continue;
                Section section = SectionOf(row.TagType);

                // The last bit of a DWORD is forbidden, but the SWP marking one exported is
                // harmless on its own — the line is ignored. What matters is a MNEMONIC landing
                // there, because that is a signal the plant will never see. So the finding needs
                // the row to be one the IO list actually targets; without that it was reporting
                // stray export ticks nobody can act on.
                if ((row.Bit == 31 || row.Bit == 63) && IsTargeted(targeted, sheet.SheetName, row.RowNumber))
                    report("forbidden bits", SwpLogSeverity.Warning,
                        "sheet " + sheet.SheetName + " row " + row.RowNumber + " (" + row.RowKey +
                        ") takes a mnemonic on bit " + row.Bit + " — the last bit of a DWORD is " +
                        "forbidden (rules §3.3)");

                if (section != Section.State)
                {
                    CheckPriority(sheet, row, section, report);

                    // What must be unique is the IDENTITY, not the sentence. A signal is its
                    // functional unit + location + device + description: sixteen safety doors on
                    // one block all read "SAFETY DOOR OPEN" and are told apart by their device
                    // (-B1, -B2, …), so keying on the text alone reported fifteen duplicates that
                    // were nothing of the kind. And the category counts too — the same identity as
                    // an alarm and as a warning is a declassed signal, not a repeat.
                    string description = row.Desc1T.Trim();
                    if (description.Length > 0)
                    {
                        string identity = IdentityOf(row);
                        string key = section + "|" + identity + "|" + description;
                        if (descriptionsSeen.TryGetValue(key, out SwpRowSnapshot first))
                            report("duplicated strings", SwpLogSeverity.Warning,
                                "sheet " + sheet.SheetName + " rows " + first.RowNumber + " and " +
                                row.RowNumber + " are the same " + section.ToString().ToLowerInvariant() +
                                ": " + (identity.Length > 0 ? identity + " " : "") + "\"" + description +
                                "\" — functional unit + location + device + description must be " +
                                "unique per block (rules §3.3)");
                        else
                            descriptionsSeen[key] = row;
                    }
                }
            }
        }

        private static void CheckPriority(SwpSheetIndex sheet, SwpRowSnapshot row, Section section,
            Action<string, SwpLogSeverity, string> report)
        {
            string where = "sheet " + sheet.SheetName + " row " + row.RowNumber + " (" + row.RowKey + ")";
            string raw = row.PriorityE.Trim();

            if (raw.Length == 0 ||
                !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double priority) ||
                priority == 0)
            {
                report("priorities", SwpLogSeverity.Warning,
                    where + ": exported but priority E is \"" + row.PriorityE +
                    "\" — must be a number other than 0 (rules §3.1)");
                return;
            }
            if (row.PriorityIsText)
            {
                report("priorities", SwpLogSeverity.Warning,
                    where + ": priority E holds \"" + raw + "\" as TEXT — set the cell category " +
                    "to General (rules §3.1)");
                return;
            }

            int expected = section == Section.Warning ? 3 : 2;
            if ((int)priority != expected)
                report("priorities", SwpLogSeverity.Info,
                    where + ": priority " + raw + " where the convention is " + expected +
                    " (alarm = 2, warning = 3 — rules §3.1)");
        }

        /// <summary>
        /// The row's electrical identity as the maintainer reads it: its OWN functional unit and
        /// location where it carries them, the block header's otherwise (a healthy bit row's H/I
        /// are formulas pointing at the header, so the cached value IS the block's), plus its
        /// device. Empty for a row that carries none — those fall back to comparing descriptions,
        /// which is all there is to compare.
        /// </summary>
        private static string IdentityOf(SwpRowSnapshot row)
        {
            string funit = Pick(row.HasCellH ? row.CachedH : "", row.BlockFUnit);
            string location = Pick(row.HasCellI ? row.CachedI : "", row.BlockLocation);
            return (funit + location + (row.DeviceJ ?? "").Trim()).Trim();
        }

        private static string Pick(string own, string fallback)
        {
            string value = (own ?? "").Trim();
            return value.Length > 0 ? value : (fallback ?? "").Trim();
        }

        private enum Section { Alarm, Warning, State }

        private static Section SectionOf(string tagType)
        {
            if (tagType.IndexOf(".Warning", StringComparison.Ordinal) >= 0) return Section.Warning;
            if (tagType.IndexOf(".Alarm", StringComparison.Ordinal) >= 0) return Section.Alarm;
            return Section.State;
        }

        /// <summary>Was a mnemonic mapped onto this SWP row? Null set = unknown, so no.</summary>
        private static bool IsTargeted(HashSet<string> targeted, string sheet, int row) =>
            targeted != null && targeted.Contains(sheet + "|" + row.ToString(CultureInfo.InvariantCulture));

        /// <summary>The key Run expects for one SWP row.</summary>
        public static string RowKeyOf(string sheet, int row) =>
            sheet + "|" + row.ToString(CultureInfo.InvariantCulture);

        private static bool IsExported(string exportB) =>
            exportB.Trim().Equals("x", StringComparison.OrdinalIgnoreCase);
    }
}
