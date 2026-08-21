using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// SWP-content audits from the official configuration rules, run over every indexed
    /// diagnostic sheet (not just the rows the IO list touches): duplicated alarm/warning
    /// strings within a device block (§3.3), exported rows on the forbidden bits 31/63 (§3.3),
    /// the priority column E of exported rows (§3.1), and '-' inside location names (§3.8).
    /// Findings are log lines only — nothing here queues an edit.
    /// </summary>
    public static class SwpContentChecks
    {
        private const int MaxDetailedPerCheck = 30;

        public static void Run(SwpWorkbookIndex index, List<SwpLogEntry> log)
        {
            if (index == null || log == null) return;

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
                CheckBlocks(sheet, Report);
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
            Action<string, SwpLogSeverity, string> report)
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

                if (row.Bit == 31 || row.Bit == 63)
                    report("forbidden bits", SwpLogSeverity.Warning,
                        "sheet " + sheet.SheetName + " row " + row.RowNumber + " (" + row.RowKey +
                        ") is exported on bit " + row.Bit + " — the last bit of a DWORD is " +
                        "forbidden (rules §3.3)");

                if (section != Section.State)
                {
                    CheckPriority(sheet, row, section, report);

                    string description = row.Desc1T.Trim();
                    if (description.Length > 0)
                    {
                        if (descriptionsSeen.TryGetValue(description, out SwpRowSnapshot first))
                            report("duplicated strings", SwpLogSeverity.Warning,
                                "sheet " + sheet.SheetName + " rows " + first.RowNumber + " and " +
                                row.RowNumber + " both export \"" + description + "\" — every " +
                                "string must appear only once per block (rules §3.3)");
                        else
                            descriptionsSeen[description] = row;
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

        private enum Section { Alarm, Warning, State }

        private static Section SectionOf(string tagType)
        {
            if (tagType.IndexOf(".Warning", StringComparison.Ordinal) >= 0) return Section.Warning;
            if (tagType.IndexOf(".Alarm", StringComparison.Ordinal) >= 0) return Section.Alarm;
            return Section.State;
        }

        private static bool IsExported(string exportB) =>
            exportB.Trim().Equals("x", StringComparison.OrdinalIgnoreCase);
    }
}
