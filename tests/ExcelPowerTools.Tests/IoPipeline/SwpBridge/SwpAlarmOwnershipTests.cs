using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// Which exported alarm lists belong to an SWP. A production folder often holds the lists for a
    /// whole plant so they can be collected from one place; each controller picks its own, because
    /// the indexes are global. The SWP names what it exports in row 5 of every diagnostic sheet
    /// ("Output filename"), and that set is the authority.
    ///
    /// The field case: a sorter SWP was handed the induct list and reported every line of it —
    /// "Alarms_S1_IL03.txt line 414: mnemonic \"IL03.BELTWR0630\" has the wrong length".
    /// </summary>
    public class SwpAlarmOwnershipTests
    {
        private static SwpWorkbookIndex Index(params (string Sheet, string Output)[] sheets)
        {
            var index = new SwpWorkbookIndex
            {
                Sheets = new Dictionary<string, SwpSheetIndex>(StringComparer.OrdinalIgnoreCase),
            };
            foreach ((string sheet, string output) in sheets)
                index.Sheets[sheet] = new SwpSheetIndex
                {
                    SheetName = sheet,
                    IsDiagnosticSheet = true,
                    OutputFileName = output,
                };
            return index;
        }

        [Fact]
        public void A_list_from_another_controller_is_skipped_with_an_info_line()
        {
            SwpWorkbookIndex swp = Index(("CABINET", "Alarms_S1_CABINET.txt"),
                                         ("CHUTE", "Alarms_S1_CHUTE.txt"));
            var log = new List<SwpLogEntry>();

            List<string> kept = SwpAlarmStrings.OwnedFiles(swp,
                new[] { "Alarms_S1_CABINET.txt", "Alarms_S1_IL03.txt" }, log);

            Assert.Equal(new[] { "Alarms_S1_CABINET.txt" }, kept);
            SwpLogEntry skip = log.Single(l => l.Message.Contains("Alarms_S1_IL03.txt"));
            Assert.Equal(SwpLogSeverity.Info, skip.Severity);
            Assert.Contains("skipped", skip.Message);
            Assert.Contains("not part of the current SWP file", skip.Message);
        }

        [Fact]
        public void An_owned_list_is_never_announced_as_skipped()
        {
            SwpWorkbookIndex swp = Index(("CABINET", "Alarms_S1_CABINET.txt"));
            var log = new List<SwpLogEntry>();

            List<string> kept = SwpAlarmStrings.OwnedFiles(swp, new[] { "Alarms_S1_CABINET.txt" }, log);

            Assert.Single(kept);
            Assert.DoesNotContain(log, l => l.Message.Contains("skipped"));
        }

        [Fact]
        public void The_name_match_ignores_case_and_surrounding_space()
        {
            SwpWorkbookIndex swp = Index(("CABINET", "  Alarms_S1_CABINET.txt  "));
            var log = new List<SwpLogEntry>();

            Assert.Single(SwpAlarmStrings.OwnedFiles(swp, new[] { "alarms_s1_cabinet.TXT" }, log));
        }

        [Fact]
        public void An_swp_that_declares_nothing_keeps_every_file()
        {
            // Older SWPs carry no row-5 parameter. Dropping everything would leave the coherence
            // leg silent; parsing a foreign list only costs warnings, so the offered set stands.
            SwpWorkbookIndex swp = Index(("CABINET", ""));
            var log = new List<SwpLogEntry>();

            List<string> kept = SwpAlarmStrings.OwnedFiles(swp,
                new[] { "Alarms_S1_CABINET.txt", "Alarms_S1_IL03.txt" }, log);

            Assert.Equal(2, kept.Count);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info &&
                l.Message.Contains("no SWP sheet names its exported alarm list"));
        }

        [Fact]
        public void A_non_diagnostic_sheet_never_claims_a_file()
        {
            SwpWorkbookIndex swp = Index(("CABINET", "Alarms_S1_CABINET.txt"));
            swp.Sheets["Cover"] = new SwpSheetIndex
            {
                SheetName = "Cover",
                IsDiagnosticSheet = false,
                OutputFileName = "Alarms_S1_IL03.txt",
            };
            var log = new List<SwpLogEntry>();

            Assert.Equal(new[] { "Alarms_S1_CABINET.txt" },
                SwpAlarmStrings.OwnedFiles(swp, new[] { "Alarms_S1_CABINET.txt", "Alarms_S1_IL03.txt" }, log));
        }

        [Fact]
        public void Losing_every_file_is_a_warning_not_a_silent_skip()
        {
            // Pointing at the wrong plant's folder must not read as "the alarms all check out".
            SwpWorkbookIndex swp = Index(("CABINET", "Alarms_S1_CABINET.txt"));
            var log = new List<SwpLogEntry>();

            Assert.Empty(SwpAlarmStrings.OwnedFiles(swp,
                new[] { "Alarms_S1_IL03.txt", "Alarms_S1_IL04.txt" }, log));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("none of the 2 offered"));
        }

        [Fact]
        public void The_skip_line_names_what_the_swp_does_export()
        {
            SwpWorkbookIndex swp = Index(("CABINET", "Alarms_S1_CABINET.txt"),
                                         ("CHUTE", "Alarms_S1_CHUTE.txt"),
                                         ("SCANNER", "Alarms_S1_SCANNER.txt"),
                                         ("ENCODER", "Alarms_S1_ENCODER.txt"),
                                         ("COIL", "Alarms_S1_COIL.txt"));
            var log = new List<SwpLogEntry>();

            SwpAlarmStrings.OwnedFiles(swp, new[] { "Alarms_S1_IL03.txt" }, log);

            // Sampled, so a 20-sheet SWP does not print a paragraph per skipped file.
            SwpLogEntry skip = log.Single(l => l.Message.Contains("Alarms_S1_IL03.txt"));
            Assert.Contains("and 2 more", skip.Message);
        }

        [Fact]
        public void A_null_index_or_no_files_is_harmless()
        {
            Assert.Empty(SwpAlarmStrings.OwnedFiles(null, new string[0], new List<SwpLogEntry>()));
            Assert.Empty(SwpAlarmStrings.OwnedFiles(Index(("CABINET", "x.txt")), null,
                new List<SwpLogEntry>()));
        }
    }
}
