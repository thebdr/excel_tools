using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// Drives the base audit over the REAL sorter pair. Silent on machines without the files, like
    /// the other real-file tests. The field case that motivated the block-header anchor:
    /// audit 'CHUTE' row 207 col R reported "F181" against the reference's "F15" — the same
    /// structural formula, each pointing at its own block header.
    /// </summary>
    public class SwpReferenceScanTests
    {
        private const string ReferenceSwp =
            @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\0 Commesse\FVZ_Intermarche\3 Software" +
            @"\3_Diagnostica\SWP\sorter\sorter_base\35_21-Mag-2026\SWP_01_Sorter.xlsx";

        private const string TargetSwp =
            @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\0 Commesse\FVZ_Intermarche\3 Software" +
            @"\3_Diagnostica\SWP\sorter\06_20-Aug-2026\SWP_01_Sorter_v2.xlsx";

        private static readonly string[] MscSheets =
        {
            "CABINET", "CHUTE", "SCANNER", "CONTROL_ZONE", "CHECK_POSITION", "CELL_TEST",
            "AIR_BLOWER", "ITEM_LOCATOR", "ENCODER", "BUS_BREAKER", "DRIVE_UNIT", "SORTERSTATE",
            "AUXZERO", "BLADE_CHECK", "ENERGY_BOX", "ENERGY_BOX_TEST", "TRAIN", "COORDINATOR",
            "MODULATOR", "COIL",
        };

        [Fact]
        public void The_real_audit_no_longer_reports_own_block_header_references()
        {
            if (!File.Exists(ReferenceSwp) || !File.Exists(TargetSwp)) return;

            var readLog = new List<SwpLogEntry>();
            SwpWorkbookIndex target = SwpSheetReader.Read(
                SwpSheetReaderTests.ReadShared(TargetSwp), MscSheets, readLog);
            SwpWorkbookIndex reference = SwpSheetReader.Read(
                SwpSheetReaderTests.ReadShared(ReferenceSwp), MscSheets, readLog);

            var log = new List<SwpLogEntry>();
            SwpFormulaCheck.AuditAgainstBase(target, reference, log);

            List<SwpLogEntry> refDiffs = log
                .Where(l => l.Message.IndexOf("CONFIG_PLANT reference differs",
                    StringComparison.Ordinal) >= 0)
                .ToList();

            // Both files were scanned exhaustively: every same-sheet reference in them resolves to
            // the row itself or to the row's own block header, so a structural difference is the
            // only thing left that can fire here.
            Assert.DoesNotContain(refDiffs,
                l => l.Message.Contains("'CHUTE' row 207") && l.Message.Contains("col R"));
            Assert.Empty(refDiffs);
        }

        [Fact]
        public void The_real_swp_names_its_own_alarm_lists_in_row_5()
        {
            if (!File.Exists(TargetSwp)) return;

            var readLog = new List<SwpLogEntry>();
            SwpWorkbookIndex swp = SwpSheetReader.Read(
                SwpSheetReaderTests.ReadShared(TargetSwp), MscSheets, readLog);

            Assert.Equal("Alarms_S1_CABINET.txt", swp.Sheets["CABINET"].OutputFileName);

            List<string> declared = swp.Sheets.Values
                .Where(s => s.IsDiagnosticSheet && s.OutputFileName.Length > 0)
                .Select(s => s.OutputFileName)
                .ToList();
            Assert.NotEmpty(declared);

            // The induct list from the field case belongs to another controller entirely.
            Assert.DoesNotContain("Alarms_S1_IL03.txt", declared);
            Assert.Empty(SwpAlarmStrings.OwnedFiles(swp,
                new[] { "Alarms_S1_IL03.txt" }, new List<SwpLogEntry>()));
        }
    }
}
