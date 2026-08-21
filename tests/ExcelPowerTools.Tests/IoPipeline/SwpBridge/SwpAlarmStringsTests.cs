using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpAlarmStringsTests
    {
        private static string Line(string tag, string desc1 = "SOME ALARM +MC1.CC1-F1", string desc2 = "")
            => "x," + tag + ",123,a,b,c,d,e,f," + desc1 + "," + desc2;

        private static List<SwpAlarmString> Parse(SwpPlantType plant, params string[] lines) =>
            SwpAlarmStrings.Parse(lines, plant, "test.txt", new List<SwpLogEntry>());

        // ---- parsing ------------------------------------------------------------------------

        [Fact]
        public void Msc_tags_decode_cabinet_and_bit()
        {
            SwpAlarmString alarm = Assert.Single(Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215")));
            Assert.Equal("CABINETAL00215", alarm.Mnemonic);
            Assert.Equal(2, alarm.CabinetNum);
            Assert.Equal(15, alarm.DiagBit);
            Assert.False(alarm.IsWarning);
            Assert.Equal("SOME ALARM +MC1.CC1-F1", alarm.Description1);
        }

        [Fact]
        public void Msc_warning_tags_carry_the_flag()
        {
            Assert.True(Assert.Single(Parse(SwpPlantType.Msc, Line("S1.CABINETWR00215"))).IsWarning);
        }

        [Fact]
        public void Ispc_line_and_group_tags_decode()
        {
            var alarms = Parse(SwpPlantType.Ispc,
                Line("S1.IL03.CABINETAL01562"),
                Line("S1.ILGRP02.CABINETAL01662"));

            SwpAlarmString line = alarms.First(a => !a.IsGroupAlarm);
            Assert.Equal(3, line.ReferenceIndex);
            Assert.Equal(15, line.CabinetNum);
            Assert.Equal(62, line.DiagBit);

            SwpAlarmString group = alarms.First(a => a.IsGroupAlarm);
            Assert.Equal(2, group.GroupIndex);
            Assert.Equal(16, group.CabinetNum);
            Assert.Equal(62, group.DiagBit);
        }

        [Fact]
        public void Ups_cabinet_belt_and_zone_tags_decode()
        {
            var alarms = Parse(SwpPlantType.Ups,
                Line("S1.UP03.CABINETAL00107"),
                Line("S1.UP01.BELTAL00550"),
                Line("S1.UP01.ZONEAL0250"));

            SwpAlarmString cabinet = alarms.First(a => !a.IsBeltAlarm && !a.IsZoneAlarm);
            Assert.Equal(3, cabinet.ReferenceIndex);
            Assert.Equal(1, cabinet.CabinetNum);
            Assert.Equal(7, cabinet.DiagBit);

            SwpAlarmString belt = alarms.First(a => a.IsBeltAlarm);
            Assert.Equal(1, belt.ReferenceIndex);
            Assert.Equal(5, belt.CabinetNum);
            Assert.Equal(50, belt.DiagBit);

            SwpAlarmString zone = alarms.First(a => a.IsZoneAlarm);
            Assert.Equal(2, zone.CabinetNum);
            Assert.Equal(50, zone.DiagBit);
        }

        [Theory]
        [InlineData("SOMETHING -  AL")]
        [InlineData("SOMETHING -  WR")]
        [InlineData("SOMETHING - Spare")]
        public void Untitled_spares_are_skipped(string desc1)
        {
            Assert.Empty(Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215", desc1)));
        }

        [Fact]
        public void Short_rows_log_the_offending_content()
        {
            var log = new List<SwpLogEntry>();
            SwpAlarmStrings.Parse(new[] { "too,short,row" }, SwpPlantType.Msc, "f.txt", log);
            // The v49 upgrade: file, line and content are all named.
            Assert.Contains(log, l => l.Message.Contains("line 1") && l.Message.Contains("[too,short,row]"));
        }

        [Fact]
        public void Comment_rows_and_blank_lines_stay_silent()
        {
            var log = new List<SwpLogEntry>();
            SwpAlarmStrings.Parse(new[] { "#comment", "" }, SwpPlantType.Msc, "f.txt", log);
            Assert.Empty(log);
        }

        // ---- matching (the v49 semantics) ---------------------------------------------------

        private static SwpMappedMnemonic Map(string mnemonic, SwpPlantType plant, int first = 1, int group = 1) =>
            SwpAddressMap.Map(mnemonic, plant, first, group);

        [Fact]
        public void Msc_matches_structurally_with_the_minus_one_offset()
        {
            var alarms = Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215"));
            // IO cabinet 1 ↔ exported cabinet 2 (v49: AS − 1), bit 15, both alarms.
            Assert.NotNull(SwpAlarmStrings.Match(Map("CABINETAL00115", SwpPlantType.Msc),
                SwpPlantType.Msc, 1, 1, alarms));
            Assert.Null(SwpAlarmStrings.Match(Map("CABINETAL00215", SwpPlantType.Msc),
                SwpPlantType.Msc, 1, 1, alarms));
            Assert.Null(SwpAlarmStrings.Match(Map("CABINETWR00115", SwpPlantType.Msc),
                SwpPlantType.Msc, 1, 1, alarms));
        }

        [Fact]
        public void Ispc_line_cabinets_align_one_to_one()
        {
            var alarms = Parse(SwpPlantType.Ispc, Line("S1.IL03.CABINETAL01562"));
            // IO line 1 with first-line 3 → IL03; cabinet 15 ↔ 15 (v49 dropped the +1).
            Assert.NotNull(SwpAlarmStrings.Match(Map("CABINETAL101562", SwpPlantType.Ispc, first: 3),
                SwpPlantType.Ispc, 3, 1, alarms));
            Assert.Null(SwpAlarmStrings.Match(Map("CABINETAL101662", SwpPlantType.Ispc, first: 3),
                SwpPlantType.Ispc, 3, 1, alarms));
        }

        [Fact]
        public void Ispc_group_cabinets_keep_the_plus_one()
        {
            var alarms = Parse(SwpPlantType.Ispc, Line("S1.ILGRP02.CABINETAL01662"));
            // Line 0 = the group; IO cabinet 17 ↔ exported 16 (+1 kept for groups).
            Assert.NotNull(SwpAlarmStrings.Match(Map("CABINETAL001762", SwpPlantType.Ispc, group: 2),
                SwpPlantType.Ispc, 1, 2, alarms));
            Assert.Null(SwpAlarmStrings.Match(Map("CABINETAL001762", SwpPlantType.Ispc, group: 9),
                SwpPlantType.Ispc, 1, 9, alarms));
        }

        [Fact]
        public void Ups_cabinets_offset_but_belts_and_zones_align()
        {
            var alarms = Parse(SwpPlantType.Ups,
                Line("S1.UP03.CABINETAL00107"),
                Line("S1.UP01.BELTAL00550"));

            // Cabinet: IO index 2 ↔ exported 1 (+1); system 1 with first-line 3 → UP03.
            Assert.NotNull(SwpAlarmStrings.Match(Map("ALLQ_1_02_07", SwpPlantType.Ups, first: 3),
                SwpPlantType.Ups, 3, 1, alarms));
            // Belt: IO belt 5 ↔ exported 5 (no offset), and the belt flag must agree.
            Assert.NotNull(SwpAlarmStrings.Match(Map("ALLB_1_005_50", SwpPlantType.Ups),
                SwpPlantType.Ups, 1, 1, alarms));
            Assert.Null(SwpAlarmStrings.Match(Map("ALLQ_1_05_50", SwpPlantType.Ups),
                SwpPlantType.Ups, 1, 1, alarms));
        }

        // ---- the coherence verdicts ---------------------------------------------------------

        private static SwpSpineRow Row(string mnemonic, string location, string device)
        {
            return new SwpSpineRow
            {
                Entry = new DiagnosticEntry
                {
                    Mnemonic = mnemonic,
                    Location = location,
                    Device = device,
                    CardSheet = "R1.S1.C1",
                },
                Mapped = Map(mnemonic, SwpPlantType.Msc),
            };
        }

        private static SwpWorkspaceModel Model(SwpSpineRow row)
        {
            var model = new SwpWorkspaceModel();
            model.Rows.Add(row);
            return model;
        }

        [Fact]
        public void A_matching_alarm_with_the_texts_reads_ok()
        {
            var alarms = Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215", "ALARM +MC1.CC1-F1 TEXT"));
            SwpSpineRow row = Row("CABINETAL00115", "+MC1.CC1", "-F1");

            SwpCoherence.Annotate(Model(row), null, null, SwpPlantType.Msc, new List<SwpLogEntry>(), alarms);
            Assert.Equal(SwpAlarmVerdict.Ok, row.AlarmVerdict);
        }

        [Fact]
        public void A_missing_export_reads_not_found()
        {
            var alarms = Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215"));
            SwpSpineRow row = Row("CABINETAL00999", "+X", "-Y");
            var log = new List<SwpLogEntry>();

            SwpCoherence.Annotate(Model(row), null, null, SwpPlantType.Msc, log, alarms);
            Assert.Equal(SwpAlarmVerdict.NotFound, row.AlarmVerdict);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("not found"));
        }

        [Fact]
        public void A_text_without_the_device_reads_mismatch()
        {
            var alarms = Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215", "ALARM +MC1.CC1 ONLY"));
            SwpSpineRow row = Row("CABINETAL00115", "+MC1.CC1", "-F1");
            var log = new List<SwpLogEntry>();

            SwpCoherence.Annotate(Model(row), null, null, SwpPlantType.Msc, log, alarms);
            Assert.Equal(SwpAlarmVerdict.DescriptionMismatch, row.AlarmVerdict);
            Assert.Contains(log, l => l.Message.Contains("device \"-F1\" missing"));
        }

        [Fact]
        public void Description_two_is_checked_only_when_exported()
        {
            var withDesc2 = Parse(SwpPlantType.Msc,
                Line("S1.CABINETAL00215", "OK +MC1.CC1-F1", "SEM O DISPOSITIVO"));
            SwpSpineRow row = Row("CABINETAL00115", "+MC1.CC1", "-F1");

            SwpCoherence.Annotate(Model(row), null, null, SwpPlantType.Msc, new List<SwpLogEntry>(), withDesc2);
            Assert.Equal(SwpAlarmVerdict.DescriptionMismatch, row.AlarmVerdict);

            var withoutDesc2 = Parse(SwpPlantType.Msc, Line("S1.CABINETAL00215", "OK +MC1.CC1-F1"));
            SwpSpineRow clean = Row("CABINETAL00115", "+MC1.CC1", "-F1");
            SwpCoherence.Annotate(Model(clean), null, null, SwpPlantType.Msc, new List<SwpLogEntry>(), withoutDesc2);
            Assert.Equal(SwpAlarmVerdict.Ok, clean.AlarmVerdict);
        }
    }
}
