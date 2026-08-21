using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpUpsGeneratorsTests
    {
        // ---- photocells ---------------------------------------------------------------------

        private static IoRowData PhtRow(string mnemonic, string address, string location, string device)
        {
            return new IoRowData
            {
                CardSheet = "R1.S1.C1",
                SourceRow = 42,
                MnemonicCell = mnemonic,
                AddressCell = address,
                FUnit = "=S1",
                Location = location,
                Device = device,
            };
        }

        [Fact]
        public void Photocells_sort_by_belt_position_and_claim_bits_from_50()
        {
            string[] sensor =
            {
                "hdr;;;;;;;",
                ";@;PREV_BELT_ID;1;7;2;0;5",       // photocell 2 on belt 5
                ";@;PREV_BELT_ID;1;7;9;0;5",       // photocell 9 on belt 5
                ";@;OFFSET_ON_BELT_Y;1;7;9;0;PB100",  // 9 sits BEFORE 2 on the belt
                ";@;OFFSET_ON_BELT_Y;1;7;2;0;NE50",
            };
            var io = new[]
            {
                PhtRow("PHT_1_002_000", "I.112.7.3", "+UL1101.01", "-B2"),
                PhtRow("PHT_1_009_000/OTHER_NAME", "I.112.7.0", "+UL1101.01", "-B9"),
            };

            var log = new List<SwpLogEntry>();
            List<DiagnosticEntry> entries = SwpUpsGenerators.GeneratePhotocellAlarms(sensor, io, log);

            Assert.Equal(2, entries.Count);
            Assert.Equal("ALLB_1_005_50", entries[0].Mnemonic);   // position-sorted: photocell 9 first
            Assert.Equal("ALLB_1_005_51", entries[1].Mnemonic);

            DiagnosticEntry first = entries[0];
            Assert.Equal("AP", first.GeneratedBy);
            Assert.Equal("TS_031", first.TsRef);
            Assert.Equal("112", first.Node);
            Assert.Equal("7", first.Offset);
            Assert.Equal("0", first.Bit);
            Assert.Equal("-B9", first.Device);
            Assert.Equal("Sensor Fault #1 +UL1101.01-B9", first.Desc1A);
            Assert.Equal("R1.S1.C1", first.CardSheet);
            Assert.Equal("Sensor Fault #2 +UL1101.01-B2", entries[1].Desc1A);
            Assert.Empty(log);
        }

        [Fact]
        public void A_position_line_before_its_belt_line_still_lands()
        {
            string[] sensor =
            {
                ";@;OFFSET_ON_BELT_Y;1;7;2;0;PB100",   // arrives first
                ";@;PREV_BELT_ID;1;7;2;0;5",
            };
            List<DiagnosticEntry> entries = SwpUpsGenerators.GeneratePhotocellAlarms(
                sensor, new IoRowData[0], new List<SwpLogEntry>());
            Assert.Single(entries);
        }

        [Fact]
        public void A_missing_pht_row_warns_and_leaves_no_address()
        {
            string[] sensor = { ";@;PREV_BELT_ID;1;7;2;0;5" };
            var log = new List<SwpLogEntry>();
            List<DiagnosticEntry> entries = SwpUpsGenerators.GeneratePhotocellAlarms(
                sensor, new IoRowData[0], log);

            DiagnosticEntry entry = Assert.Single(entries);
            Assert.Equal("", entry.SwpAddress());
            Assert.Equal("Sensor Fault #1", entry.Desc1A);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("PHT_1_002_000"));
        }

        [Fact]
        public void An_unassigned_sensor_position_warns()
        {
            string[] sensor = { ";@;OFFSET_ON_BELT_Y;1;7;2;0;PB100" };
            var log = new List<SwpLogEntry>();
            SwpUpsGenerators.GeneratePhotocellAlarms(sensor, new IoRowData[0], log);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("no belt"));
        }

        [Theory]
        [InlineData("PB100", 10100)]
        [InlineData("PE250", 20250)]
        [InlineData("NB250", 20250)]
        [InlineData("NE10", 40010)]
        [InlineData("XX10", 0)]
        [InlineData("PBabc", 0)]
        public void Absolute_positions_follow_the_half_belt_arithmetic(string offset, long expected)
        {
            Assert.Equal(expected, SwpUpsGenerators.AbsolutePosition(offset));
        }

        // ---- strobe devices -----------------------------------------------------------------

        [Fact]
        public void Strobe_devices_expand_their_template()
        {
            string[] strobe =
            {
                ";@;TYPE_ID;2;0;3;0;1",
                ";@;REFERENCE_BELT_ID;2;0;3;0;12",
                ";@;REFERENCE_CABINET_ID;2;0;3;0;4",
            };
            var templates = new List<SwpStrobeTemplate>
            {
                new SwpStrobeTemplate
                {
                    TypeId = "1",
                    TypeDescription = "Scale",
                    Mnemonics = { "WARQ_%sys%_%cabinet%_01", "ALLB_%sys%_%belt%_02" },
                    Descriptions = { "SCALE CERT", "BELT THING" },
                },
            };

            List<DiagnosticEntry> entries = SwpUpsGenerators.GenerateStrobeAlarms(
                strobe, templates, new List<SwpLogEntry>());

            Assert.Equal(2, entries.Count);
            Assert.Equal("WARQ_2_04_01", entries[0].Mnemonic);
            Assert.Equal("SCALE CERT", entries[0].Desc1A);
            Assert.Equal("ASD", entries[0].GeneratedBy);
            Assert.Equal("ALLB_2_012_02", entries[1].Mnemonic);
        }

        [Fact]
        public void An_unknown_strobe_type_logs_an_error()
        {
            string[] strobe = { ";@;TYPE_ID;1;0;3;0;9" };
            var log = new List<SwpLogEntry>();
            List<DiagnosticEntry> entries = SwpUpsGenerators.GenerateStrobeAlarms(
                strobe, SwpDefaults.DefaultStrobeTemplates(), log);
            Assert.Empty(entries);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("type 9"));
        }

        [Fact]
        public void The_harvested_default_templates_are_complete()
        {
            List<SwpStrobeTemplate> templates = SwpDefaults.DefaultStrobeTemplates();
            Assert.Equal(2, templates.Count);
            Assert.Equal(14, templates[0].Mnemonics.Count);       // Scale: 7 WARQ + 7 ALLQ
            Assert.Equal(120, templates[1].Mnemonics.Count);      // BottomCamera: 60 ALLQ + 60 WARQ
            Assert.All(templates, t => Assert.Equal(t.Mnemonics.Count, t.Descriptions.Count));
        }
    }
}
