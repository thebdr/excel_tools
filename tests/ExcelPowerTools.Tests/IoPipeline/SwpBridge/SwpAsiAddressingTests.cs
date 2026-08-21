using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// AS-I identities: the Slot.ID composition ("1M.24A", empty slot = line 1) and the gateway
    /// image layout that numbers sub-devices — the IO list's "ASI Addressing" sheet, kept in code
    /// because not every IO list ships it (line 1: A 1..31, B 33..63; line 2: A 65..95, B 97..127).
    /// </summary>
    public class SwpAsiAddressingTests
    {
        [Theory]
        [InlineData("1M", "24A", "1M.24A")]
        [InlineData("", "1A", "1M.1A")]          // empty slot = line 1
        [InlineData("  ", "1A", "1M.1A")]
        [InlineData("2m", "3b", "2M.3B")]        // case-normalized
        [InlineData("2M", "31B", "2M.31B")]
        [InlineData("3M", "1A", "")]             // a gateway has two lines
        [InlineData("-K35601", "1A", "")]        // a device tag is not a line
        [InlineData("1M", "32A", "")]            // addresses run 1..31
        [InlineData("1M", "0A", "")]
        [InlineData("1M", "1C", "")]
        [InlineData("1M", "90", "")]             // a Profinet node, not a slave
        [InlineData("1M", "", "")]
        public void Compose_builds_line_dot_slave(string slot, string id, string expected)
        {
            Assert.Equal(expected, SwpAsiAddressing.Compose(slot, id));
        }

        [Theory]
        [InlineData("1M.1A", 1)]
        [InlineData("1M.24A", 24)]
        [InlineData("1M.31A", 31)]
        [InlineData("1M.1B", 33)]
        [InlineData("1M.24B", 56)]
        [InlineData("1M.31B", 63)]
        [InlineData("2M.1A", 65)]
        [InlineData("2M.24A", 88)]
        [InlineData("2M.31A", 95)]
        [InlineData("2M.1B", 97)]
        [InlineData("2M.22B", 118)]
        [InlineData("2M.31B", 127)]
        [InlineData("1m.3a", 3)]
        [InlineData("1M.MasterFlags", -1)]
        [InlineData("1A", -1)]
        [InlineData("", -1)]
        [InlineData(null, -1)]
        public void SubDevice_follows_the_gateway_image_layout(string node, int expected)
        {
            Assert.Equal(expected, SwpAsiAddressing.SubDevice(node));
        }

        [Fact]
        public void The_filter_composes_the_slave_identity_and_keeps_the_gateway_node()
        {
            var row = new IoRowData
            {
                CardSheet = "R1.S1.C1",
                SourceRow = 619,
                MnemonicCell = "CLASSDIAGAL00620030130/CLASSDIAGAL00620030230",
                AddressCell = "",
                IdCell = "1A",
                SlotCell = "1M",
                ModuleNode = "90",
            };
            var log = new List<SwpLogEntry>();
            List<DiagnosticEntry> entries = SwpMnemonicFilter.Run(new[] { row },
                new SwpFilterOptions { PlantType = SwpPlantType.Msc }, log);

            Assert.Equal(2, entries.Count);
            foreach (DiagnosticEntry e in entries)
            {
                Assert.True(e.IsBusModule);
                Assert.Equal("1A", e.Node);
                Assert.Equal("1M.1A", e.SwpAddress());    // G receives line.slave — no ".0.0"
                Assert.Equal("1M.1A", e.AsiNode);
                Assert.Equal("90", e.AsiMasterNode);
            }
        }

        [Fact]
        public void An_empty_slot_is_line_one_and_a_foreign_slot_warns()
        {
            var log = new List<SwpLogEntry>();
            DiagnosticEntry bare = Assert.Single(SwpMnemonicFilter.Run(
                new[] { new IoRowData { CardSheet = "R1.S1.C1", SourceRow = 5, MnemonicCell = "CLASSDIAGAL00620030130", IdCell = "24a" } },
                new SwpFilterOptions { PlantType = SwpPlantType.Msc }, log));
            Assert.Equal("1M.24A", bare.AsiNode);
            Assert.Equal("", bare.AsiMasterNode);
            Assert.DoesNotContain(log, l => l.Severity == SwpLogSeverity.Warning);

            DiagnosticEntry foreign = Assert.Single(SwpMnemonicFilter.Run(
                new[] { new IoRowData { CardSheet = "R1.S1.C1", SourceRow = 6, MnemonicCell = "CLASSDIAGAL00620030130", IdCell = "2B", SlotCell = "-K35601" } },
                new SwpFilterOptions { PlantType = SwpPlantType.Msc }, log));
            Assert.Equal("", foreign.AsiNode);
            Assert.Equal("2B", foreign.Node);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("-K35601"));
        }
    }
}
