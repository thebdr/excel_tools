using System.Collections.Generic;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The fourth coherence leg: BUS_ExtendedDiagnostic.csv binds AS-I slave states to CLASSDIAG
    /// mnemonics ("@;mnemonic;card;gateway node;sub-device;type;mask;# comment"). A slave's
    /// mnemonic must be bound, to its gateway and to the sub-device its line + address fix.
    /// </summary>
    public class SwpBusExtendedDiagnosticTests
    {
        private static CsvDocument Doc(params string[] lines) =>
            CsvFile.Load(Encoding.UTF8.GetBytes(string.Join("\r\n", lines)));

        private static CsvDocument Extended() => Doc(
            "#;Para #1;Para #2;Para #3;Para #4;Para #5;Para #6;Para #7;#",
            "#;Signal to set;Card;Node;Sub Device 0=master 1..127=device;Type 0=Bit 1=Mask;Bit or HEX Value;Comment;#",
            "#Suggestion:;xxx;xxx;xxx;xxx;1;FFFFFFFF;xxx;#",
            "@;CLASSDIAGAL00620030130;1;90;1;1;00000028;# Asi slave fault master:90 channel:1 slave:1A chute 301 +CH301-K1;#",
            "@;CLASSDIAGAL00620030230;1;90;1;1;00000028;# Asi slave fault master:90 channel:1 slave:1A chute 302 +CH301-K1;#",
            "@;CLASSDIAGAL00620032630;1;90;13;1;00000028;# Asi slave fault master:90 channel:1 slave:13A chute 326 +CH326-K1;#",
            "@;CLASSDIAGAL00620020130;1;91;65;1;00000028;# Asi slave fault master:91 channel:2 slave:1A chute 201 +CH201-K1;#");

        private static CsvDocument Busfield() => Doc(
            "#;Board Number;Physical Node Address;Node Type;1 = Warning on fault;Linked Node Address;Linked Master ID;Linked Slave ID;Reference Index;DB Class ID;Element;Bit;#;# Description",
            "@;1;0;0x011E0203;0;0;0;0;0;0;0;0;#;Hilscher Master Profinet Card 1",
            "@;1;90;0x0136AC01;2;0;0;0;0;0;0;0;#;[n0090-pc1-cc1-k35601] GATEWAY PROFINET - AS-i");

        private static SwpSpineRow Slave(string mnemonic, string asiNode, string master, string card = "R1.S1.C1")
        {
            return new SwpSpineRow
            {
                Entry = new DiagnosticEntry
                {
                    CardSheet = card,
                    SourceRow = 619,
                    RawMnemonic = mnemonic,
                    Mnemonic = mnemonic,
                    IsBusModule = true,
                    Node = asiNode.Substring(asiNode.IndexOf('.') + 1),
                    AsiNode = asiNode,
                    AsiMasterNode = master,
                },
            };
        }

        private static SwpWorkspaceModel Model(params SwpSpineRow[] rows)
        {
            var model = new SwpWorkspaceModel();
            model.Rows.AddRange(rows);
            return model;
        }

        private static List<SwpLogEntry> Check(params SwpSpineRow[] rows)
        {
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(Model(rows), null, null, SwpPlantType.Msc, log, busExtended: Extended());
            return log;
        }

        [Fact]
        public void Bound_slaves_read_ok_on_both_gateway_lines()
        {
            SwpSpineRow a = Slave("CLASSDIAGAL00620030130", "1M.1A", "90");
            SwpSpineRow b = Slave("CLASSDIAGAL00620030230", "1M.1A", "90");
            SwpSpineRow c = Slave("CLASSDIAGAL00620020130", "2M.1A", "91");   // line 2 → sub-device 65
            SwpSpineRow d = Slave("CLASSDIAGAL00620032630", "1M.13A", "90");

            List<SwpLogEntry> log = Check(a, b, c, d);

            Assert.All(new[] { a, b, c, d }, row => Assert.Equal(SwpBusExtVerdict.Ok, row.BusExtVerdict));
            Assert.DoesNotContain(log, l => l.Severity == SwpLogSeverity.Error);

            // The matched line is kept for the double-click jump (three header lines first).
            Assert.Equal(4, a.BusExtLine);
            Assert.Equal(5, b.BusExtLine);
            Assert.Equal(7, c.BusExtLine);
            Assert.Equal(6, d.BusExtLine);
        }

        [Fact]
        public void Plant_inputs_and_a_missing_binding_keep_the_line_or_leave_it_zero()
        {
            CsvDocument plantInputs = Doc(
                "#;Signal;x;x;x;x;x;x;x;x;x;Node;Offset;Bit",
                "@;S1.CABINETAL00200;;;;;;;;;;81;2;7",
                "@;S1.CABINETAL00100;;;;;;;;;;80;3;1");
            var signal = new SwpSpineRow
            {
                Entry = new DiagnosticEntry
                {
                    CardSheet = "R1.S1.C1", SourceRow = 12, Mnemonic = "CABINETAL00100",
                    Node = "80", Offset = "3", Bit = "1",
                },
            };
            SwpSpineRow unbound = Slave("CLASSDIAGAL00620039930", "1M.13A", "90");

            SwpCoherence.Annotate(Model(signal, unbound), plantInputs, null, SwpPlantType.Msc,
                new List<SwpLogEntry>(), busExtended: Extended());

            Assert.Equal(SwpCsvVerdict.Ok, signal.CsvVerdict);
            Assert.Equal(3, signal.CsvLine);
            Assert.Equal(SwpBusExtVerdict.NotFound, unbound.BusExtVerdict);
            Assert.Equal(0, unbound.BusExtLine);
        }

        [Fact]
        public void An_unbound_mnemonic_reads_not_bound_with_the_expected_sub_device()
        {
            SwpSpineRow row = Slave("CLASSDIAGAL00620039930", "1M.13A", "90");

            List<SwpLogEntry> log = Check(row);

            Assert.Equal(SwpBusExtVerdict.NotFound, row.BusExtVerdict);
            SwpLogEntry error = Assert.Single(log, l => l.Severity == SwpLogSeverity.Error);
            Assert.Equal("AS-I slave 1M.13A (sub-device 13): mnemonic not bound in BUS_ExtendedDiagnostic", error.Message);
            Assert.Equal("CLASSDIAGAL00620039930", error.Mnemonic);
            Assert.Equal(619, error.SourceRow);
        }

        [Fact]
        public void A_slave_bound_to_another_gateway_reads_node_differs()
        {
            SwpSpineRow row = Slave("CLASSDIAGAL00620030130", "1M.1A", "91");

            List<SwpLogEntry> log = Check(row);

            Assert.Equal(SwpBusExtVerdict.NodeMismatch, row.BusExtVerdict);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error &&
                l.Message == "BUS_ExtendedDiagnostic binds it to gateway node 90; the IO list hangs slave 1M.1A off node 91");
        }

        [Fact]
        public void A_slave_on_the_wrong_address_or_line_reads_sub_device_differs()
        {
            SwpSpineRow wrongAddress = Slave("CLASSDIAGAL00620030130", "1M.2A", "90");   // bound to 1
            SwpSpineRow wrongLine = Slave("CLASSDIAGAL00620032630", "1M.13B", "90");     // B line → 45, bound to 13

            List<SwpLogEntry> log = Check(wrongAddress, wrongLine);

            Assert.Equal(SwpBusExtVerdict.SubDeviceMismatch, wrongAddress.BusExtVerdict);
            Assert.Equal(SwpBusExtVerdict.SubDeviceMismatch, wrongLine.BusExtVerdict);
            Assert.Contains(log, l => l.Message.StartsWith("BUS_ExtendedDiagnostic sub-device 1 vs 2 for AS-I slave 1M.2A"));
            Assert.Contains(log, l => l.Message.StartsWith("BUS_ExtendedDiagnostic sub-device 13 vs 45 for AS-I slave 1M.13B"));
        }

        [Fact]
        public void A_slave_on_another_card_reads_node_differs()
        {
            SwpSpineRow row = Slave("CLASSDIAGAL00620030130", "1M.1A", "90", card: "R1.S1.C2");

            List<SwpLogEntry> log = Check(row);

            Assert.Equal(SwpBusExtVerdict.NodeMismatch, row.BusExtVerdict);
            Assert.Contains(log, l => l.Message == "BUS_ExtendedDiagnostic binds it to card 1; the IO list has slave 1M.1A on card 2");
        }

        [Fact]
        public void An_unknown_gateway_passes_with_an_info_line()
        {
            SwpSpineRow row = Slave("CLASSDIAGAL00620030130", "1M.1A", "");

            List<SwpLogEntry> log = Check(row);

            Assert.Equal(SwpBusExtVerdict.Ok, row.BusExtVerdict);
            Assert.DoesNotContain(log, l => l.Severity == SwpLogSeverity.Error);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info && l.Message.Contains("no gateway row above it"));
        }

        [Fact]
        public void Without_the_file_the_verdict_stays_not_checked()
        {
            SwpSpineRow row = Slave("CLASSDIAGAL00620030130", "1M.1A", "90");
            SwpCoherence.Annotate(Model(row), null, null, SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpBusExtVerdict.NotChecked, row.BusExtVerdict);
        }

        [Fact]
        public void The_busfield_check_skips_asi_slaves()
        {
            // "1A" is not a Profinet node — the busfield never lists it, so it must not warn.
            SwpSpineRow row = Slave("CLASSDIAGAL00620030130", "1M.1A", "90");
            row.Mapped = new SwpMappedMnemonic { ObjectType = "CHUTE", RawDeviceNumber = 301 };
            var log = new List<SwpLogEntry>();

            SwpCoherence.Annotate(Model(row), null, Busfield(), SwpPlantType.Msc, log);

            Assert.Equal(SwpBusfieldVerdict.NotChecked, row.BusfieldVerdict);
            Assert.DoesNotContain(log, l => l.Message.Contains("not found in the busfield"));
        }
    }
}
