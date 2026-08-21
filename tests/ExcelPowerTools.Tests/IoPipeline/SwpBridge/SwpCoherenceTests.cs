using System.Collections.Generic;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpCoherenceTests
    {
        private static CsvDocument Doc(params string[] lines) =>
            CsvFile.Load(Encoding.UTF8.GetBytes(string.Join("\r\n", lines)));

        private static CsvDocument PlantInputs() => Doc(
            "#;Signal ID;Type;Logic;RE;FE;A1;A0;Events;Queue;Flags;Node;Offset;Bit;Default;Card;#;",
            "@;CABINETAL00100;1;1;0;0;0;0;0;0;0;80;3;1;0;1;#;desc",
            "@;OTHERSIGNAL_01;1;1;0;0;0;0;0;0;0;90;0;0;0;1;#;other");

        private static CsvDocument Busfield() => Doc(
            "#;Board Number;Phisical Node Address;Node Type;1 = Warning on fault;a;b;c;Reference Index;DB Class ID;Element;Bit;#;TEXT",
            "@;1;236;32880;0;0;0;0;0;60;1;61;#;Controller");

        private static SwpSpineRow SignalRow(string mnemonic = "CABINETAL00100",
            string node = "80", string offset = "3", string bit = "1")
        {
            return new SwpSpineRow
            {
                Entry = new DiagnosticEntry
                {
                    Mnemonic = mnemonic,
                    Node = node,
                    Offset = offset,
                    Bit = bit,
                    CardSheet = "R1.S1.C1",
                },
                Mapped = SwpAddressMap.Map(mnemonic, SwpPlantType.Msc),
            };
        }

        private static SwpSpineRow ModuleRow(string mnemonic = "CABINETAL00161",
            SwpPlantType plantType = SwpPlantType.Msc, string cardSheet = "R1.S1.C1")
        {
            return new SwpSpineRow
            {
                Entry = new DiagnosticEntry
                {
                    Mnemonic = mnemonic,
                    IsBusModule = true,
                    Node = "236",
                    CardSheet = cardSheet,
                },
                Mapped = SwpAddressMap.Map(mnemonic, plantType),
            };
        }

        private static SwpWorkspaceModel Model(params SwpSpineRow[] rows)
        {
            var model = new SwpWorkspaceModel();
            model.Rows.AddRange(rows);
            return model;
        }

        [Fact]
        public void A_matching_signal_reads_ok()
        {
            SwpSpineRow row = SignalRow();
            SwpCoherence.Annotate(Model(row), PlantInputs(), null, SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpCsvVerdict.Ok, row.CsvVerdict);
            Assert.Equal(SwpBusfieldVerdict.NotChecked, row.BusfieldVerdict);
        }

        [Fact]
        public void A_moved_address_reads_mismatch_with_an_error()
        {
            SwpSpineRow row = SignalRow(node: "81");
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(Model(row), PlantInputs(), null, SwpPlantType.Msc, log);
            Assert.Equal(SwpCsvVerdict.AddressMismatch, row.CsvVerdict);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("Plant_Inputs"));
        }

        [Fact]
        public void A_missing_mnemonic_reads_name_missing()
        {
            SwpSpineRow row = SignalRow("CABINETAL00900");
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(Model(row), PlantInputs(), null, SwpPlantType.Msc, log);
            Assert.Equal(SwpCsvVerdict.NameMissing, row.CsvVerdict);
            Assert.Contains(log, l => l.Message.Contains("not found in Plant_Inputs"));
        }

        [Fact]
        public void A_consistent_module_reads_ok_in_the_busfield()
        {
            SwpSpineRow row = ModuleRow();   // element 1, bit 61, alarm — matches the fixture row
            SwpCoherence.Annotate(Model(row), null, Busfield(), SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpBusfieldVerdict.Ok, row.BusfieldVerdict);
            Assert.Equal(SwpCsvVerdict.NotChecked, row.CsvVerdict);
        }

        [Fact]
        public void A_node_absent_from_the_busfield_warns()
        {
            SwpSpineRow row = ModuleRow();
            row.Entry.Node = "999";
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(Model(row), null, Busfield(), SwpPlantType.Msc, log);
            Assert.Equal(SwpBusfieldVerdict.NodeMissing, row.BusfieldVerdict);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("999"));
        }

        [Fact]
        public void A_wrong_card_is_a_missing_node()
        {
            SwpSpineRow row = ModuleRow(cardSheet: "R1.S1.C2");
            SwpCoherence.Annotate(Model(row), null, Busfield(), SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpBusfieldVerdict.NodeMissing, row.BusfieldVerdict);
        }

        [Fact]
        public void Disagreeing_element_or_bit_reads_data_mismatch()
        {
            SwpSpineRow row = ModuleRow("CABINETAL00262");   // element 2 ≠ busfield element 1
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(Model(row), null, Busfield(), SwpPlantType.Msc, log);
            Assert.Equal(SwpBusfieldVerdict.DataMismatch, row.BusfieldVerdict);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("disagrees"));
        }

        [Fact]
        public void A_warning_flag_disagreement_is_a_data_mismatch()
        {
            SwpSpineRow row = ModuleRow("CABINETWR00161");   // warning mnemonic vs busfield flag 0
            SwpCoherence.Annotate(Model(row), null, Busfield(), SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpBusfieldVerdict.DataMismatch, row.BusfieldVerdict);
        }

        [Fact]
        public void Upstream_requires_the_reference_index_to_agree()
        {
            CsvDocument busfield = Doc(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;236;32880;0;0;0;0;2;11;2;0;#;x");   // reference index 2, element 2

            SwpSpineRow matching = ModuleRow("ALLQ_2_002_00", SwpPlantType.Ups);   // system 2
            SwpSpineRow foreign = ModuleRow("ALLQ_1_002_00", SwpPlantType.Ups);    // system 1
            SwpCoherence.Annotate(Model(matching, foreign), null, busfield, SwpPlantType.Ups, new List<SwpLogEntry>());

            Assert.Equal(SwpBusfieldVerdict.Ok, matching.BusfieldVerdict);
            Assert.Equal(SwpBusfieldVerdict.NodeMissing, foreign.BusfieldVerdict);
        }

        [Fact]
        public void An_induction_line_mask_is_accepted_and_decoded()
        {
            CsvDocument busfield = Doc(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;236;32880;0;0;0;0;10;21;1;62;#;x");   // 10 ≥ 10 → a mask: L2 + L4

            SwpSpineRow row = ModuleRow("CABINETAL100162", SwpPlantType.Ispc);
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(Model(row), null, busfield, SwpPlantType.Ispc, log);

            Assert.Equal(SwpBusfieldVerdict.Ok, row.BusfieldVerdict);
            Assert.Contains(log, l => l.Message.Contains("L2 L4"));
        }

        [Fact]
        public void Generated_rows_and_absent_documents_stay_unchecked()
        {
            SwpSpineRow generated = SignalRow();
            generated.Entry.GeneratedBy = "AP";
            SwpSpineRow unchecked1 = SignalRow();

            SwpCoherence.Annotate(Model(generated), PlantInputs(), Busfield(), SwpPlantType.Msc, new List<SwpLogEntry>());
            SwpCoherence.Annotate(Model(unchecked1), null, null, SwpPlantType.Msc, new List<SwpLogEntry>());

            Assert.Equal(SwpCsvVerdict.NotChecked, generated.CsvVerdict);
            Assert.Equal(SwpCsvVerdict.NotChecked, unchecked1.CsvVerdict);
        }
    }
}
