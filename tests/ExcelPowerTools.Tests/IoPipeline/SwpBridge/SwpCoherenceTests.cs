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
        public void Plant_inputs_matching_ignores_case_like_every_other_mnemonic_match()
        {
            // The containment match is legacy and deliberate; its case-sensitivity was not.
            CsvDocument lowercase = Doc(
                "#;Signal ID;Type;Logic;RE;FE;A1;A0;Events;Queue;Flags;Node;Offset;Bit;Default;Card;#;",
                "@;cabinetal00100;1;1;0;0;0;0;0;0;0;80;3;1;0;1;#;desc");
            SwpSpineRow row = SignalRow();
            SwpCoherence.Annotate(Model(row), lowercase, null, SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpCsvVerdict.Ok, row.CsvVerdict);
            Assert.Equal(2, row.CsvLine);
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

        /// <summary>The real FVZ ISPC2 IODriver\BusField.csv, cut down: the modern layout, reference
        /// index / element / bit all 0 on every row — the nodes the check used to report missing.</summary>
        private static CsvDocument Ispc2Busfield() => Doc(
            "#;Board Number;Physical Node Address;Node Type;1 = Warning on fault;Linked Node Address;Linked Master ID;Linked Slave ID;Reference Index;DB Class ID;Element;Bit;#;# Description",
            "@;1;0;0x011E0203;0;0;0;0;0;0;0;0;#;Hilscher Master Profinet Card 1",
            "#;1;6;0x002A0604;0;0;0;0;0;0;0;0;#;[n0006-mc1-cc1-k60201] PN-PN COUPLER (SIEMENS)  +MC1.CC1-K60201 (192.168.3.6)",
            "@;1;61;0x0136AC50;0;0;0;0;0;0;0;0;#;[n0061-in21-xn1] DISTRIBUTED I-O LINK DEVICE (IFM)  +IN21-XN1 (192.168.3.61)",
            "@;1;101;0x010A0006;0;0;0;0;0;0;0;0;#;[n0101-in21-01-t1] MOVIMOT ADVANCE (SEW-EURODRIVE)  +IN21.01-T1 (192.168.3.101)",
            "@;1;111;0x010A0006;0;0;0;0;0;0;0;0;#;[n0111-in22-01-t1] MOVIMOT ADVANCE (SEW-EURODRIVE)  +IN22.01-T1 (192.168.3.111)");

        private static SwpSpineRow ModuleRowOn(string node, string mnemonic, SwpPlantType plantType)
        {
            SwpSpineRow row = ModuleRow(mnemonic, plantType);
            row.Entry.Node = node;
            return row;
        }

        [Theory]
        [InlineData("61", "CABINETAL100162")]    // induction line 1: IN21's I/O-Link master
        [InlineData("101", "CABINETAL100162")]   // …and its first MOVIMOT
        [InlineData("111", "CABINETAL200162")]   // induction line 2: IN22's first MOVIMOT
        public void A_modern_induction_node_is_found_on_node_and_card_alone(string node, string mnemonic)
        {
            SwpSpineRow row = ModuleRowOn(node, mnemonic, SwpPlantType.Ispc);
            // A LINE row: exactly what the old rule compared against the busfield's index 0.
            Assert.True(row.Mapped.Ok);
            Assert.NotEqual(0, row.Mapped.ReferenceIndex);
            var log = new List<SwpLogEntry>();

            SwpCoherence.Annotate(Model(row), null, Ispc2Busfield(), SwpPlantType.Ispc, log);

            Assert.Equal(SwpBusfieldVerdict.Ok, row.BusfieldVerdict);
            Assert.DoesNotContain(log, l => l.Message.Contains("not found in the busfield"));
        }

        [Fact]
        public void A_modern_upstream_node_is_found_without_a_reference_index()
        {
            // The FUR_Jadlog group busfields: every row index 0, element 0, bit 0.
            CsvDocument busfield = Doc("#;h;h;h;h;h;h;h;h;h;h;h;#;h", "@;1;236;32880;0;0;0;0;0;0;0;0;#;x");
            SwpSpineRow row = ModuleRow("ALLQ_2_002_00", SwpPlantType.Ups);
            Assert.Equal(2, row.Mapped.ReferenceIndex);

            SwpCoherence.Annotate(Model(row), null, busfield, SwpPlantType.Ups, new List<SwpLogEntry>());

            Assert.Equal(SwpBusfieldVerdict.Ok, row.BusfieldVerdict);
        }

        [Fact]
        public void A_node_the_busfield_comments_out_or_lacks_is_still_missing()
        {
            // "#" rows are disabled in the IODriver: node 6 is not on the network, 121 never was.
            SwpSpineRow disabled = ModuleRowOn("6", "CABINETAL100162", SwpPlantType.Ispc);
            SwpSpineRow absent = ModuleRowOn("121", "CABINETAL100162", SwpPlantType.Ispc);

            SwpCoherence.Annotate(Model(disabled, absent), null, Ispc2Busfield(), SwpPlantType.Ispc, new List<SwpLogEntry>());

            Assert.Equal(SwpBusfieldVerdict.NodeMissing, disabled.BusfieldVerdict);
            Assert.Equal(SwpBusfieldVerdict.NodeMissing, absent.BusfieldVerdict);
        }

        [Fact]
        public void A_legacy_induction_row_still_has_to_name_its_line()
        {
            // The deprecated layout carries element/bit in the busfield row, and there the index
            // routes the fault to ONE line: line 5 is not this row's line.
            CsvDocument busfield = Doc("#;h;h;h;h;h;h;h;h;h;h;h;#;h", "@;1;236;32880;0;0;0;0;5;21;1;62;#;x");
            SwpSpineRow row = ModuleRow("CABINETAL100162", SwpPlantType.Ispc);

            SwpCoherence.Annotate(Model(row), null, busfield, SwpPlantType.Ispc, new List<SwpLogEntry>());

            Assert.Equal(SwpBusfieldVerdict.NodeMissing, row.BusfieldVerdict);
        }

        /// <summary>The real FNU Zalando IODriver\BusField.csv head: the OLDER layout — no Linked
        /// Master/Slave ID, so Reference Index / DB Class / Element / Bit sit two fields earlier,
        /// behind two parameter lines, with the export's own "Phisical" spelling.</summary>
        private static CsvDocument OlderLayoutBusfield(string row) => Doc(
            "#;56;;;;;;;;;#;;;;;#",
            "#;Para #1;Para #2;Para #3;Para #4;Para #5;Para #6;Para #7;Para #8;Para #9;#;;;;;#",
            "#;Board Number;Phisical Node Address;Node Type;1 = Warning on fault;Linked Node Address;Reference Index;DB Class ID;Element;Bit;#;Profinet IP;Profinet name;Desc;;#",
            row);

        [Fact]
        public void An_older_busfield_layout_is_read_by_its_own_column_names()
        {
            // Real FNU node 5: index, class, element and bit all 0 — the modern linkage. Fixed
            // positions read the "#" separator as element and the IP address as bit: DataMismatch.
            SwpSpineRow row = ModuleRowOn("5", "CABINETAL00063", SwpPlantType.Msc);
            var log = new List<SwpLogEntry>();

            SwpCoherence.Annotate(Model(row), null,
                OlderLayoutBusfield("@;1;5;0x013418C5;0;0;0;0;0;0;#;192.168.1.5;N0005+MS1-K2501;UR20;CABINETAL00063;#"),
                SwpPlantType.Msc, log);

            Assert.Equal(SwpBusfieldVerdict.Ok, row.BusfieldVerdict);
            Assert.DoesNotContain(log, l => l.Message.Contains("disagrees") || l.Message.Contains("missing columns"));
        }

        [Fact]
        public void An_older_layout_reference_index_is_read_where_that_layout_keeps_it()
        {
            // Index 2 in field 6 (the older position) with element/bit set: the legacy linkage, which
            // routes the fault to line 2 — found for the line-2 row, missing for the line-1 row.
            CsvDocument busfield = OlderLayoutBusfield("@;1;236;0x010C0002;0;0;2;21;1;62;#;192.168.1.236;N0236;x;x;#");
            SwpSpineRow lineTwo = ModuleRow("CABINETAL200162", SwpPlantType.Ispc);
            SwpSpineRow lineOne = ModuleRow("CABINETAL100162", SwpPlantType.Ispc);
            Assert.Equal(2, lineTwo.Mapped.ReferenceIndex);
            Assert.Equal(1, lineOne.Mapped.ReferenceIndex);

            SwpCoherence.Annotate(Model(lineTwo, lineOne), null, busfield, SwpPlantType.Ispc, new List<SwpLogEntry>());

            Assert.NotEqual(SwpBusfieldVerdict.NodeMissing, lineTwo.BusfieldVerdict);
            Assert.Equal(SwpBusfieldVerdict.NodeMissing, lineOne.BusfieldVerdict);
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
