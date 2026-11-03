using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline
{
    public class IoListLogicTests
    {
        [Theory]
        [InlineData("R1.S1.C1", "1", "1", "1")]              // 8FVZ_IO0001 sorter
        [InlineData("R1.S2.C1 - PC2", "1", "2", "1")]        // 8FVZ_IO0002 induction: label suffix
        [InlineData("R1.S3.C1 - PC3 ", "1", "3", "1")]       // ...and a trailing space
        [InlineData("R1.S12.C10", "1", "12", "10")]          // multi-digit rack/slot/card
        [InlineData(" r2.s3.c04 ", "2", "3", "4")]           // case + zero padding
        public void Parses_card_from_sheet_name(string sheetName, string rack, string slot, string card)
        {
            Assert.True(IoListLogic.TryParseCardSheetName(sheetName, out string r, out string s, out string c));
            Assert.Equal(rack, r);
            Assert.Equal(slot, s);
            Assert.Equal(card, c);
        }

        [Theory]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("COVER")]
        [InlineData("NET SAFETY 50")]
        [InlineData("R1.S1")]              // no card
        [InlineData("Chutes R1.S1.C1")]    // must start with the card token, not merely contain it
        public void Rejects_non_card_sheet_names(string sheetName)
        {
            Assert.False(IoListLogic.TryParseCardSheetName(sheetName, out _, out _, out _));
        }

        [Theory]
        [InlineData("I.2.0.0", 'I', "2", "0", "0")]
        [InlineData("Q.2.0.0", 'O', "2", "0", "0")]     // Siemens Q == output
        [InlineData("Q.90.1.4", 'O', "90", "1", "4")]
        [InlineData("O.0080.3.7", 'O', "80", "3", "7")] // padded components normalize
        [InlineData("i.5.0.15", 'I', "5", "0", "15")]
        public void Parses_addresses(string input, char dir, string node, string offset, string bit)
        {
            Assert.True(IoListLogic.TryParseAddress(input, out char d, out string n, out string o, out string b));
            Assert.Equal(dir, d);
            Assert.Equal(node, n);
            Assert.Equal(offset, o);
            Assert.Equal(bit, b);
        }

        [Theory]
        [InlineData("80.3.0")]        // no direction
        [InlineData("I.80.3")]        // too short
        [InlineData("I.80.3.0.1")]    // too long
        [InlineData("X.80.3.0")]      // bad direction
        [InlineData("I.a.3.0")]       // non-numeric
        [InlineData("I0.0")]          // Siemens byte.bit — a PLC address, not an ASC one
        [InlineData("")]
        [InlineData(null)]
        public void Rejects_non_addresses(string input)
        {
            Assert.False(IoListLogic.TryParseAddress(input, out _, out _, out _, out _));
        }

        // The safety PLC's lists address inputs directly — no node, because a distributed
        // input is unique PLC-wide: "I.0.4" on the FVT list, "I0.4" on the FVZ one, "%I12.7"
        // as TIA exports it, E/A in German projects. Q/A are outputs.
        [Theory]
        [InlineData("I.0.4", 'I', "0", "4")]
        [InlineData("I0.4", 'I', "0", "4")]
        [InlineData("%I12.7", 'I', "12", "7")]
        [InlineData(" i.012.07 ", 'I', "12", "7")]      // padded components normalize
        [InlineData("E1.0", 'I', "1", "0")]
        [InlineData("Q.0.1", 'O', "0", "1")]
        [InlineData("A2.3", 'O', "2", "3")]
        public void Parses_plc_addresses(string input, char dir, string byteOffset, string bit)
        {
            Assert.True(IoListLogic.TryParsePlcAddress(input, out char d, out string b, out string t));
            Assert.Equal(dir, d);
            Assert.Equal(byteOffset, b);
            Assert.Equal(bit, t);
            if (dir == 'I') Assert.Equal("I" + byteOffset + "." + bit, IoListLogic.BuildPlcAddress(b, t));
        }

        [Theory]
        [InlineData("I.80.3.1")]      // the ASC four-part shape belongs to TryParseAddress
        [InlineData("80.3")]          // no direction
        [InlineData("IW0")]           // a word, not a bit
        [InlineData("I.0")]           // no bit
        [InlineData("X0.1")]          // bad direction
        [InlineData("I.a.1")]
        [InlineData("")]
        [InlineData(null)]
        public void Rejects_non_plc_addresses(string input)
        {
            Assert.False(IoListLogic.TryParsePlcAddress(input, out _, out _, out _));
        }

        [Fact]
        public void Builds_address_and_direction_aware_join_key()
        {
            Assert.Equal("I.80.3.0", IoListLogic.BuildAddress('I', "0080", "3", "0"));
            Assert.Equal("80.3.0", IoListLogic.BuildAddress('\0', "80", "3", "0"));

            // CSV side (DEI_GATEWAY_CELL_TEST_1: node 2, offset 0, bit 0, card 1) meets the IO-list side
            // ("I.2.0.0" on sheet "R1.S1.C1") — verified against 8FVZ_IO0001_R1.3 + Plant_Inputs.csv.
            Assert.True(IoListLogic.TryParseAddress("I.2.0.0", out char d, out string n, out string o, out string b));
            Assert.True(IoListLogic.TryParseCardSheetName("R1.S1.C1", out _, out _, out string card));
            Assert.Equal(IoListLogic.JoinKey(d, n, o, card, b),
                         IoListLogic.JoinKey('I', "2", "0", "1", "0"));

            // The Outputs CSV meets the same row written as "Q.2.0.0".
            Assert.True(IoListLogic.TryParseAddress("Q.2.0.0", out char dq, out _, out _, out _));
            Assert.Equal(IoListLogic.JoinKey('O', "2", "0", "1", "0"),
                         IoListLogic.JoinKey(dq, "2", "0", "1", "0"));

            // Same point, other direction — must NOT collide.
            Assert.NotEqual(IoListLogic.JoinKey('I', "80", "3", "2", "0"),
                            IoListLogic.JoinKey('O', "80", "3", "2", "0"));
        }

        [Fact]
        public void Controller_key_separates_slots_so_card_numbers_do_not_collide()
        {
            // A slot holds a controller; a controller holds several cards. "R1.S2.C1 - PC2" and
            // "R1.S3.C1 - PC3" are DIFFERENT cards that happen to share the number 1, and in the real
            // 8FVZ_IO0002 induction list they share 124 identical addresses. Only the controller tells
            // them apart, so an index must be confined to one before card numbers mean anything.
            Assert.True(IoListLogic.TryParseCardSheetName("R1.S2.C1 - PC2", out string r2, out string s2, out string c2));
            Assert.True(IoListLogic.TryParseCardSheetName("R1.S3.C1 - PC3 ", out string r3, out string s3, out string c3));

            Assert.Equal(c2, c3);                                     // same card NUMBER...
            Assert.NotEqual(IoListLogic.ControllerKey(r2, s2),        // ...different controller
                            IoListLogic.ControllerKey(r3, s3));

            // Without the controller split these two addresses are indistinguishable.
            Assert.Equal(IoListLogic.JoinKey('I', "5", "0", c2, "0"),
                         IoListLogic.JoinKey('I', "5", "0", c3, "0"));
        }

        [Theory]
        [InlineData("1", "1", "R1.S1")]
        [InlineData("01", "02", "R1.S2")]      // normalizes numerically
        [InlineData("1", "12", "R1.S12")]
        public void Controller_key_is_rack_and_slot(string rack, string slot, string expected)
        {
            Assert.Equal(expected, IoListLogic.ControllerKey(rack, slot));
        }

        [Fact]
        public void Node_key_normalizes_and_stays_distinct_per_card()
        {
            Assert.Equal("1#236", IoListLogic.NodeKey("01", "0236"));
            Assert.NotEqual(IoListLogic.NodeKey("1", "236"), IoListLogic.NodeKey("2", "236"));
            Assert.NotEqual(IoListLogic.NodeKey("1", "236"), IoListLogic.NodeKey("1", "237"));
        }

        [Theory]
        // The taxonomy: 1 DI · 2 DO · 3 AI · 4 AO · 5 VDI · 6 VDO · 7 DEI · 8 DEO · 9 VAI · 10 VAO · 11 BUS.
        [InlineData("5", true)]     // VDI
        [InlineData("6", true)]     // VDO
        [InlineData("9", true)]     // AIS — software analog, exported at 0.0.0
        [InlineData("10", true)]    // AOS
        [InlineData("13", true)]    // VAI — the controller's number, not 9
        [InlineData("14", true)]    // VAO
        [InlineData("05", true)]
        [InlineData("1", false)]    // DI
        [InlineData("4", false)]    // AO
        [InlineData("7", false)]    // DEI — a struct, not a virtual
        [InlineData("8", false)]    // DEO
        [InlineData("11", false)]   // BUS
        [InlineData("12", false)]   // skipped by the defines (IN odd, OUT = IN + 1)
        [InlineData("", false)]
        public void Classifies_virtual_types(string type, bool expected)
        {
            Assert.Equal(expected, IoListLogic.IsVirtualType(type));
        }

        [Theory]
        [InlineData("", true, "0")]      // the list's normal case is an EMPTY cell
        [InlineData("0", true, "0")]
        [InlineData(" 1 ", true, "1")]
        [InlineData("NC", false, null)]  // nothing the CSV accepts
        [InlineData("2", false, null)]
        [InlineData(null, true, "0")]
        public void Normal_condition_reads_as_csv_logic(string io, bool readable, string logic)
        {
            Assert.Equal(readable, IoListLogic.TryNormalConditionToLogic(io, out string actual));
            Assert.Equal(logic, actual);
        }

        [Theory]
        [InlineData("0", true)]
        [InlineData("1", true)]
        [InlineData("", false)]
        [InlineData("01", false)]
        [InlineData("NO", false)]
        public void Csv_logic_is_exactly_zero_or_one(string logic, bool valid)
        {
            Assert.Equal(valid, IoListLogic.IsValidLogic(logic));
        }

        [Fact]
        public void Signal_type_table_follows_the_controller_defines()
        {
            // SIGNAL_TYPE_* : 1 DI 2 DO 3 AI 4 AO 5 VI 6 VO 7 DEI 8 DEO 9 AIS 10 AOS 11 BUS (12 skipped) 13 VAI 14 VAO
            Assert.Equal(new[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14 },
                System.Linq.Enumerable.ToArray(System.Linq.Enumerable.Select(CsvSignalType.All, e => e.Number)));
            Assert.Equal("AIS", CsvSignalType.Find("9")!.Acronym);
            Assert.Equal("VAI", CsvSignalType.Find("13")!.Acronym);
            Assert.Null(CsvSignalType.Find("12"));
            Assert.True(CsvSignalType.Find("13")!.IsInput);
            Assert.False(CsvSignalType.Find("14")!.IsInput);
            Assert.Equal(5, CsvSignalType.NumberOf("VI"));
            Assert.Equal(5, CsvSignalType.NumberOf("VDI"));   // the tooling's spelling
            Assert.Equal(0, CsvSignalType.NumberOf("ZZZ"));
            Assert.Equal("1 — DI (digital input)", CsvSignalType.Label(CsvSignalType.Find("1")));
        }

        [Theory]
        [InlineData("11", true)]
        [InlineData("011", true)]
        [InlineData("1", false)]
        [InlineData("5", false)]
        [InlineData("", false)]
        [InlineData(null, false)]
        public void Classifies_bus_diagnostic_types(string type, bool expected)
        {
            Assert.Equal(expected, IoListLogic.IsBusDiagnosticType(type));
        }

        [Theory]
        [InlineData("BUS_CABINETAL00162", "CABINETAL00162")]
        [InlineData("bus_CABINETAL00162", "CABINETAL00162")]
        [InlineData("BUS_CABINETAL00161/ISPC_GROUP_1_NUMLINES", "CABINETAL00161/ISPC_GROUP_1_NUMLINES")]
        [InlineData("CABINETWR00300", "CABINETWR00300")]      // untouched when absent
        [InlineData("BUSBAR_ALARM", "BUSBAR_ALARM")]          // prefix must be exactly "BUS_"
        [InlineData("", "")]
        [InlineData(null, "")]
        public void Strips_bus_prefix_per_part(string input, string expected)
        {
            Assert.Equal(expected, IoListLogic.StripBusPrefix(input));
        }

        [Theory]
        [InlineData("CABINETAL00162", "CABINETAL00162")]
        [InlineData("BUS_CABINETAL00162", "CABINETAL00162")]
        [InlineData("  bus_cabinetal00162  ", "CABINETAL00162")]
        [InlineData("BUSBAR_ALARM", "BUSBAR_ALARM")]      // prefix must be exactly "BUS_"
        [InlineData("", "")]
        [InlineData(null, "")]
        public void Mnemonic_key_normalizes_case_and_bus_prefix(string input, string expected)
        {
            Assert.Equal(expected, IoListLogic.MnemonicKey(input));
        }

        [Theory]
        // identical
        [InlineData("A/B", "A/B", true)]
        // order does not matter
        [InlineData("A/B/C", "C/A/B", true)]
        // BUS_ is ignored on either side
        [InlineData("BUS_CABINETAL00161/ISPC_GROUP_1_NUMLINES", "ISPC_GROUP_1_NUMLINES/CABINETAL00161", true)]
        [InlineData("CABINETAL00162", "BUS_CABINETAL00162", true)]
        // case is ignored
        [InlineData("dj00301", "DJ00301", true)]
        // duplicates and blank parts collapse
        [InlineData("A/A/B", "B/A", true)]
        [InlineData("A//B", "A/B", true)]
        [InlineData(" A / B ", "B/A", true)]
        // genuinely different sets
        [InlineData("A/B", "A", false)]
        [InlineData("A", "A/B", false)]
        [InlineData("LP100301", "LP200301", false)]         // the real off-by-one-bit drift
        [InlineData("A", "B", false)]
        [InlineData("A", "", false)]
        [InlineData("", "A", false)]
        // nothing on either side is not a mismatch
        [InlineData("", "", true)]
        [InlineData(null, null, true)]
        public void Mnemonic_sets_match_ignores_order_case_and_bus_prefix(string a, string b, bool expected)
        {
            Assert.Equal(expected, IoListLogic.MnemonicSetsMatch(a, b));
            Assert.Equal(expected, IoListLogic.MnemonicSetsMatch(b, a));   // symmetric
        }

        [Fact]
        public void Mnemonic_keys_are_distinct_and_ordered_as_written()
        {
            Assert.Equal(new[] { "A", "B" }, IoListLogic.MnemonicKeys("A/B/a/"));
            Assert.Empty(IoListLogic.MnemonicKeys(""));
            Assert.Empty(IoListLogic.MnemonicKeys("///"));
        }

        [Theory]
        // covered: the expected name is there, alone or accompanied
        [InlineData("AUTOM/MANUAL", "AUTOM", true)]
        [InlineData("CABINETAL00161/ISPC_GROUP_1_NUMLINES", "BUS_CABINETAL00161", true)]
        [InlineData("A", "A", true)]
        [InlineData("A/B", "A/B", true)]
        [InlineData("A", "", true)]                          // nothing expected is always covered
        // not covered: the expected name is missing — the two sides disagree about the signal
        [InlineData("LP100301", "LP000301", false)]          // the real drift: CSV name absent from the list
        [InlineData("A/B", "A/C", false)]
        [InlineData("", "A", false)]
        public void Mnemonic_set_covers_separates_accompanied_from_absent(string pulled, string expected, bool covered)
        {
            Assert.Equal(covered, IoListLogic.MnemonicSetCovers(pulled, expected));
        }

        [Fact]
        public void Covered_but_not_matching_is_the_accompanied_case()
        {
            // One IO-list row carrying several mnemonics: a mismatch, but a benign one — worth telling
            // apart from a row whose signal is simply not there.
            Assert.False(IoListLogic.MnemonicSetsMatch("AUTOM/MANUAL", "AUTOM"));
            Assert.True(IoListLogic.MnemonicSetCovers("AUTOM/MANUAL", "AUTOM"));
        }

        [Theory]
        // Agreement is blank — only disagreements draw the eye.
        [InlineData("LP100301", "LP100301", "LP100301", "")]
        [InlineData("A/B", "B/A", "A/B", "")]                     // order and duplicates do not matter
        [InlineData("CABINETAL00162", "BUS_CABINETAL00162", "BUS_CABINETAL00162", "")]
        // Nothing matched by address at all — no verdict is possible.
        [InlineData("LP100301", "", "", "")]
        // The real drift: my name is not what sits at my address.
        [InlineData("LP100301", "LP200301", "LP200301", "0/1")]
        // My name is there but shares the address; the IO-list side (2 mnemonics) sets the total.
        [InlineData("AUTOM", "AUTOM/MANUAL", "AUTOM/MANUAL", "1/2")]
        // IO-list sheet: the row's own cell is the reference, so it sets the total.
        [InlineData("DJ00301/DJ00302/DJ00303/DJ00304", "DJ00301/DJ00302", "DJ00301/DJ00302/DJ00303/DJ00304", "2/4")]
        [InlineData("A/B", "C", "A/B", "0/2")]
        public void Match_ratio_reports_only_disagreements(
            string rowMnemonics, string byAddress, string ioListSide, string expected)
        {
            Assert.Equal(expected, IoListLogic.MatchRatio(rowMnemonics, byAddress, ioListSide));
        }

        [Fact]
        public void Match_ratio_ignores_mnemonics_found_only_by_name()
        {
            // A signal that moved is reported in the by-name column; the ratio judges the address side
            // alone, so its presence elsewhere must not turn "0/1" into agreement.
            Assert.Equal("0/1", IoListLogic.MatchRatio("LP100301", "LP200301", "LP200301"));
        }

        [Theory]
        [InlineData("LP100301", "O.90.0.1", "LP100301@O.90.0.1")]
        [InlineData("CABINETAL00161", "#236", "CABINETAL00161@#236")]
        public void Mnemonic_is_tagged_with_where_it_lives(string mnemonic, string address, string expected)
        {
            Assert.Equal(expected, IoListLogic.AtAddress(mnemonic, address));
        }

        [Theory]
        [InlineData("236", "#236")]
        [InlineData("0236", "#236")]
        public void Module_address_is_marked_as_a_node(string node, string expected)
        {
            // Deliberately unlike a signal address, so "#236" and "I.236.0.0" are never confused.
            Assert.Equal(expected, IoListLogic.ModuleAddress(node));
            Assert.NotEqual(IoListLogic.BuildAddress('I', node, "0", "0"), IoListLogic.ModuleAddress(node));
        }

        [Theory]
        [InlineData("Plant_Inputs.csv", 'I')]
        [InlineData("PLANT_OUTPUTS - Copy.csv", 'O')]
        [InlineData("whatever.csv", '\0')]
        public void Direction_from_name(string name, char expected)
        {
            Assert.Equal(expected, IoListLogic.DirectionFromName(name));
        }

        [Fact]
        public void Finds_columns_in_the_current_io_list_header()
        {
            var headers = new[]
            {
                "Description Module ", "Manufacturer ", "Part No.", "Cod. Fives", "Slot", "ID", "Bit",
                "Normal condition", "Connector ", "Pin No. ",
                "Description language 1 (1° part) (permanent part)",
                "Description language 1 (2° part) (variable part)",
                "Description language 2 (1° part) (permanent part)",
                "Description language 2 (2° part) (variable part)",
                "Functional unit", "Location", "Device", "Type", "Drawing name", "Sheet",
                "T.S. ref.", "Profinet IP", "Profinet name", "Mnemonic", "Alarm filter", "Tag filter",
            };

            Assert.Equal(6, IoListLogic.FindColumn(headers, "bit"));
            Assert.Equal(5, IoListLogic.FindColumn(headers, "id"));
            Assert.Equal(23, IoListLogic.FindColumn(headers, "mnemonic"));
            Assert.Equal(14, IoListLogic.FindColumn(headers, "functional unit"));
            Assert.Equal(17, IoListLogic.FindColumn(headers, "type"));
            Assert.Equal(18, IoListLogic.FindColumn(headers, "drawing name", "drawing"));
            Assert.Equal(19, IoListLogic.FindColumn(headers, "sheet"));
            Assert.Equal(-1, IoListLogic.FindColumn(headers, "rack.slot.card"));   // gone from the format
        }

        [Fact]
        public void Description_fragments_pick_the_language_1_parts_not_description_module()
        {
            var headers = new[]
            {
                "Description Module ", "Manufacturer ", "Bit",
                "Description language 1 (1° part) (permanent part)",
                "Description language 1 (2° part) (variable part)",
                "Description language 2 (1° part) (permanent part)",
            };

            // "Description Module" is column 0 and would win a bare "description" contains-match.
            Assert.Equal(3, IoListLogic.FindColumn(headers, "Description language 1 (1° part", "description english language"));
            Assert.Equal(4, IoListLogic.FindColumn(headers, "Description language 1 (2° part"));
        }

        [Fact]
        public void Finds_both_csv_schemas()
        {
            // Plant_Inputs.csv (comma-delimited)
            var inputs = new[]
            {
                "#", "Signal ID", "Type", "Logic", "RE Filter", "FE Filter", "A1 Filter", "A0 Filter",
                "Events", "Queue", "Flags", "Node", "Offset", "Bit", "Default Virtual Value", "Card", "Description",
            };
            Assert.Equal(1, IoListLogic.FindColumn(inputs, "signal id", "input name", "output name", "name"));
            Assert.Equal(2, IoListLogic.FindColumn(inputs, "type", "input type", "output type"));
            Assert.Equal(11, IoListLogic.FindColumn(inputs, "node"));
            Assert.Equal(13, IoListLogic.FindColumn(inputs, "bit"));
            Assert.Equal(15, IoListLogic.FindColumn(inputs, "card", "board number", "board"));

            // Plant_Outputs.csv (semicolon-delimited, different schema)
            var outputs = new[]
            {
                "#", "Output Name", "Output Type", "NC/NO Logic", "Node", "Offset", "Bit",
                "Default Value (For Virtual I/Os)", "Board number", "Description",
            };
            Assert.Equal(1, IoListLogic.FindColumn(outputs, "signal id", "input name", "output name", "name"));
            Assert.Equal(2, IoListLogic.FindColumn(outputs, "type", "input type", "output type"));
            Assert.Equal(4, IoListLogic.FindColumn(outputs, "node"));
            Assert.Equal(6, IoListLogic.FindColumn(outputs, "bit"));
            Assert.Equal(8, IoListLogic.FindColumn(outputs, "card", "board number", "board"));
        }

        [Fact]
        public void Exact_header_match_beats_contains()
        {
            var headers = new[] { "Script Type", "Type" };
            Assert.Equal(1, IoListLogic.FindColumn(headers, "type"));   // exact "Type", not "Script Type"
        }

        [Fact]
        public void Mnemonic_map_joins_with_slash()
        {
            var map = new Dictionary<string, List<string>>();
            IoListLogic.AddToMap(map, "I.80.3.2.0", "AUTO1_Z01");
            IoListLogic.AddToMap(map, "I.80.3.2.0", "AUTO1_Z01_MIRROR");
            IoListLogic.AddToMap(map, "I.80.3.2.0", "AUTO1_Z01");   // duplicate ignored

            Assert.Equal("AUTO1_Z01/AUTO1_Z01_MIRROR", IoListLogic.Joined(map, "I.80.3.2.0"));
            Assert.Equal("", IoListLogic.Joined(map, "I.9.9.9.9"));
        }
    }
}
