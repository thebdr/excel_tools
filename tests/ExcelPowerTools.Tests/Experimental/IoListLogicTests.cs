using System.Collections.Generic;
using ExcelPowerTools.Core.Experimental;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental
{
    public class IoListLogicTests
    {
        [Theory]
        [InlineData("R1.S1.C2", "1", "1", "2")]     // the real legacy format
        [InlineData("R1.S12.C1", "1", "12", "1")]
        [InlineData("80.3.2", "80", "3", "2")]      // bare numeric variant
        [InlineData(" r2.s3.c04 ", "2", "3", "4")]  // case + zero padding
        public void Parses_rack_slot_card(string input, string rack, string slot, string card)
        {
            Assert.True(IoListLogic.TryParseRackSlotCard(input, out string r, out string s, out string c));
            Assert.Equal(rack, r);
            Assert.Equal(slot, s);
            Assert.Equal(card, c);
        }

        [Theory]
        [InlineData("")]
        [InlineData("R1.S1")]
        [InlineData(null)]
        public void Rejects_incomplete_rack_slot_card(string input)
        {
            Assert.False(IoListLogic.TryParseRackSlotCard(input, out _, out _, out _));
        }

        [Theory]
        [InlineData("I.0080.3.0", 'I', "80", "3", "0")]   // as stored in the legacy list (padded node)
        [InlineData("O.0080.3.7", 'O', "80", "3", "7")]
        [InlineData("i.5.0.15", 'I', "5", "0", "15")]
        public void Parses_full_addresses(string input, char dir, string node, string offset, string bit)
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
        [InlineData("")]
        public void Rejects_non_addresses(string input)
        {
            Assert.False(IoListLogic.TryParseAddress(input, out _, out _, out _, out _));
        }

        [Fact]
        public void Builds_address_and_direction_aware_join_key()
        {
            Assert.Equal("I.80.3.0", IoListLogic.BuildAddress('I', "0080", "3", "0"));
            Assert.Equal("80.3.0", IoListLogic.BuildAddress('\0', "80", "3", "0"));

            // CSV side (node 80, offset 3, board 2, bit 0) meets IO-list side (I.0080.3.0 + R1.S1.C2).
            Assert.True(IoListLogic.TryParseAddress("I.0080.3.0", out char d, out string n, out string o, out string b));
            Assert.True(IoListLogic.TryParseRackSlotCard("R1.S1.C2", out _, out _, out string card));
            Assert.Equal(IoListLogic.JoinKey(d, n, o, card, b),
                         IoListLogic.JoinKey('I', "80", "3", "2", "0"));

            // Same point, other direction — must NOT collide.
            Assert.NotEqual(IoListLogic.JoinKey('I', "80", "3", "2", "0"),
                            IoListLogic.JoinKey('O', "80", "3", "2", "0"));
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
        public void Finds_columns_by_normalized_headers()
        {
            var headers = new[]
            {
                "Modello Modulo", "Targa Modulo", "PC Board Rack.Slot.Card", "ID", "Bit",
                "Normal condition", "Slot", "Pin No.", "Description English language",
                "Description language 2", "Functional unit", "Location", "Device", "Type",
                "Drawing", "Sheet", "Mnemonic",
            };

            Assert.Equal(16, IoListLogic.FindColumn(headers, "mnemonic"));
            Assert.Equal(3, IoListLogic.FindColumn(headers, "id"));
            Assert.Equal(8, IoListLogic.FindColumn(headers, "description english language", "description english", "description"));
            Assert.Equal(10, IoListLogic.FindColumn(headers, "functional unit"));
            Assert.Equal(2, IoListLogic.FindColumn(headers, "rack.slot.card", "pc board"));
            Assert.Equal(-1, IoListLogic.FindColumn(headers, "does not exist"));
        }

        [Fact]
        public void Exact_header_match_beats_contains()
        {
            var headers = new[] { "Input Type", "Type" };
            Assert.Equal(1, IoListLogic.FindColumn(headers, "type"));   // exact "Type", not "Input Type"
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
