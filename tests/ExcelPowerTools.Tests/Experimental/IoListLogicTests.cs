using System.Collections.Generic;
using ExcelPowerTools.Core.Experimental;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental
{
    public class IoListLogicTests
    {
        [Theory]
        [InlineData("80.3.2", "80", "3", "2")]
        [InlineData(" 8.0.2 ", "8", "0", "2")]
        [InlineData("08.03.02", "8", "3", "2")]     // numeric normalization
        [InlineData("1.2.3.4", "1.2", "3", "4")]    // dotted node tolerated (card = last)
        public void Parses_rack_slot_card(string input, string node, string offset, string card)
        {
            Assert.True(IoListLogic.TryParseRackSlotCard(input, out string n, out string o, out string c));
            Assert.Equal(node, n);
            Assert.Equal(offset, o);
            Assert.Equal(card, c);
        }

        [Theory]
        [InlineData("")]
        [InlineData("80.3")]
        [InlineData(null)]
        public void Rejects_incomplete_rack_slot_card(string input)
        {
            Assert.False(IoListLogic.TryParseRackSlotCard(input, out _, out _, out _));
        }

        [Fact]
        public void Builds_address_and_join_key()
        {
            Assert.Equal("I.80.3.0", IoListLogic.BuildAddress('I', "80", "3", "0"));
            Assert.Equal("O.8.0.15", IoListLogic.BuildAddress('O', "08", "00", "15"));
            Assert.Equal("80.3.0", IoListLogic.BuildAddress('\0', "80", "3", "0"));
            Assert.Equal("80.3.2.0", IoListLogic.JoinKey("80", "03", "2", "00"));
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
                "Mnemonic", "ID", "Bit", "Description English language",
                "Functional unit", "Location", "Device", "Type",
                "Drawing", "Sheet", "PC Board\nRack.Slot.Card",
            };

            Assert.Equal(0, IoListLogic.FindColumn(headers, "mnemonic"));
            Assert.Equal(1, IoListLogic.FindColumn(headers, "id"));
            Assert.Equal(3, IoListLogic.FindColumn(headers, "description english", "description"));
            Assert.Equal(4, IoListLogic.FindColumn(headers, "functional unit"));
            Assert.Equal(10, IoListLogic.FindColumn(headers, "rack.slot.card", "pc board"));
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
            IoListLogic.AddToMap(map, "80.3.2.0", "AUTO1_Z01");
            IoListLogic.AddToMap(map, "80.3.2.0", "AUTO1_Z01_MIRROR");
            IoListLogic.AddToMap(map, "80.3.2.0", "AUTO1_Z01");   // duplicate ignored

            Assert.Equal("AUTO1_Z01/AUTO1_Z01_MIRROR", IoListLogic.Joined(map, "80.3.2.0"));
            Assert.Equal("", IoListLogic.Joined(map, "9.9.9.9"));
        }
    }
}
