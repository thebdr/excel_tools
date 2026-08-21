using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.IoBridge
{
    /// <summary>
    /// The exclude regex filters MNEMONICS, not rows: a generator key written in plain mnemonic
    /// syntax beside a real mnemonic ("BUS_CABINETAL00161/SAFETY_NUMDOORS_16") drops out of the
    /// comparison while the mnemonic beside it still joins. Only a cell whose every part matches is
    /// excluded whole.
    /// </summary>
    public class IoBridgeExcludeMnemonicsTests
    {
        private const string InputsText =
            "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n" +
            "@,LP100301,1,0,0,0,0,0,0,0,0,90,0,1,0,1,LAMP 301\r\n" +
            "@,CABINETAL00161,11,0,0,0,0,0,0,0,0,236,0,0,0,1,FIELD BUS FAILURE\r\n";

        private const string OutputsText =
            "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n";

        private static CardSheetGrid Card(params string[][] rows)
        {
            var all = new List<string[]> { new[] { "ID", "Bit", "Normal condition", "Mnemonic" } };
            all.AddRange(rows);
            return new CardSheetGrid { SheetName = "R1.S1.C1", FirstRowNumber = 1, AllRows = all };
        }

        private static BridgeSheetModel Build(CardSheetGrid card, string excludeRegex) =>
            IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)),
                new IoBridgeJoinOptions { ExcludeIoMnemonicsMatching = excludeRegex }).Sheets.Single();

        [Fact]
        public void A_generator_key_beside_a_module_mnemonic_drops_out_and_the_module_still_joins()
        {
            BridgeRow row = Build(Card(new[] { "236", "", "", "BUS_CABINETAL00161/SAFETY_NUMDOORS_16" }), "^SAFETY_NUMDOORS")
                .SpineRows.Single();

            Assert.Equal(BridgeRowKind.Module, row.Kind);
            Assert.Equal("CABINETAL00161", row.Signal.Mnemonic);          // the effective mnemonic, BUS_ stripped
            Assert.Equal("SAFETY_NUMDOORS_16", row.ExcludedMnemonics);
            Assert.Equal("Module · filtered", row.RowLabel);
            Assert.NotNull(row.Csv);                                       // joined by node to the type-11 record
            Assert.Equal("", row.MatchVerdict);                            // the key no longer counts as a missing name
        }

        [Fact]
        public void Without_the_regex_the_key_reads_as_a_missing_name()
        {
            BridgeRow row = Build(Card(new[] { "236", "", "", "BUS_CABINETAL00161/SAFETY_NUMDOORS_16" }), "")
                .SpineRows.Single();

            Assert.Equal("CABINETAL00161/SAFETY_NUMDOORS_16", row.Signal.Mnemonic);
            Assert.Equal("", row.ExcludedMnemonics);
            Assert.Equal("1/2", row.MatchVerdict);                         // the noise the regex exists to remove
        }

        [Fact]
        public void A_key_beside_a_signal_mnemonic_is_filtered_the_same_way()
        {
            BridgeRow row = Build(Card(new[] { "", "I.90.0.1", "1", "LP100301/<SPECIAL_KEY>" }), "^<.*>$")
                .SpineRows.Single();

            Assert.Equal("LP100301", row.Signal.Mnemonic);
            Assert.Equal("<SPECIAL_KEY>", row.ExcludedMnemonics);
            Assert.Equal("Signal · filtered", row.RowLabel);
            Assert.Equal("", row.MatchVerdict);
            Assert.Equal("R1.S1.C1|I.90.0.1", row.Signal.IdentityKey());
        }

        [Fact]
        public void A_cell_whose_every_part_matches_is_excluded_whole()
        {
            BridgeSheetModel sheet = Build(Card(new[] { "", "I.90.0.1", "1", "SAFETY_NUMDOORS_16/SAFETY_NUMDOORS_17" }), "^SAFETY_");
            BridgeRow row = sheet.SpineRows.Single();

            Assert.Equal("Signal · excluded", row.RowLabel);
            Assert.Null(row.Csv);
            Assert.Equal("", row.MatchVerdict);
            Assert.Equal("SAFETY_NUMDOORS_16/SAFETY_NUMDOORS_17", row.Signal.Mnemonic);   // the name stays for display
            Assert.Equal("SAFETY_NUMDOORS_16/SAFETY_NUMDOORS_17", row.ExcludedMnemonics);
            Assert.Contains(sheet.CsvOnlyRows, r => r.Signal.Mnemonic == "LP100301");   // its record stays unclaimed
        }

        [Theory]
        [InlineData("BUS_CABINETAL00161/SAFETY_NUMDOORS_16", "^SAFETY_", "BUS_CABINETAL00161", "SAFETY_NUMDOORS_16", false)]
        [InlineData("SAFETY_A / CAB1 / SAFETY_B", "^SAFETY_", "CAB1", "SAFETY_A/SAFETY_B", false)]
        [InlineData("SAFETY_A/SAFETY_B", "^SAFETY_", "SAFETY_A/SAFETY_B", "SAFETY_A/SAFETY_B", true)]   // whole-cell exclusion keeps the name
        [InlineData("CAB1", "^SAFETY_", "CAB1", "", false)]
        [InlineData("", "^SAFETY_", "", "", false)]
        [InlineData("/ /", "^SAFETY_", "", "", false)]
        public void Filtering_splits_the_cell_on_slashes_and_keeps_each_groups_order(
            string cell, string regex, string kept, string excluded, bool allExcluded)
        {
            IoBridgeJoin.FilterMnemonics(cell, new Regex(regex), out string k, out string x, out bool all);
            Assert.Equal(kept, k);
            Assert.Equal(excluded, x);
            Assert.Equal(allExcluded, all);
        }
    }
}
