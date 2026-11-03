using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.IoBridge
{
    /// <summary>
    /// "Found at a different address": an IO row whose name the CSV carries at ANOTHER address of the
    /// same card (FVT LOWER SORTER: MODULATOR_DEI_00061 at I.35.1.0 in the list, I.35.4.0 in the
    /// CSV) is paired by name and flagged "moved" instead of reading "missing" beside a CSV-only
    /// leftover. The address pair of row actions then aligns either side.
    /// </summary>
    public class IoBridgeMovedAddressTests
    {
        private const string Header =
            "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n";

        private const string OutputsText =
            "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
            "@;DOOR_CMD_01;2;0;235;10;1;0;9;DOOR CMD ON ANOTHER CARD\r\n";

        private static CardSheetGrid Card(params string[][] rows)
        {
            var all = new List<string[]> { new[] { "ID", "Bit", "Normal condition", "Mnemonic" } };
            all.AddRange(rows);
            return new CardSheetGrid { SheetName = "R1.S1.C1", FirstRowNumber = 1, AllRows = all };
        }

        private static BridgeSheetModel Build(string inputs, CardSheetGrid card) =>
            IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(inputs)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)))
                .Sheets.Single();

        private static string Rec(string name, int node, int off, int bit, string card = "1", string type = "1") =>
            "@," + name + "," + type + ",0,0,0,0,0,0,0,0," + node + "," + off + "," + bit + ",0," + card + ",DESC\r\n";

        [Fact]
        public void A_name_at_another_address_pairs_flagged_and_leaves_no_csv_only_row()
        {
            string inputs = Header + Rec("MODULATOR_DEI_00061", 35, 4, 0, type: "7");
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.35.1.0", "1", "MODULATOR_DEI_00061" }));

            BridgeRow row = sheet.SpineRows.Single();
            Assert.NotNull(row.Csv);
            Assert.True(row.Csv!.Moved);
            Assert.False(row.Csv.IsMulti);
            Assert.Equal("Signal · moved", row.RowLabel);
            Assert.Equal("moved", row.MatchVerdict);                  // never the clean match
            Assert.Equal("4", row.CsvValues[CsvCanonical.Offset]);
            Assert.False(row.Csv.IoAddressOccupiedInCsv);
            Assert.False(row.Csv.CsvAddressOccupiedInIo);
            Assert.Empty(sheet.CsvOnlyRows);
        }

        [Fact]
        public void The_same_address_still_wins_and_reads_clean()
        {
            string inputs = Header + Rec("EBOX_TEST_1", 57, 2, 0);
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.57.2.0", "1", "EBOX_TEST_1" }));

            BridgeRow row = sheet.SpineRows.Single();
            Assert.False(row.Csv!.Moved);
            Assert.Equal("Signal", row.RowLabel);
            Assert.Equal("", row.MatchVerdict);
        }

        [Fact]
        public void A_shifted_record_is_not_stolen_by_the_unrelated_row_now_at_its_old_address()
        {
            // List: EBOX_TEST_1 at 57.1.0, MODULATOR_DEI_00331 at 57.2.0.
            // CSV : EBOX_TEST_1 at 57.2.0, MODULATOR_DEI_00331 at 57.4.0 (the FVT shift).
            string inputs = Header + Rec("EBOX_TEST_1", 57, 2, 0) + Rec("MODULATOR_DEI_00331", 57, 4, 0, type: "7");
            BridgeSheetModel sheet = Build(inputs, Card(
                new[] { "", "I.57.1.0", "1", "EBOX_TEST_1" },
                new[] { "", "I.57.2.0", "1", "MODULATOR_DEI_00331" }));

            BridgeRow ebox = sheet.SpineRows[0];
            BridgeRow dei = sheet.SpineRows[1];
            Assert.Equal("EBOX_TEST_1", ebox.CsvValues[CsvCanonical.Name]);
            Assert.Equal("MODULATOR_DEI_00331", dei.CsvValues[CsvCanonical.Name]);
            Assert.All(new[] { ebox, dei }, r => Assert.Equal("moved", r.MatchVerdict));
            Assert.Empty(sheet.CsvOnlyRows);

            // EBOX_TEST_1: its IO address 57.1.0 is free in the CSV (CSV side may move), but the CSV
            // address 57.2.0 is the DEI row's address (IO side may not take it).
            Assert.False(ebox.Csv!.IoAddressOccupiedInCsv);
            Assert.True(ebox.Csv.CsvAddressOccupiedInIo);
            // MODULATOR_DEI_00331: its IO address 57.2.0 holds EBOX_TEST_1's record (CSV side blocked);
            // its CSV address 57.4.0 is free in the list (IO side may take it).
            Assert.True(dei.Csv!.IoAddressOccupiedInCsv);
            Assert.False(dei.Csv.CsvAddressOccupiedInIo);
        }

        [Fact]
        public void A_swapped_pair_pairs_by_name_and_blocks_both_alignments()
        {
            string inputs = Header + Rec("PAD_JOG_DX_1", 32, 1, 4) + Rec("PAD_JOG_SX_1", 32, 1, 5);
            BridgeSheetModel sheet = Build(inputs, Card(
                new[] { "", "I.32.1.4", "1", "PAD_JOG_SX_1" },
                new[] { "", "I.32.1.5", "1", "PAD_JOG_DX_1" }));

            Assert.All(sheet.SpineRows, r =>
            {
                Assert.True(r.Csv!.Moved);
                Assert.True(r.Csv.IoAddressOccupiedInCsv);
                Assert.True(r.Csv.CsvAddressOccupiedInIo);
            });
            Assert.Equal("PAD_JOG_SX_1", sheet.SpineRows[0].CsvValues[CsvCanonical.Name]);
            Assert.Empty(sheet.CsvOnlyRows);
        }

        [Fact]
        public void A_name_carried_by_two_records_is_ambiguous_and_stays_missing()
        {
            string inputs = Header + Rec("TWICE", 40, 1, 0) + Rec("TWICE", 40, 2, 0);
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.40.9.0", "1", "TWICE" }));

            Assert.Null(sheet.SpineRows.Single().Csv);
            Assert.Equal("missing", sheet.SpineRows.Single().MatchVerdict);
            Assert.Equal(2, sheet.CsvOnlyRows.Count);
        }

        [Fact]
        public void A_record_on_another_card_pairs_flagged_off_card_with_only_the_csv_side_alignable()
        {
            string inputs = Header + Rec("ELSEWHERE", 40, 1, 0, card: "2");
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.40.9.0", "1", "ELSEWHERE" }));

            BridgeRow row = sheet.SpineRows.Single();
            Assert.True(row.Csv!.Moved);
            Assert.True(row.Csv.OtherCard);
            Assert.Equal("2", row.Csv.CsvCard);
            Assert.Equal("Signal · moved · card 2", row.RowLabel);
            Assert.Equal("moved", row.MatchVerdict);
        }

        [Fact]
        public void A_record_on_another_selected_card_stays_with_its_own_row()
        {
            // Card 2 has its own row at the record's address: it must win, whatever the sheet order.
            string inputs = Header + Rec("SHARED", 40, 1, 0, card: "2");
            var card1 = Card(new[] { "", "I.40.9.0", "1", "SHARED" });
            var card2 = new CardSheetGrid
            {
                SheetName = "R1.S1.C2", FirstRowNumber = 1,
                AllRows = new List<string[]> { new[] { "ID", "Bit", "Normal condition", "Mnemonic" }, new[] { "", "I.40.1.0", "1", "SHARED" } },
            };
            BridgeWorkspaceModel model = IoBridgeJoin.Build(new[] { card1, card2 },
                CsvFile.Load(Encoding.UTF8.GetBytes(inputs)), CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)));

            Assert.Equal("", model.Sheets[1].SpineRows.Single().MatchVerdict);
            Assert.Null(model.Sheets[0].SpineRows.Single().Csv);
            Assert.Equal("missing", model.Sheets[0].SpineRows.Single().MatchVerdict);
        }

        [Fact]
        public void Aligning_the_csv_to_the_io_row_is_a_set_field_resolved_by_mnemonic()
        {
            string inputs = Header + Rec("MODULATOR_DEI_00061", 35, 4, 0, type: "7");
            CsvDocument doc = CsvFile.Load(Encoding.UTF8.GetBytes(inputs));
            CsvColumnsMap map = CsvColumnsMap.Locate(doc, PlantCsvKind.Inputs);

            // The action carries the IO row's identity (I.35.1.0); the record still sits at 35.4.0.
            var action = new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = BridgeTarget.InputsCsv,
                Signal = new SignalRef { CardKey = "R1.S1.C1", Address = "I.35.1.0", Mnemonic = "MODULATOR_DEI_00061" },
                FieldHeader = "offset",
                OldValue = "4",
                NewValue = "1",
            };

            CsvPatchResult result = CsvPatchApplier.Apply(doc, new[] { action }, map);

            Assert.Equal(BridgeActionStatus.Applied, Assert.Single(result.Outcomes).Outcome);
            Assert.Equal("1", doc.Records[1].Fields[map.Offset]);
        }
    }
}
