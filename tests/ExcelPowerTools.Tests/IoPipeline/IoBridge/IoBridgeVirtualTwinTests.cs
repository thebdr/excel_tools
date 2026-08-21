using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.IoBridge
{
    /// <summary>
    /// A VIRTUAL record (type 5/6/9/10) carrying an IO point's mnemonic is the point's twin: it
    /// looks present but acquires nothing from the field. The join binds it to the IO row flagged
    /// (label "· virtual twin", verdict "virtual", CsvBinding.Virtual) instead of leaving the row
    /// "missing" — which used to end in an integrate that the apply then reported as "applied
    /// externally" (the FIRE_ALARM_STOP_REQUEST case). The type change that follows is a plain
    /// SetField, resolved to the twin by its mnemonic.
    /// </summary>
    public class IoBridgeVirtualTwinTests
    {
        private const string Header =
            "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n";

        // An outputs file with nothing on card 1, so the CSV-only bucket reflects the inputs alone.
        private const string OutputsText =
            "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
            "@;DOOR_CMD_01;2;0;235;10;1;0;9;DOOR CMD ON ANOTHER CARD\r\n";

        private static CardSheetGrid Card(params string[][] rows)
        {
            var all = new List<string[]> { new[] { "ID", "Bit", "Normal condition", "Mnemonic" } };
            all.AddRange(rows);
            return new CardSheetGrid { SheetName = "R1.S1.C1", FirstRowNumber = 1, AllRows = all };
        }

        private static BridgeWorkspaceModel Build(string inputs, CardSheetGrid card, IoBridgeJoinOptions options = null) =>
            IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(inputs)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)),
                options);

        [Fact]
        public void A_virtual_record_with_the_io_points_name_binds_as_its_twin_flagged()
        {
            string inputs = Header +
                "@,FIRE_ALARM_STOP_REQUEST,5,0,0,0,0,0,0,0,0,0,0,0,0,1,FIRE ALARM STOP REQUEST\r\n";
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.90.0.2", "1", "FIRE_ALARM_STOP_REQUEST" }))
                .Sheets.Single();

            BridgeRow row = sheet.SpineRows.Single();
            Assert.NotNull(row.Csv);
            Assert.True(row.Csv!.Virtual);
            Assert.False(row.Csv.IsMulti);
            Assert.Equal("Signal · virtual twin", row.RowLabel);
            Assert.Equal("virtual", row.MatchVerdict);                      // never a clean (green) match
            Assert.Equal("FIRE_ALARM_STOP_REQUEST", row.CsvValues[CsvCanonical.Name]);
            Assert.Equal("5", row.CsvValues[CsvCanonical.Type]);
            Assert.Empty(sheet.CsvOnlyRows);                                // imported into the row, not left over
        }

        [Fact]
        public void A_real_record_at_the_address_wins_and_the_virtual_stays_csv_only()
        {
            string inputs = Header +
                "@,FIRE_ALARM_STOP_REQUEST,1,0,0,0,0,0,0,0,0,90,0,2,0,1,REAL ONE\r\n" +
                "@,FIRE_ALARM_STOP_REQUEST,5,0,0,0,0,0,0,0,0,0,0,0,0,1,VIRTUAL COPY\r\n";
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.90.0.2", "1", "FIRE_ALARM_STOP_REQUEST" }))
                .Sheets.Single();

            BridgeRow row = sheet.SpineRows.Single();
            Assert.False(row.Csv!.Virtual);
            Assert.Equal("Signal", row.RowLabel);
            Assert.Equal("", row.MatchVerdict);
            BridgeRow leftover = Assert.Single(sheet.CsvOnlyRows);
            Assert.Equal("CSV-only · virtual", leftover.RowLabel);
        }

        [Fact]
        public void A_virtual_with_another_name_is_no_twin_and_the_row_still_reads_missing()
        {
            string inputs = Header +
                "@,SOMETHING_ELSE,5,0,0,0,0,0,0,0,0,0,0,0,0,1,UNRELATED VIRTUAL\r\n";
            BridgeSheetModel sheet = Build(inputs, Card(new[] { "", "I.90.0.2", "1", "FIRE_ALARM_STOP_REQUEST" }))
                .Sheets.Single();

            Assert.Null(sheet.SpineRows.Single().Csv);
            Assert.Equal("missing", sheet.SpineRows.Single().MatchVerdict);
            Assert.Equal("CSV-only · virtual", Assert.Single(sheet.CsvOnlyRows).RowLabel);
        }

        [Fact]
        public void Opted_in_virtuals_read_clean_while_every_other_virtual_twin_is_flagged()
        {
            // The compare-virtuals regex declares which virtuals are EXPECTED (the CLASSDIAGAL
            // pattern): those join by name as ordinary matches. Every other virtual carrying an IO
            // point's name is a twin — flagged. Neither is ever claimed twice.
            string inputs = Header +
                "@,CLASSDIAGAL00620030130,5,0,0,0,0,0,0,0,0,0,0,0,0,1,ASI FAULT\r\n" +
                "@,FIRE_ALARM_STOP_REQUEST,5,0,0,0,0,0,0,0,0,0,0,0,0,1,FIRE ALARM\r\n";
            BridgeSheetModel sheet = Build(inputs,
                Card(new[] { "", "I.90.1.2", "1", "CLASSDIAGAL00620030130" },
                     new[] { "", "I.90.0.2", "1", "FIRE_ALARM_STOP_REQUEST" }),
                new IoBridgeJoinOptions { IncludeVirtualsMatching = "^CLASSDIAGAL" }).Sheets.Single();

            BridgeRow classdiag = sheet.SpineRows[0];
            BridgeRow fire = sheet.SpineRows[1];
            Assert.False(classdiag.Csv!.Virtual);
            Assert.Equal("", classdiag.MatchVerdict);
            Assert.Equal("Signal", classdiag.RowLabel);
            Assert.True(fire.Csv!.Virtual);
            Assert.Equal("virtual", fire.MatchVerdict);
            Assert.Single(classdiag.Csv.RecordIndexes);
            Assert.Single(fire.Csv.RecordIndexes);
            Assert.Empty(sheet.CsvOnlyRows);
        }

        [Fact]
        public void Changing_the_twins_type_is_a_set_field_resolved_by_mnemonic()
        {
            string inputs = Header +
                "@,FIRE_ALARM_STOP_REQUEST,5,0,0,0,0,0,0,0,0,0,0,0,0,1,FIRE ALARM STOP REQUEST\r\n";
            CsvDocument doc = CsvFile.Load(Encoding.UTF8.GetBytes(inputs));
            CsvColumnsMap map = CsvColumnsMap.Locate(doc);

            // The action carries the IO point's identity (its address) — the twin sits at 0.0.0.
            var action = new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = BridgeTarget.InputsCsv,
                Signal = new SignalRef { CardKey = "R1.S1.C1", Address = "I.90.0.2", Mnemonic = "FIRE_ALARM_STOP_REQUEST" },
                FieldHeader = "type",
                OldValue = "5",
                NewValue = "1",
            };

            CsvPatchResult result = CsvPatchApplier.Apply(doc, new[] { action }, map);

            CsvActionOutcome outcome = Assert.Single(result.Outcomes);
            Assert.Equal(BridgeActionStatus.Applied, outcome.Outcome);
            Assert.True(result.AnyMutation);
            Assert.Equal("1", doc.Records[1].Fields[map.Type]);
        }
    }
}
