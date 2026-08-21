using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.IoBridge
{
    public class IoBridgeJoinTests
    {
        private const string InputsText =
            "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n" +
            "@,LP100301,1,0,0,0,0,0,0,0,0,90,0,1,0,1,LAMP 301\r\n" +
            "@,CABINETAL00161,11,0,0,0,0,0,0,0,0,236,0,0,0,1,FIELD BUS FAILURE\r\n" +
            "@,VIRT_SIG,5,0,0,0,0,0,0,0,0,0,0,0,0,1,VIRTUAL ONE\r\n" +
            "#,DISABLED_SIG,1,1,0,0,0,0,0,0,0,11,1,1,0,1,DISABLED ONE\r\n" +
            "@,ORPHAN_SIG,1,0,0,0,0,0,0,0,0,50,2,3,0,1,NOT IN LIST\r\n" +
            "@,OTHER_CARD,1,0,0,0,0,0,0,0,0,7,0,0,0,9,DIFFERENT CARD\r\n" +
            "!,COMMON_CSV\\Common_Inputs.csv,,,,,,,,,,,,,,,#\r\n";

        private const string OutputsText =
            "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
            "@;DOOR_CMD_01;2;0;235;10;1;0;1;DOOR CMD\r\n";

        private static CardSheetGrid Card()
        {
            return new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                FirstRowNumber = 1,
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "236", "", "", "BUS_CABINETAL00161" },          // module row
                    new[] { "", "I.90.0.1", "1", "LP100301" },
                    new[] { "", "Q.235.10.1", "0", "DOOR_CMD_01" },
                    new[] { "", "I.90.0.2", "1", "NEW_SIGNAL" },            // not in any CSV
                    new[] { "", "I.11.1.1", "1", "OTHERNAME" },             // joins the disabled row, name differs
                    new[] { "", "", "", "" },                               // decoration — skipped
                },
            };
        }

        private static BridgeWorkspaceModel Build() =>
            IoBridgeJoin.Build(
                new[] { Card() },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)));

        [Theory]
        // The user-confirmed reference example, piece by piece.
        [InlineData("CIRCUIT BREAKER TRIPPED", "POWER SUPPLY COOLING FAN", "+MS1.CC1", "-F22001", "8FVZ_Q0001",
            "CIRCUIT BREAKER TRIPPED POWER SUPPLY COOLING FAN +MS1.CC1-F22001 [Q0001]")]
        // Missing variable part.
        [InlineData("OVERTEMPERATURE ALARM", "", "+MS1.CC1", "-S70301", "8FVZ_Q0001",
            "OVERTEMPERATURE ALARM +MS1.CC1-S70301 [Q0001]")]
        // No drawing.
        [InlineData("BLADE DETECTION ALARM", "", "+PC1.CC1", "-F30201", "",
            "BLADE DETECTION ALARM +PC1.CC1-F30201")]
        // No location/device.
        [InlineData("PROFINET CARD", "", "", "", "8xxx_Qxxxx",
            "PROFINET CARD [Qxxxx]")]
        // Everything empty stays empty.
        [InlineData("", "", "", "", "", "")]
        // Untrimmed inputs and a drawing without a prefix.
        [InlineData(" A ", " B ", " +L1 ", " -D1 ", " P1006 ",
            "A B +L1-D1 [P1006]")]
        public void Csv_description_composes_as_the_export_tool_writes_it(
            string desc1a, string desc1b, string location, string device, string drawing, string expected)
        {
            Assert.Equal(expected, CsvDescription.Compose(desc1a, desc1b, location, device, drawing));
        }

        [Theory]
        [InlineData("8FVZ_Q0001", "Q0001")]
        [InlineData("8xxx_C0002", "C0002")]
        [InlineData("A_B_C123", "C123")]       // the LAST underscore splits
        [InlineData("P1006", "P1006")]         // no prefix — used whole
        [InlineData("TRAILING_", "TRAILING_")] // nothing after the underscore — used whole
        [InlineData("", "")]
        public void Drawing_code_is_the_part_after_the_last_underscore(string drawingName, string expected)
        {
            Assert.Equal(expected, CsvDescription.ShortDrawing(drawingName));
        }

        [Fact]
        public void Header_row_is_found_by_bit_plus_mnemonic()
        {
            Assert.Equal(0, IoBridgeJoin.FindHeaderRow(Card().AllRows));
            Assert.Equal(-1, IoBridgeJoin.FindHeaderRow(new List<string[]> { new[] { "just", "text" } }));
        }

        [Fact]
        public void Spine_preserves_io_list_order_and_skips_decoration_rows()
        {
            BridgeSheetModel sheet = Build().Sheets.Single();

            Assert.Equal("R1.S1.C1", sheet.CardKey);
            Assert.Equal("1", sheet.Card);
            Assert.Equal(1, sheet.BitColumn);
            Assert.Equal(1, sheet.IoHeaderSourceRow);

            Assert.Equal(
                new[] { "#236", "I.90.0.1", "O.235.10.1", "I.90.0.2", "I.11.1.1" },
                sheet.SpineRows.Select(r => r.Signal.Address));
            Assert.Equal(new[] { 2, 3, 4, 5, 6 }, sheet.SpineRows.Select(r => r.IoSourceRow));
        }

        [Fact]
        public void Module_rows_join_by_node_and_strip_the_bus_prefix()
        {
            BridgeRow module = Build().Sheets.Single().SpineRows[0];

            Assert.Equal(BridgeRowKind.Module, module.Kind);
            Assert.Equal("CABINETAL00161", module.Signal.Mnemonic);
            Assert.NotNull(module.Csv);
            Assert.Equal(BridgeTarget.InputsCsv, module.Csv!.File);
            Assert.Equal("CABINETAL00161", module.CsvValues[CsvCanonical.Name]);
            Assert.Equal("", module.MatchVerdict);   // agree once BUS_ is ignored
        }

        [Fact]
        public void Directions_route_to_their_own_csv()
        {
            BridgeSheetModel sheet = Build().Sheets.Single();

            BridgeRow input = sheet.SpineRows[1];
            Assert.Equal(BridgeTarget.InputsCsv, input.Csv!.File);
            Assert.Equal("LP100301", input.CsvValues[CsvCanonical.Name]);
            Assert.Equal("", input.MatchVerdict);

            BridgeRow output = sheet.SpineRows[2];   // written as Q on the card sheet
            Assert.Equal(BridgeTarget.OutputsCsv, output.Csv!.File);
            Assert.Equal("DOOR_CMD_01", output.CsvValues[CsvCanonical.Name]);
            Assert.Equal("0", output.CsvValues[CsvCanonical.Logic]);   // "NC/NO Logic" resolved
            Assert.Equal("1", output.CsvValues[CsvCanonical.Card]);    // "Board number" resolved
        }

        [Fact]
        public void A_named_module_without_its_type11_record_reads_missing_not_matching()
        {
            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "999", "", "", "BUS_CABINETAL09999" },   // no CSV record at node 999
                },
            };
            BridgeRow module = IoBridgeJoin.Build(new[] { card },
                    CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                    CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)))
                .Sheets.Single().SpineRows.Single();

            Assert.Equal(BridgeRowKind.Module, module.Kind);
            Assert.Null(module.Csv);
            Assert.Equal("missing", module.MatchVerdict);   // a real gap — must not fade green
        }

        [Fact]
        public void A_spine_signal_without_a_csv_row_reads_missing()
        {
            BridgeRow missing = Build().Sheets.Single().SpineRows[3];
            Assert.Null(missing.Csv);
            Assert.Equal("missing", missing.MatchVerdict);
            Assert.Empty(missing.CsvValues);
        }

        [Fact]
        public void A_disabled_csv_row_still_joins_and_is_labelled()
        {
            BridgeRow row = Build().Sheets.Single().SpineRows[4];

            Assert.True(row.Csv!.Disabled);
            Assert.Contains("disabled", row.RowLabel);
            Assert.Equal("0/1", row.MatchVerdict);   // DISABLED_SIG vs OTHERNAME
        }

        [Fact]
        public void Csv_only_rows_collect_on_their_card_and_other_cards_drop()
        {
            BridgeSheetModel sheet = Build().Sheets.Single();

            Assert.Equal(2, sheet.CsvOnlyRows.Count);
            BridgeRow virt = sheet.CsvOnlyRows.Single(r => r.Signal.Mnemonic == "VIRT_SIG");
            Assert.Contains("virtual", virt.RowLabel);
            Assert.Equal("", virt.Signal.Address);   // virtual addresses are stale — identity is the name

            BridgeRow orphan = sheet.CsvOnlyRows.Single(r => r.Signal.Mnemonic == "ORPHAN_SIG");
            Assert.Equal("I.50.2.3", orphan.Signal.Address);
            Assert.Equal(BridgeRowKind.CsvOnly, orphan.Kind);

            // OTHER_CARD (board 9) belongs to no selected sheet; the '!' include row is never a signal.
            Assert.DoesNotContain(sheet.CsvOnlyRows, r => r.Signal.Mnemonic == "OTHER_CARD");
        }

        [Fact]
        public void One_name_of_a_multi_name_point_still_resolves_to_its_row()
        {
            // Per-record actions (expand-sheet deletes/creates) identify by a SINGLE name; the index
            // must alias each part of a multi-mnemonic point to the same spine row.
            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "", "I.60.0.0", "1", "AUTOM/MANUAL" },
                },
            };
            var model = IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)));

            BridgeRow byPart = model.FindByIdentity(
                new SignalRef { CardKey = "R1.S1.C1", Address = "", Mnemonic = "MANUAL" })!;
            Assert.NotNull(byPart);
            Assert.Equal("AUTOM/MANUAL", byPart.Signal.Mnemonic);
        }

        [Fact]
        public void The_signal_index_resolves_by_address_and_by_name()
        {
            BridgeWorkspaceModel model = Build();

            BridgeRow byAddress = model.FindByIdentity(
                new SignalRef { CardKey = "R1.S1.C1", Address = "I.90.0.1" })!;
            Assert.Equal("LP100301", byAddress.Signal.Mnemonic);

            // Same signal, stale address, known name — the mnemonic fallback finds it.
            BridgeRow byName = model.FindByIdentity(
                new SignalRef { CardKey = "R1.S1.C1", Address = "", Mnemonic = "lp100301" })!;
            Assert.Same(byAddress, byName);

            Assert.Null(model.FindByIdentity(
                new SignalRef { CardKey = "R1.S1.C1", Address = "I.9.9.9", Mnemonic = "NOBODY" }));
        }

        [Fact]
        public void Canonical_field_headers_carry_each_files_real_names()
        {
            BridgeWorkspaceModel model = Build();

            Assert.Equal("signalid", model.InputsFieldHeaders[CsvCanonical.Name]);
            Assert.Equal("outputname", model.OutputsFieldHeaders[CsvCanonical.Name]);
            Assert.Equal("logic", model.InputsFieldHeaders[CsvCanonical.Logic]);
            Assert.Equal("ncnologic", model.OutputsFieldHeaders[CsvCanonical.Logic]);
            Assert.Equal("card", model.InputsFieldHeaders[CsvCanonical.Card]);
            Assert.Equal("boardnumber", model.OutputsFieldHeaders[CsvCanonical.Card]);
        }

        [Fact]
        public void Several_csv_rows_at_one_address_all_join_and_the_names_merge()
        {
            // Two CSV records share the physical point 60.0.0 — the AUTOM/MANUAL pattern.
            const string twinInputs =
                "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n" +
                "@,AUTOM,1,0,0,0,0,0,0,0,0,60,0,0,0,1,AUTO MODE\r\n" +
                "@,MANUAL,1,1,0,0,0,0,0,0,0,60,0,0,0,1,MANUAL MODE\r\n";

            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "", "I.60.0.0", "1", "AUTOM/MANUAL" },
                },
            };
            var model = IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(twinInputs)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)));

            BridgeSheetModel sheet = model.Sheets.Single();
            BridgeRow row = sheet.SpineRows.Single();

            Assert.NotNull(row.Csv);
            Assert.True(row.Csv!.IsMulti);
            Assert.Equal(2, row.Csv.RecordIndexes.Count);
            Assert.Equal("AUTOM/MANUAL", row.CsvValues[CsvCanonical.Name]);
            Assert.Equal("60", row.CsvValues[CsvCanonical.Node]);      // identical values collapse
            Assert.Equal("0/1", row.CsvValues[CsvCanonical.Logic]);    // differing values join
            Assert.Equal("", row.MatchVerdict);                        // {AUTOM,MANUAL} == {AUTOM,MANUAL}
            Assert.Contains("2 rows", row.RowLabel);

            // Both twin records are CLAIMED, not orphaned (the unrelated Outputs row still buckets).
            Assert.DoesNotContain(sheet.CsvOnlyRows, r => r.Csv!.File == BridgeTarget.InputsCsv);
        }

        [Fact]
        public void A_true_mismatch_on_a_multi_row_point_still_shows_a_ratio()
        {
            const string twinInputs =
                "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n" +
                "@,AUTOM,1,0,0,0,0,0,0,0,0,60,0,0,0,1,AUTO MODE\r\n" +
                "@,SOMETHING_ELSE,1,1,0,0,0,0,0,0,0,60,0,0,0,1,OTHER\r\n";

            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "", "I.60.0.0", "1", "AUTOM/MANUAL" },
                },
            };
            BridgeRow row = IoBridgeJoin.Build(new[] { card },
                    CsvFile.Load(Encoding.UTF8.GetBytes(twinInputs)),
                    CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)))
                .Sheets.Single().SpineRows.Single();

            Assert.Equal("1/2", row.MatchVerdict);   // MANUAL is genuinely absent from the CSV side
        }

        [Fact]
        public void A_struct_binding_row_never_steals_the_named_signals_record()
        {
            // The SIREN pattern: two IO rows share one address — a DEO struct binding point first,
            // the actual DO (SIREN01) below. The CSV carries both records at that address.
            const string sirenOutputs =
                "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
                "@;SIREN01;2;0;90;5;0;0;1;THE ACTUAL SIREN\r\n" +
                "@;DEO_SIREN_STRUCT;8;0;90;5;0;0;1;STRUCT BINDING\r\n";

            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "", "Q.90.5.0", "0", "DEO_SIREN_STRUCT" },   // the struct row, FIRST
                    new[] { "", "Q.90.5.0", "0", "SIREN01" },            // the real DO below it
                },
            };
            var model = IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(sirenOutputs)));

            var rows = model.Sheets.Single().SpineRows;
            Assert.Equal("DEO_SIREN_STRUCT", rows[0].CsvValues[CsvCanonical.Name]);
            Assert.Equal("", rows[0].MatchVerdict);
            Assert.Equal("SIREN01", rows[1].CsvValues[CsvCanonical.Name]);   // NOT stolen by the row above
            Assert.Equal("", rows[1].MatchVerdict);
        }

        [Fact]
        public void With_only_the_do_record_present_the_named_row_wins_and_the_struct_reads_missing()
        {
            const string sirenOutputs =
                "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
                "@;SIREN01;2;0;90;5;0;0;1;THE ACTUAL SIREN\r\n";

            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "", "Q.90.5.0", "0", "DEO_SIREN_STRUCT" },
                    new[] { "", "Q.90.5.0", "0", "SIREN01" },
                },
            };
            var rows = IoBridgeJoin.Build(new[] { card },
                    CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                    CsvFile.Load(Encoding.UTF8.GetBytes(sirenOutputs)))
                .Sheets.Single().SpineRows;

            Assert.Equal("missing", rows[0].MatchVerdict);               // the struct row has no record
            Assert.Equal("SIREN01", rows[1].CsvValues[CsvCanonical.Name]);
            Assert.Equal("", rows[1].MatchVerdict);
        }

        [Fact]
        public void Opted_in_virtuals_join_clean_and_are_flagged_twins_without_the_regex()
        {
            const string virtInputs =
                "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n" +
                "@,CLASSDIAGAL00620030130,5,0,0,0,0,0,0,0,0,0,0,0,0,1,ASI FAULT A\r\n" +
                "@,CLASSDIAGAL00620030230,5,0,0,0,0,0,0,0,0,0,0,0,0,1,ASI FAULT B\r\n";

            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "ID", "Bit", "Normal condition", "Mnemonic" },
                    new[] { "", "I.90.1.2", "1", "CLASSDIAGAL00620030130/CLASSDIAGAL00620030230" },
                },
            };

            // Without the regex the virtuals are not EXPECTED — they still carry the row's names,
            // so they bind as its (flagged) twins rather than leaving the row "missing".
            BridgeRow without = IoBridgeJoin.Build(new[] { card },
                    CsvFile.Load(Encoding.UTF8.GetBytes(virtInputs)),
                    CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)))
                .Sheets.Single().SpineRows.Single();
            Assert.NotNull(without.Csv);
            Assert.True(without.Csv!.Virtual);
            Assert.Equal(2, without.Csv.RecordIndexes.Count);
            Assert.Equal("virtual", without.MatchVerdict);
            Assert.EndsWith("· virtual twin", without.RowLabel);

            // With the regex: both virtual records attach BY NAME and the row agrees.
            BridgeRow with = IoBridgeJoin.Build(new[] { card },
                    CsvFile.Load(Encoding.UTF8.GetBytes(virtInputs)),
                    CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)),
                    new IoBridgeJoinOptions { IncludeVirtualsMatching = "^CLASSDIAGAL" })
                .Sheets.Single().SpineRows.Single();

            Assert.NotNull(with.Csv);
            Assert.Equal(2, with.Csv!.RecordIndexes.Count);
            Assert.Equal("CLASSDIAGAL00620030130/CLASSDIAGAL00620030230", with.CsvValues[CsvCanonical.Name]);
            Assert.Equal("", with.MatchVerdict);
        }

        [Fact]
        public void Excluded_io_rows_are_shown_but_ignored_by_the_comparison()
        {
            var model = IoBridgeJoin.Build(new[] { Card() },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)),
                new IoBridgeJoinOptions { ExcludeIoMnemonicsMatching = "^LP1" });

            BridgeSheetModel sheet = model.Sheets.Single();
            BridgeRow lp = sheet.SpineRows.Single(r => r.Signal.Mnemonic == "LP100301");

            Assert.Contains("excluded", lp.RowLabel);
            Assert.Null(lp.Csv);                        // claims nothing
            Assert.Equal("", lp.MatchVerdict);          // no verdict, not even "missing"

            // Its CSV record is unclaimed and lands in the CSV-only bucket instead.
            Assert.Contains(sheet.CsvOnlyRows, r => r.Signal.Mnemonic == "LP100301");
        }

        [Fact]
        public void Spare_points_with_empty_descriptions_and_device_read_unused()
        {
            var card = new CardSheetGrid
            {
                SheetName = "R1.S1.C1",
                AllRows = new List<string[]>
                {
                    new[] { "Bit", "Mnemonic", "Description language 1 (1° part) (permanent part)",
                            "Description language 1 (2° part) (variable part)", "Device" },
                    new[] { "I.90.0.1", "LP100301", "LAMP", "301", "-K1" },      // in use
                    new[] { "I.90.0.2", "", "", "", "" },                        // spare — unused
                    new[] { "I.90.0.3", "SIG3", "", "SOMETHING", "" },           // part 2 filled → used
                },
            };
            var model = IoBridgeJoin.Build(new[] { card },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)));

            var rows = model.Sheets.Single().SpineRows;
            Assert.DoesNotContain("unused", rows[0].RowLabel);
            Assert.EndsWith("· unused", rows[1].RowLabel);
            Assert.DoesNotContain("unused", rows[2].RowLabel);
        }

        [Fact]
        public void Unused_is_not_judged_when_the_columns_are_absent()
        {
            // The default test card sheet has no description/device columns at all.
            foreach (BridgeRow row in Build().Sheets.Single().SpineRows)
                Assert.DoesNotContain("unused", row.RowLabel);
        }

        [Fact]
        public void A_sheet_whose_name_is_not_a_card_is_ignored()
        {
            var model = IoBridgeJoin.Build(
                new[] { new CardSheetGrid { SheetName = "COVER", AllRows = Card().AllRows } },
                CsvFile.Load(Encoding.UTF8.GetBytes(InputsText)),
                CsvFile.Load(Encoding.UTF8.GetBytes(OutputsText)));

            Assert.Empty(model.Sheets);
        }
    }
}
