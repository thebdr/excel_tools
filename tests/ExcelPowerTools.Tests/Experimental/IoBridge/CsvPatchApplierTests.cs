using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.Experimental.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental.IoBridge
{
    public class CsvPatchApplierTests
    {
        private const string InputsText =
            "#,Signal ID,Type,Logic,RE Filter,FE Filter,A1 Filter,A0 Filter,Events,Queue,Flags,Node,Offset,Bit,Default Virtual Value,Card,Description\r\n" +
            "@,LP100301,1,0,0,0,0,0,0,0,0,90,0,1,0,1,LAMP 301\r\n" +
            "@,CABINETAL00161,11,0,0,0,0,0,0,0,0,236,0,0,0,1,FIELD BUS FAILURE\r\n" +
            "#,DISABLED_SIG,1,1,0,0,0,0,0,0,0,11,1,1,0,1,DISABLED ONE\r\n" +
            "@,SHORT_ROW,1,0,0,0,0,0,0,0,0,50,2,3,0,1\r\n" +          // 16 fields — no Description
            "@,TWIN,1,0,0,0,0,0,0,0,0,60,0,0,0,1,FIRST TWIN\r\n" +
            "@,TWIN,1,0,0,0,0,0,0,0,0,60,0,1,0,1,SECOND TWIN\r\n";

        private static CsvDocument Load() => CsvFile.Load(Encoding.UTF8.GetBytes(InputsText));

        private static SignalRef At(string address, string mnemonic = "") =>
            new SignalRef { CardKey = "R1.S1.C1", Address = address, Mnemonic = mnemonic };

        private static BridgeAction Set(SignalRef signal, string field, string oldV, string newV) =>
            new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = BridgeTarget.InputsCsv,
                Signal = signal,
                FieldHeader = field,
                OldValue = oldV,
                NewValue = newV,
            };

        private static CsvPatchResult Apply(CsvDocument doc, params BridgeAction[] actions) =>
            CsvPatchApplier.Apply(doc, actions, CsvColumnsMap.Locate(doc)!);

        [Fact]
        public void Map_locates_both_export_schemas()
        {
            CsvColumnsMap inputs = CsvColumnsMap.Locate(Load())!;
            Assert.Equal(1, inputs.Name);
            Assert.Equal(15, inputs.Board);
            Assert.Equal(16, inputs.FieldFor("description"));

            var outputs = CsvColumnsMap.Locate(CsvFile.Load(Encoding.UTF8.GetBytes(
                "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
                "@;DOOR_CMD_01;2;0;235;10;1;0;1;DOOR CMD\r\n")))!;
            Assert.Equal(1, outputs.Name);
            Assert.Equal(8, outputs.Board);
            Assert.Equal(3, outputs.FieldFor("ncnologic"));
        }

        [Fact]
        public void Set_field_applies_when_the_file_still_holds_the_old_value()
        {
            CsvDocument doc = Load();
            CsvPatchResult result = Apply(doc, Set(At("I.90.0.1"), "signalid", "LP100301", "LP100302"));

            Assert.True(result.AnyMutation);
            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes.Single().Outcome);
            Assert.Equal("LP100302", doc.Records[1].Fields[1]);

            byte[] outBytes = doc.Save();
            Assert.Null(CsvDiff.VerifyRoundTrip(outBytes, doc.Format, doc.Records));
        }

        [Fact]
        public void Set_field_conflicts_when_the_file_drifted()
        {
            CsvDocument doc = Load();
            CsvPatchResult result = Apply(doc, Set(At("I.90.0.1"), "signalid", "SOMETHING_ELSE", "LP100302"));

            Assert.False(result.AnyMutation);
            CsvActionOutcome outcome = result.Outcomes.Single();
            Assert.Equal(BridgeActionStatus.Conflict, outcome.Outcome);
            Assert.Contains("LP100301", outcome.Note);
            Assert.Equal("LP100301", doc.Records[1].Fields[1]);   // untouched
        }

        [Fact]
        public void Set_field_detects_an_external_apply()
        {
            CsvDocument doc = Load();
            CsvPatchResult result = Apply(doc, Set(At("I.90.0.1"), "signalid", "OLD_NAME", "LP100301"));

            Assert.False(result.AnyMutation);
            Assert.Equal(BridgeActionStatus.AppliedExternally, result.Outcomes.Single().Outcome);
        }

        [Fact]
        public void Missing_column_and_missing_signal_conflict()
        {
            CsvDocument doc = Load();
            CsvPatchResult result = Apply(doc,
                Set(At("I.90.0.1"), "nosuchcolumn", "a", "b"),
                Set(At("I.99.9.9"), "signalid", "a", "b"));

            Assert.Equal(BridgeActionStatus.Conflict, result.Outcomes[0].Outcome);
            Assert.Contains("nosuchcolumn", result.Outcomes[0].Note);
            Assert.Equal(BridgeActionStatus.Conflict, result.Outcomes[1].Outcome);
            Assert.Contains("not found", result.Outcomes[1].Note);
        }

        [Fact]
        public void A_module_signal_resolves_by_node()
        {
            CsvDocument doc = Load();
            CsvPatchResult result = Apply(doc,
                Set(At("#236", "CABINETAL00161"), "description", "FIELD BUS FAILURE", "FIELD BUS FAILURE 2"));

            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes.Single().Outcome);
            Assert.Equal("FIELD BUS FAILURE 2", doc.Records[2].Fields[16]);
        }

        [Fact]
        public void A_readdressed_signal_falls_back_to_its_mnemonic()
        {
            CsvDocument doc = Load();
            // The action remembers the OLD address; the file has since moved the signal — the name finds it.
            CsvPatchResult result = Apply(doc,
                Set(At("I.90.9.9", "LP100301"), "description", "LAMP 301", "LAMP 301 MOVED"));

            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes.Single().Outcome);
            Assert.Equal("LAMP 301 MOVED", doc.Records[1].Fields[16]);
        }

        [Fact]
        public void An_ambiguous_mnemonic_is_an_error_not_a_guess()
        {
            CsvDocument doc = Load();
            CsvPatchResult result = Apply(doc,
                Set(At("", "TWIN"), "description", "FIRST TWIN", "X"));

            CsvActionOutcome outcome = result.Outcomes.Single();
            Assert.Equal(BridgeActionStatus.Conflict, outcome.Outcome);
            Assert.Contains("ambiguous", outcome.Note);
            Assert.False(result.AnyMutation);
        }

        [Fact]
        public void Widening_a_short_record_keeps_its_terminator_and_round_trips()
        {
            CsvDocument doc = Load();
            string? terminator = doc.Records[4].Terminator;
            Assert.Equal("\r\n", terminator);
            Assert.Equal(16, doc.Records[4].FieldCount);

            CsvPatchResult result = Apply(doc,
                Set(At("I.50.2.3", "SHORT_ROW"), "description", "", "NOW DESCRIBED"));

            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes.Single().Outcome);
            Assert.Equal(17, doc.Records[4].FieldCount);
            Assert.Equal("NOW DESCRIBED", doc.Records[4].Fields[16]);
            Assert.Equal(terminator, doc.Records[4].Terminator);
            Assert.Null(CsvDiff.VerifyRoundTrip(doc.Save(), doc.Format, doc.Records));
        }

        [Fact]
        public void Disable_swaps_the_marker_and_enable_swaps_it_back()
        {
            CsvDocument doc = Load();
            var disable = new BridgeAction
            {
                Kind = BridgeActionKind.DisableCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.90.0.1"),
            };

            Assert.Equal(BridgeActionStatus.Applied, Apply(doc, disable).Outcomes.Single().Outcome);
            Assert.Equal("#", doc.Records[1].Fields[0]);

            Assert.Equal(BridgeActionStatus.AppliedExternally, Apply(doc, disable).Outcomes.Single().Outcome);

            var enable = new BridgeAction
            {
                Kind = BridgeActionKind.EnableCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.90.0.1"),
            };
            Assert.Equal(BridgeActionStatus.Applied, Apply(doc, enable).Outcomes.Single().Outcome);
            Assert.Equal("@", doc.Records[1].Fields[0]);
        }

        [Fact]
        public void A_disabled_row_still_resolves_so_enable_can_find_it()
        {
            CsvDocument doc = Load();
            var enable = new BridgeAction
            {
                Kind = BridgeActionKind.EnableCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.11.1.1", "DISABLED_SIG"),
            };

            Assert.Equal(BridgeActionStatus.Applied, Apply(doc, enable).Outcomes.Single().Outcome);
            Assert.Equal("@", doc.Records[3].Fields[0]);
        }

        [Fact]
        public void Create_appends_and_an_existing_identity_reports_externally_applied()
        {
            CsvDocument doc = Load();
            int before = doc.Records.Count;
            var create = new BridgeAction
            {
                Kind = BridgeActionKind.CreateCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.70.0.0", "NEW_SIGNAL"),
                OldValue = "",
                NewValue = "+",
                RowFields = new[] { "@", "NEW_SIGNAL", "1", "0", "0", "0", "0", "0", "0", "0", "0", "70", "0", "0", "0", "1", "FRESH" },
            };

            Assert.Equal(BridgeActionStatus.Applied, Apply(doc, create).Outcomes.Single().Outcome);
            Assert.Equal(before + 1, doc.Records.Count);
            Assert.Equal("NEW_SIGNAL", doc.Records[before].Fields[1]);
            Assert.Null(CsvDiff.VerifyRoundTrip(doc.Save(), doc.Format, doc.Records));

            var again = new BridgeAction
            {
                Kind = BridgeActionKind.CreateCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.90.0.1", "LP100301"),
                RowFields = create.RowFields,
            };
            Assert.Equal(BridgeActionStatus.AppliedExternally, Apply(doc, again).Outcomes.Single().Outcome);
        }

        [Fact]
        public void A_create_payload_without_a_signal_name_is_refused()
        {
            CsvDocument doc = Load();
            int before = doc.Records.Count;
            var nameless = new BridgeAction
            {
                Kind = BridgeActionKind.CreateCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.70.0.0", ""),
                RowFields = new[] { "@", "  ", "1", "0", "0", "0", "0", "0", "0", "0", "0", "70", "0", "0", "0", "1", "" },
            };

            CsvActionOutcome outcome = Apply(doc, nameless).Outcomes.Single();

            Assert.Equal(BridgeActionStatus.Conflict, outcome.Outcome);
            Assert.Contains("no signal name", outcome.Note);
            Assert.Equal(before, doc.Records.Count);   // nothing appended
        }

        [Fact]
        public void Delete_removes_and_later_actions_still_resolve()
        {
            CsvDocument doc = Load();
            var delete = new BridgeAction
            {
                Kind = BridgeActionKind.DeleteCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = At("I.90.0.1", "LP100301"),
            };
            // The second action targets a record AFTER the deleted one — its index shifts.
            BridgeAction after = Set(At("I.11.1.1", "DISABLED_SIG"), "description", "DISABLED ONE", "STILL FOUND");

            CsvPatchResult result = Apply(doc, delete, after);

            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes[0].Outcome);
            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes[1].Outcome);
            Assert.DoesNotContain(doc.Records, r => r.Fields.Length > 1 && r.Fields[1] == "LP100301");
            Assert.Contains(doc.Records, r => r.Fields.Length > 16 && r.Fields[16] == "STILL FOUND");

            Assert.Equal(BridgeActionStatus.AppliedExternally, Apply(doc, delete).Outcomes.Single().Outcome);
        }

        [Fact]
        public void The_semicolon_schema_round_trips_after_patching()
        {
            CsvDocument doc = CsvFile.Load(Encoding.UTF8.GetBytes(
                "#;Output Name;Output Type;NC/NO Logic;Node;Offset;Bit;Default Value (For Virtual I/Os);Board number;Description\r\n" +
                "@;DOOR_CMD_01;2;0;235;10;1;0;1;DOOR CMD\r\n"));

            CsvPatchResult result = CsvPatchApplier.Apply(doc, new[]
            {
                new BridgeAction
                {
                    Kind = BridgeActionKind.SetField,
                    Target = BridgeTarget.OutputsCsv,
                    Signal = At("O.235.10.1"),
                    FieldHeader = "outputname",
                    OldValue = "DOOR_CMD_01",
                    NewValue = "DOOR_CMD_01B",
                },
            }, CsvColumnsMap.Locate(doc)!);

            Assert.Equal(BridgeActionStatus.Applied, result.Outcomes.Single().Outcome);
            byte[] outBytes = doc.Save();
            Assert.Null(CsvDiff.VerifyRoundTrip(outBytes, doc.Format, doc.Records));
            Assert.Contains(";DOOR_CMD_01B;", Encoding.UTF8.GetString(outBytes));
        }
    }
}
