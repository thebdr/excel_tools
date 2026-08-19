using System.Linq;
using ExcelPowerTools.Core.Experimental.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental.IoBridge
{
    public class IoBridgeQueueJsonTests
    {
        private static IoBridgeSidecar Sample()
        {
            var sidecar = new IoBridgeSidecar
            {
                GeneratedUtc = "2026-08-19T10:00:00.0000000Z",
                IoListPath = @"C:\docs\8FVZ_IO0001 - sorter.xlsx",
                InputsCsvPath = @"C:\sw\Plant_Inputs.csv",
                OutputsCsvPath = @"C:\sw\Plant_Outputs.csv",
            };
            sidecar.CardSheets.Add("R1.S1.C1");

            sidecar.Queue.Actions.Add(new BridgeAction
            {
                ActionId = "a1",
                Kind = BridgeActionKind.SetField,
                Target = BridgeTarget.InputsCsv,
                Signal = new SignalRef { CardKey = "R1.S1.C1", Address = "I.80.3.0", Mnemonic = "LP100301" },
                Linked = LinkedField.Mnemonic,
                FieldHeader = "signalid",
                OldValue = "LP100301",
                NewValue = "LP100302",
                QueuedUtc = "2026-08-19T10:01:00.0000000Z",
            });
            sidecar.Queue.Actions.Add(new BridgeAction
            {
                ActionId = "a2",
                Kind = BridgeActionKind.CreateCsvRow,
                Target = BridgeTarget.OutputsCsv,
                Signal = new SignalRef { CardKey = "R1.S1.C1", Address = "O.90.0.1", Mnemonic = "NEW_SIGNAL" },
                OldValue = "",
                NewValue = "+",
                RowFields = new[] { "@", "NEW_SIGNAL", "2", "0", "90", "0", "1", "0", "1", "descr \"quoted\" / slash" },
                Include = false,
                Status = BridgeActionStatus.Conflict,
                StatusNote = "file now holds \"X\"",
            });
            return sidecar;
        }

        [Fact]
        public void Round_trips_every_field()
        {
            string json = IoBridgeQueueJson.Serialize(Sample());
            Assert.True(IoBridgeQueueJson.TryParse(json, out IoBridgeSidecar back, out string error), error);

            Assert.Equal("2026-08-19T10:00:00.0000000Z", back.GeneratedUtc);
            Assert.Equal(@"C:\docs\8FVZ_IO0001 - sorter.xlsx", back.IoListPath);
            Assert.Equal(new[] { "R1.S1.C1" }, back.CardSheets);
            Assert.Equal(2, back.Queue.Actions.Count);
            Assert.Empty(back.LoadNotes);

            BridgeAction a1 = back.Queue.Actions[0];
            Assert.Equal("a1", a1.ActionId);
            Assert.Equal(BridgeActionKind.SetField, a1.Kind);
            Assert.Equal(BridgeTarget.InputsCsv, a1.Target);
            Assert.Equal("R1.S1.C1|I.80.3.0", a1.Signal.IdentityKey());
            Assert.Equal(LinkedField.Mnemonic, a1.Linked);
            Assert.Equal("signalid", a1.FieldHeader);
            Assert.Equal("LP100301", a1.OldValue);
            Assert.Equal("LP100302", a1.NewValue);
            Assert.True(a1.Include);
            Assert.Equal(BridgeActionStatus.Pending, a1.Status);
            Assert.Null(a1.RowFields);

            BridgeAction a2 = back.Queue.Actions[1];
            Assert.Equal(BridgeActionKind.CreateCsvRow, a2.Kind);
            Assert.False(a2.Include);
            Assert.Equal(BridgeActionStatus.Conflict, a2.Status);
            Assert.Equal("file now holds \"X\"", a2.StatusNote);
            Assert.Equal("descr \"quoted\" / slash", a2.RowFields![9]);
        }

        [Fact]
        public void Serialization_is_deterministic()
        {
            Assert.Equal(IoBridgeQueueJson.Serialize(Sample()), IoBridgeQueueJson.Serialize(Sample()));
        }

        [Fact]
        public void Unknown_properties_are_ignored()
        {
            string json = IoBridgeQueueJson.Serialize(Sample());
            json = json.Replace("\"generatedUtc\"", "\"futureThing\":{\"nested\":[1,2]},\"generatedUtc\"");

            Assert.True(IoBridgeQueueJson.TryParse(json, out IoBridgeSidecar back, out string error), error);
            Assert.Equal(2, back.Queue.Actions.Count);
        }

        [Fact]
        public void Unknown_enum_values_quarantine_the_action_not_the_file()
        {
            string json = IoBridgeQueueJson.Serialize(Sample());
            json = json.Replace("\"kind\":\"CreateCsvRow\"", "\"kind\":\"TeleportRow\"");

            Assert.True(IoBridgeQueueJson.TryParse(json, out IoBridgeSidecar back, out string error), error);
            Assert.Single(back.Queue.Actions);
            Assert.Equal("a1", back.Queue.Actions[0].ActionId);
            string note = Assert.Single(back.LoadNotes);
            Assert.Contains("TeleportRow", note);
        }

        [Fact]
        public void Numeric_enum_aliases_are_not_accepted()
        {
            string json = IoBridgeQueueJson.Serialize(Sample());
            json = json.Replace("\"kind\":\"SetField\"", "\"kind\":\"0\"");

            Assert.True(IoBridgeQueueJson.TryParse(json, out IoBridgeSidecar back, out _));
            Assert.Single(back.Queue.Actions);          // the tampered action is skipped
            Assert.Single(back.LoadNotes);
        }

        [Fact]
        public void A_different_schema_token_refuses_the_load()
        {
            string json = IoBridgeQueueJson.Serialize(Sample())
                .Replace(IoBridgeQueueJson.SchemaToken, "io-bridge-queue-v99");

            Assert.False(IoBridgeQueueJson.TryParse(json, out _, out string error));
            Assert.Contains("io-bridge-queue-v99", error);
        }

        [Fact]
        public void Malformed_json_reports_a_parse_error()
        {
            Assert.False(IoBridgeQueueJson.TryParse("{\"schema\":", out _, out string error));
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void An_empty_sidecar_round_trips()
        {
            string json = IoBridgeQueueJson.Serialize(new IoBridgeSidecar());
            Assert.True(IoBridgeQueueJson.TryParse(json, out IoBridgeSidecar back, out string error), error);
            Assert.Empty(back.Queue.Actions);
            Assert.Empty(back.CardSheets);
        }
    }
}
