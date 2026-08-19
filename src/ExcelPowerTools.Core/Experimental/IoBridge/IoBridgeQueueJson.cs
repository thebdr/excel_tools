using System;
using System.Collections.Generic;
using System.Text;
using ExcelPowerTools.Core.Text;

namespace ExcelPowerTools.Core.Experimental.IoBridge
{
    /// <summary>Everything the sidecar persists: the document set and the action queue.</summary>
    public sealed class IoBridgeSidecar
    {
        public string GeneratedUtc = "";
        public string IoListPath = "";
        public string InputsCsvPath = "";
        public string OutputsCsvPath = "";
        public List<string> CardSheets = new List<string>();
        public IoBridgeQueue Queue = new IoBridgeQueue();

        /// <summary>Load-time observations (skipped actions, tolerated oddities). Never serialized.</summary>
        public List<string> LoadNotes = new List<string>();
    }

    /// <summary>
    /// Sidecar (de)serialization. Versioning follows the house pattern: the literal schema token only
    /// changes for breaking shape changes; ADDITIVE properties do not bump it, and the reader ignores
    /// properties it does not know. An action carrying an unknown enum value is skipped (reported via
    /// LoadNotes) rather than failing the whole file; a different schema token refuses the load.
    /// </summary>
    public static class IoBridgeQueueJson
    {
        public const string SchemaToken = "io-bridge-queue-v1";

        public static string Serialize(IoBridgeSidecar sidecar)
        {
            if (sidecar == null) throw new ArgumentNullException(nameof(sidecar));

            var json = new StringBuilder(4096);
            json.Append('{');
            MiniJsonWriter.Property(json, "schema", SchemaToken, first: true);
            MiniJsonWriter.Property(json, "generatedUtc", sidecar.GeneratedUtc, first: false);
            MiniJsonWriter.Property(json, "ioListPath", sidecar.IoListPath, first: false);
            MiniJsonWriter.Property(json, "inputsCsvPath", sidecar.InputsCsvPath, first: false);
            MiniJsonWriter.Property(json, "outputsCsvPath", sidecar.OutputsCsvPath, first: false);

            json.Append(',');
            MiniJsonWriter.AppendString(json, "cardSheets");
            json.Append(":[");
            for (int i = 0; i < sidecar.CardSheets.Count; i++)
            {
                if (i > 0) json.Append(',');
                MiniJsonWriter.AppendString(json, sidecar.CardSheets[i] ?? "");
            }
            json.Append(']');

            json.Append(',');
            MiniJsonWriter.AppendString(json, "actions");
            json.Append(":[");
            for (int i = 0; i < sidecar.Queue.Actions.Count; i++)
            {
                if (i > 0) json.Append(',');
                AppendAction(json, sidecar.Queue.Actions[i]);
            }
            json.Append(']');

            json.Append('}');
            return json.ToString();
        }

        private static void AppendAction(StringBuilder json, BridgeAction a)
        {
            json.Append('{');
            MiniJsonWriter.Property(json, "actionId", a.ActionId, first: true);
            MiniJsonWriter.Property(json, "kind", a.Kind.ToString(), first: false);
            MiniJsonWriter.Property(json, "target", a.Target.ToString(), first: false);
            MiniJsonWriter.Property(json, "cardKey", a.Signal.CardKey, first: false);
            MiniJsonWriter.Property(json, "address", a.Signal.Address, first: false);
            MiniJsonWriter.Property(json, "mnemonic", a.Signal.Mnemonic, first: false);
            MiniJsonWriter.Property(json, "linked", a.Linked.ToString(), first: false);
            MiniJsonWriter.Property(json, "fieldHeader", a.FieldHeader, first: false);
            MiniJsonWriter.Property(json, "oldValue", a.OldValue, first: false);
            MiniJsonWriter.Property(json, "newValue", a.NewValue, first: false);

            if (a.RowFields != null)
            {
                json.Append(',');
                MiniJsonWriter.AppendString(json, "rowFields");
                json.Append(":[");
                for (int i = 0; i < a.RowFields.Length; i++)
                {
                    if (i > 0) json.Append(',');
                    MiniJsonWriter.AppendString(json, a.RowFields[i] ?? "");
                }
                json.Append(']');
            }

            MiniJsonWriter.BoolProperty(json, "include", a.Include, first: false);
            MiniJsonWriter.Property(json, "status", a.Status.ToString(), first: false);
            MiniJsonWriter.Property(json, "statusNote", a.StatusNote, first: false);
            MiniJsonWriter.Property(json, "queuedUtc", a.QueuedUtc, first: false);
            MiniJsonWriter.Property(json, "appliedUtc", a.AppliedUtc, first: false);
            json.Append('}');
        }

        public static bool TryParse(string json, out IoBridgeSidecar sidecar, out string error)
        {
            sidecar = null;
            if (!MiniJson.TryParse(json, out JsonValue root, out error)) return false;
            if (root.Kind != JsonKind.Object)
            {
                error = "The sidecar is not a JSON object.";
                return false;
            }

            string schema = root.Str("schema");
            if (schema != SchemaToken)
            {
                error = schema.Length == 0
                    ? "The sidecar carries no schema token."
                    : "Unsupported sidecar schema \"" + schema + "\" (this build reads \"" + SchemaToken + "\").";
                return false;
            }

            var result = new IoBridgeSidecar
            {
                GeneratedUtc = root.Str("generatedUtc"),
                IoListPath = root.Str("ioListPath"),
                InputsCsvPath = root.Str("inputsCsvPath"),
                OutputsCsvPath = root.Str("outputsCsvPath"),
            };

            foreach (JsonValue sheet in root.Arr("cardSheets"))
                if (sheet.Kind == JsonKind.String) result.CardSheets.Add(sheet.Text);

            int index = 0;
            foreach (JsonValue entry in root.Arr("actions"))
            {
                index++;
                if (entry.Kind != JsonKind.Object)
                {
                    result.LoadNotes.Add("Action " + index + " skipped: not an object.");
                    continue;
                }

                if (!TryParseEnum(entry.Str("kind"), out BridgeActionKind kind) ||
                    !TryParseEnum(entry.Str("target"), out BridgeTarget target) ||
                    !TryParseEnum(entry.Str("status"), out BridgeActionStatus status))
                {
                    result.LoadNotes.Add(
                        "Action " + index + " skipped: unknown kind/target/status " +
                        "(\"" + entry.Str("kind") + "\"/\"" + entry.Str("target") + "\"/\"" + entry.Str("status") + "\").");
                    continue;
                }

                TryParseEnum(entry.Str("linked"), out LinkedField linked);   // unknown → None, harmless

                var action = new BridgeAction
                {
                    ActionId = FirstNonEmpty(entry.Str("actionId"), Guid.NewGuid().ToString("N")),
                    Kind = kind,
                    Target = target,
                    Signal = new SignalRef
                    {
                        CardKey = entry.Str("cardKey"),
                        Address = entry.Str("address"),
                        Mnemonic = entry.Str("mnemonic"),
                    },
                    Linked = linked,
                    FieldHeader = entry.Str("fieldHeader"),
                    OldValue = entry.Str("oldValue"),
                    NewValue = entry.Str("newValue"),
                    Include = entry.BoolOf("include", true),
                    Status = status,
                    StatusNote = entry.Str("statusNote"),
                    QueuedUtc = entry.Str("queuedUtc"),
                    AppliedUtc = entry.Str("appliedUtc"),
                };

                JsonValue rowFields = entry.Get("rowFields");
                if (rowFields != null && rowFields.Kind == JsonKind.Array)
                {
                    var fields = new List<string>(rowFields.Items.Count);
                    foreach (JsonValue f in rowFields.Items)
                        fields.Add(f.Kind == JsonKind.String || f.Kind == JsonKind.Number ? f.Text : "");
                    action.RowFields = fields.ToArray();
                }

                result.Queue.Actions.Add(action);
            }

            sidecar = result;
            error = null;
            return true;
        }

        private static bool TryParseEnum<T>(string text, out T value) where T : struct
        {
            // Exact names only. Enum.TryParse also accepts numeric strings, which would let a value
            // from a newer build round-trip as a bogus member — reject anything not starting with a
            // letter before parsing.
            value = default(T);
            if (string.IsNullOrEmpty(text) || !char.IsLetter(text[0])) return false;
            return Enum.TryParse(text, ignoreCase: false, out value) && Enum.IsDefined(typeof(T), value);
        }

        private static string FirstNonEmpty(string preferred, string fallback) =>
            string.IsNullOrEmpty(preferred) ? fallback : preferred;
    }
}
