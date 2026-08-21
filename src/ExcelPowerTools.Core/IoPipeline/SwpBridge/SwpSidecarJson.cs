using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using ExcelPowerTools.Core.Text;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// Everything the swp-bridge.json sidecar persists for one system. The SWP and CONFIG_PLANT
    /// live as folder-relative names plus the current dated version folder, so version-folder
    /// hops only update <see cref="CurrentVersionFolder"/>; the IO list is an absolute path.
    /// </summary>
    public sealed class SwpBridgeSettings
    {
        public SwpPlantType PlantType = SwpPlantType.Msc;
        public string CurrentVersionFolder = "";
        public string SwpFileName = "";
        public string ConfigFileName = "";
        public string IoListPath = "";
        public List<string> CardSheets = new List<string>();
        public string PlantInputsCsvPath = "";
        public string BusfieldCsvPath = "";
        public string DestinationParamsCsvPath = "";
        public string SensorCsvPath = "";
        public string StrobeDeviceCsvPath = "";
        public int FirstLineId = 1;
        public int GroupIndex = 1;
        public string IncludeRegex = "";
        public string ExcludeRegex = "";
        public bool TrustConfigPlant;
        public bool UpdateDescriptionFormulas;
        public SwpMismatchMode MismatchMode = SwpMismatchMode.AppendToDescription;
        public bool ApplyToVersionFolder;    // false = in-place with .bak

        /// <summary>The second language's NAME in the global _Languages table ("" = none).</summary>
        public string SecondLanguage = "";

        /// <summary>Optional glossary workbooks: the standard one and the project shadow.</summary>
        public string GlossaryPath1 = "";
        public string GlossaryPath2 = "";

        /// <summary>true = the glossary translation replaces the IO list's language-2 text;
        /// false = it only fills the rows whose language 2 is empty.</summary>
        public bool GlossaryTranslateAlways;

        /// <summary>Optional reference SWP for the formula-structure audit ("" = no audit).</summary>
        public string ReferenceSwpPath = "";

        /// <summary>Optional exported Alarms_*.txt files for the alarm-strings coherence.</summary>
        public List<string> AlarmStringsPaths = new List<string>();
    }

    /// <summary>One strobe-device alarm template (the legacy StrobeDevices sheet blocks).</summary>
    public sealed class SwpStrobeTemplate
    {
        public string TypeId = "";
        public string TypeDescription = "";
        public List<string> Mnemonics = new List<string>();      // with %sys%/%belt%/%cabinet%
        public List<string> Descriptions = new List<string>();   // parallel to Mnemonics
    }

    public sealed class SwpSidecarDocument
    {
        public string GeneratedUtc = "";
        public SwpBridgeSettings Settings = new SwpBridgeSettings();
        public List<SwpRowOverride> Overrides = new List<SwpRowOverride>();
        public List<SwpDictionaryEntry> Dictionary1 = new List<SwpDictionaryEntry>();
        public List<SwpDictionaryEntry> Dictionary2 = new List<SwpDictionaryEntry>();
        public List<SwpStrobeTemplate> StrobeTemplates = new List<SwpStrobeTemplate>();
        public List<SwpGlossaryTweak> GlossaryTweaks = new List<SwpGlossaryTweak>();
        public SwpQueue Queue = new SwpQueue();

        /// <summary>Load-time observations (skipped actions, tolerated oddities). Never serialized.</summary>
        public List<string> LoadNotes = new List<string>();
    }

    /// <summary>
    /// swp-bridge.json (de)serialization, following the sidecar house pattern: the schema token
    /// changes only for breaking shape changes, additive properties never bump it, the reader
    /// ignores unknown properties, an action with an unknown enum member is skipped into
    /// LoadNotes, and a foreign schema token refuses the load.
    /// </summary>
    public static class SwpSidecarJson
    {
        public const string SchemaToken = "swp-bridge-v1";

        public static string Serialize(SwpSidecarDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            SwpBridgeSettings s = document.Settings ?? new SwpBridgeSettings();

            var json = new StringBuilder(8192);
            json.Append('{');
            MiniJsonWriter.Property(json, "schema", SchemaToken, first: true);
            MiniJsonWriter.Property(json, "generatedUtc", document.GeneratedUtc, first: false);

            json.Append(',');
            MiniJsonWriter.AppendString(json, "settings");
            json.Append(":{");
            MiniJsonWriter.Property(json, "plantType", s.PlantType.ToString(), first: true);
            MiniJsonWriter.Property(json, "currentVersionFolder", s.CurrentVersionFolder, first: false);
            MiniJsonWriter.Property(json, "swpFileName", s.SwpFileName, first: false);
            MiniJsonWriter.Property(json, "configFileName", s.ConfigFileName, first: false);
            MiniJsonWriter.Property(json, "ioListPath", s.IoListPath, first: false);
            AppendStringArray(json, "cardSheets", s.CardSheets);
            MiniJsonWriter.Property(json, "plantInputsCsvPath", s.PlantInputsCsvPath, first: false);
            MiniJsonWriter.Property(json, "busfieldCsvPath", s.BusfieldCsvPath, first: false);
            MiniJsonWriter.Property(json, "destinationParamsCsvPath", s.DestinationParamsCsvPath, first: false);
            MiniJsonWriter.Property(json, "sensorCsvPath", s.SensorCsvPath, first: false);
            MiniJsonWriter.Property(json, "strobeDeviceCsvPath", s.StrobeDeviceCsvPath, first: false);
            MiniJsonWriter.NumberProperty(json, "firstLineId", s.FirstLineId, first: false);
            MiniJsonWriter.NumberProperty(json, "groupIndex", s.GroupIndex, first: false);
            MiniJsonWriter.Property(json, "includeRegex", s.IncludeRegex, first: false);
            MiniJsonWriter.Property(json, "excludeRegex", s.ExcludeRegex, first: false);
            MiniJsonWriter.BoolProperty(json, "trustConfigPlant", s.TrustConfigPlant, first: false);
            MiniJsonWriter.BoolProperty(json, "updateDescriptionFormulas", s.UpdateDescriptionFormulas, first: false);
            MiniJsonWriter.Property(json, "mismatchMode", s.MismatchMode.ToString(), first: false);
            MiniJsonWriter.BoolProperty(json, "applyToVersionFolder", s.ApplyToVersionFolder, first: false);
            MiniJsonWriter.Property(json, "secondLanguage", s.SecondLanguage, first: false);
            MiniJsonWriter.Property(json, "glossaryPath1", s.GlossaryPath1, first: false);
            MiniJsonWriter.Property(json, "glossaryPath2", s.GlossaryPath2, first: false);
            MiniJsonWriter.BoolProperty(json, "glossaryTranslateAlways", s.GlossaryTranslateAlways, first: false);
            MiniJsonWriter.Property(json, "referenceSwpPath", s.ReferenceSwpPath, first: false);
            AppendStringArray(json, "alarmStringsPaths", s.AlarmStringsPaths);
            json.Append('}');

            json.Append(',');
            MiniJsonWriter.AppendString(json, "overrides");
            json.Append(":[");
            for (int i = 0; i < document.Overrides.Count; i++)
            {
                SwpRowOverride o = document.Overrides[i];
                if (i > 0) json.Append(',');
                json.Append('{');
                MiniJsonWriter.Property(json, "mnemonic", o.Mnemonic, first: true);
                MiniJsonWriter.BoolProperty(json, "skipLabelCheck", o.SkipLabelCheck, first: false);
                if (o.SwpDbIndex.HasValue)
                    MiniJsonWriter.NumberProperty(json, "swpDbIndex", o.SwpDbIndex.Value, first: false);
                MiniJsonWriter.BoolProperty(json, "forceElectricalIds", o.ForceElectricalIds, first: false);
                MiniJsonWriter.Property(json, "desc1", o.Desc1Override, first: false);
                MiniJsonWriter.Property(json, "desc2", o.Desc2Override, first: false);
                json.Append('}');
            }
            json.Append(']');

            AppendDictionary(json, "dictionary1", document.Dictionary1);
            AppendDictionary(json, "dictionary2", document.Dictionary2);

            json.Append(',');
            MiniJsonWriter.AppendString(json, "glossaryTweaks");
            json.Append(":[");
            for (int i = 0; i < document.GlossaryTweaks.Count; i++)
            {
                SwpGlossaryTweak tweak = document.GlossaryTweaks[i];
                if (i > 0) json.Append(',');
                json.Append('{');
                MiniJsonWriter.Property(json, "en", tweak.En, first: true);
                MiniJsonWriter.BoolProperty(json, "enabled", tweak.Enabled, first: false);
                MiniJsonWriter.Property(json, "translation", tweak.Translation, first: false);
                json.Append('}');
            }
            json.Append(']');

            json.Append(',');
            MiniJsonWriter.AppendString(json, "strobeTemplates");
            json.Append(":[");
            for (int i = 0; i < document.StrobeTemplates.Count; i++)
            {
                SwpStrobeTemplate t = document.StrobeTemplates[i];
                if (i > 0) json.Append(',');
                json.Append('{');
                MiniJsonWriter.Property(json, "typeId", t.TypeId, first: true);
                MiniJsonWriter.Property(json, "typeDescription", t.TypeDescription, first: false);
                AppendStringArray(json, "mnemonics", t.Mnemonics);
                AppendStringArray(json, "descriptions", t.Descriptions);
                json.Append('}');
            }
            json.Append(']');

            json.Append(',');
            MiniJsonWriter.AppendString(json, "actions");
            json.Append(":[");
            for (int i = 0; i < document.Queue.Actions.Count; i++)
            {
                SwpAction a = document.Queue.Actions[i];
                if (i > 0) json.Append(',');
                json.Append('{');
                MiniJsonWriter.Property(json, "actionId", a.ActionId, first: true);
                MiniJsonWriter.Property(json, "sheet", a.Sheet, first: false);
                MiniJsonWriter.Property(json, "coordinate", a.CoordinateKey, first: false);
                MiniJsonWriter.Property(json, "field", a.Field.ToString(), first: false);
                MiniJsonWriter.Property(json, "mnemonic", a.Mnemonic, first: false);
                MiniJsonWriter.Property(json, "oldValue", a.OldValue, first: false);
                MiniJsonWriter.Property(json, "newValue", a.NewValue, first: false);
                MiniJsonWriter.BoolProperty(json, "include", a.Include, first: false);
                MiniJsonWriter.Property(json, "status", a.Status.ToString(), first: false);
                MiniJsonWriter.Property(json, "statusNote", a.StatusNote, first: false);
                MiniJsonWriter.Property(json, "queuedUtc", a.QueuedUtc, first: false);
                MiniJsonWriter.Property(json, "appliedUtc", a.AppliedUtc, first: false);
                json.Append('}');
            }
            json.Append(']');

            json.Append('}');
            return MiniJsonWriter.Indent(json.ToString());
        }

        public static bool TryParse(string json, out SwpSidecarDocument document, out string error)
        {
            document = null;
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

            var result = new SwpSidecarDocument { GeneratedUtc = root.Str("generatedUtc") };

            JsonValue settings = root.Get("settings");
            if (settings != null && settings.Kind == JsonKind.Object)
            {
                SwpBridgeSettings s = result.Settings;
                if (TryParseEnum(settings.Str("plantType"), out SwpPlantType plantType))
                    s.PlantType = plantType;
                else if (settings.Str("plantType").Length > 0)
                    result.LoadNotes.Add("Unknown plant type \"" + settings.Str("plantType") + "\" — Msc assumed.");
                s.CurrentVersionFolder = settings.Str("currentVersionFolder");
                s.SwpFileName = settings.Str("swpFileName");
                s.ConfigFileName = settings.Str("configFileName");
                s.IoListPath = settings.Str("ioListPath");
                foreach (JsonValue sheet in settings.Arr("cardSheets"))
                    if (sheet.Kind == JsonKind.String) s.CardSheets.Add(sheet.Text);
                s.PlantInputsCsvPath = settings.Str("plantInputsCsvPath");
                s.BusfieldCsvPath = settings.Str("busfieldCsvPath");
                s.DestinationParamsCsvPath = settings.Str("destinationParamsCsvPath");
                s.SensorCsvPath = settings.Str("sensorCsvPath");
                s.StrobeDeviceCsvPath = settings.Str("strobeDeviceCsvPath");
                s.FirstLineId = IntOf(settings, "firstLineId", 1);
                s.GroupIndex = IntOf(settings, "groupIndex", 1);
                s.IncludeRegex = settings.Str("includeRegex");
                s.ExcludeRegex = settings.Str("excludeRegex");
                s.TrustConfigPlant = settings.BoolOf("trustConfigPlant", false);
                s.UpdateDescriptionFormulas = settings.BoolOf("updateDescriptionFormulas", false);
                if (TryParseEnum(settings.Str("mismatchMode"), out SwpMismatchMode mismatchMode))
                    s.MismatchMode = mismatchMode;
                else
                    // Pre-mode sidecars carried three switches; the append flag maps cleanly.
                    s.MismatchMode = settings.BoolOf("mismatchAppendsLabel", true)
                        ? SwpMismatchMode.AppendToDescription
                        : SwpMismatchMode.Overwrite;
                s.ApplyToVersionFolder = settings.BoolOf("applyToVersionFolder", false);
                s.SecondLanguage = settings.Str("secondLanguage");
                s.GlossaryPath1 = settings.Str("glossaryPath1");
                s.GlossaryPath2 = settings.Str("glossaryPath2");
                s.GlossaryTranslateAlways = settings.BoolOf("glossaryTranslateAlways", false);
                s.ReferenceSwpPath = settings.Str("referenceSwpPath");
                foreach (JsonValue path in settings.Arr("alarmStringsPaths"))
                    if (path.Kind == JsonKind.String && path.Text.Length > 0)
                        s.AlarmStringsPaths.Add(path.Text);
            }

            foreach (JsonValue entry in root.Arr("overrides"))
            {
                if (entry.Kind != JsonKind.Object) continue;
                var o = new SwpRowOverride
                {
                    Mnemonic = entry.Str("mnemonic"),
                    SkipLabelCheck = entry.BoolOf("skipLabelCheck", false),
                    ForceElectricalIds = entry.BoolOf("forceElectricalIds", false),
                    Desc1Override = entry.Str("desc1"),
                    Desc2Override = entry.Str("desc2"),
                };
                JsonValue dbIndex = entry.Get("swpDbIndex");
                if (dbIndex != null && dbIndex.Kind == JsonKind.Number &&
                    int.TryParse(dbIndex.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
                    o.SwpDbIndex = index;
                if (o.Mnemonic.Length > 0) result.Overrides.Add(o);
            }

            ParseDictionary(root, "dictionary1", result.Dictionary1);
            ParseDictionary(root, "dictionary2", result.Dictionary2);

            foreach (JsonValue entry in root.Arr("glossaryTweaks"))
            {
                if (entry.Kind != JsonKind.Object) continue;
                string en = entry.Str("en");
                if (en.Length == 0) continue;
                result.GlossaryTweaks.Add(new SwpGlossaryTweak
                {
                    En = en,
                    Enabled = entry.BoolOf("enabled", true),
                    Translation = entry.Str("translation"),
                });
            }

            foreach (JsonValue entry in root.Arr("strobeTemplates"))
            {
                if (entry.Kind != JsonKind.Object) continue;
                var template = new SwpStrobeTemplate
                {
                    TypeId = entry.Str("typeId"),
                    TypeDescription = entry.Str("typeDescription"),
                };
                foreach (JsonValue m in entry.Arr("mnemonics"))
                    if (m.Kind == JsonKind.String) template.Mnemonics.Add(m.Text);
                foreach (JsonValue d in entry.Arr("descriptions"))
                    if (d.Kind == JsonKind.String) template.Descriptions.Add(d.Text);
                result.StrobeTemplates.Add(template);
            }

            int actionIndex = 0;
            foreach (JsonValue entry in root.Arr("actions"))
            {
                actionIndex++;
                if (entry.Kind != JsonKind.Object)
                {
                    result.LoadNotes.Add("Action " + actionIndex + " skipped: not an object.");
                    continue;
                }
                if (!TryParseEnum(entry.Str("field"), out SwpField field) ||
                    !TryParseEnum(entry.Str("status"), out SwpActionStatus status))
                {
                    result.LoadNotes.Add("Action " + actionIndex + " skipped: unknown field/status (\"" +
                        entry.Str("field") + "\"/\"" + entry.Str("status") + "\").");
                    continue;
                }

                result.Queue.Actions.Add(new SwpAction
                {
                    ActionId = FirstNonEmpty(entry.Str("actionId"), Guid.NewGuid().ToString("N")),
                    Sheet = entry.Str("sheet"),
                    CoordinateKey = entry.Str("coordinate"),
                    Field = field,
                    Mnemonic = entry.Str("mnemonic"),
                    OldValue = entry.Str("oldValue"),
                    NewValue = entry.Str("newValue"),
                    Include = entry.BoolOf("include", true),
                    Status = status,
                    StatusNote = entry.Str("statusNote"),
                    QueuedUtc = entry.Str("queuedUtc"),
                    AppliedUtc = entry.Str("appliedUtc"),
                });
            }

            document = result;
            error = null;
            return true;
        }

        private static void AppendStringArray(StringBuilder json, string name, List<string> values)
        {
            json.Append(',');
            MiniJsonWriter.AppendString(json, name);
            json.Append(":[");
            for (int i = 0; i < values.Count; i++)
            {
                if (i > 0) json.Append(',');
                MiniJsonWriter.AppendString(json, values[i] ?? "");
            }
            json.Append(']');
        }

        private static void AppendDictionary(StringBuilder json, string name, List<SwpDictionaryEntry> entries)
        {
            json.Append(',');
            MiniJsonWriter.AppendString(json, name);
            json.Append(":[");
            for (int i = 0; i < entries.Count; i++)
            {
                if (i > 0) json.Append(',');
                json.Append('{');
                MiniJsonWriter.Property(json, "find", entries[i].Find, first: true);
                MiniJsonWriter.Property(json, "replace", entries[i].Replace, first: false);
                json.Append('}');
            }
            json.Append(']');
        }

        private static void ParseDictionary(JsonValue root, string name, List<SwpDictionaryEntry> target)
        {
            foreach (JsonValue entry in root.Arr(name))
            {
                if (entry.Kind != JsonKind.Object) continue;
                string find = entry.Str("find");
                if (find.Length == 0) continue;
                target.Add(new SwpDictionaryEntry { Find = find, Replace = entry.Str("replace") });
            }
        }

        private static int IntOf(JsonValue obj, string name, int fallback)
        {
            JsonValue v = obj.Get(name);
            return v != null && v.Kind == JsonKind.Number &&
                   int.TryParse(v.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)
                ? result
                : fallback;
        }

        private static bool TryParseEnum<T>(string text, out T value) where T : struct
        {
            // Exact names only — a numeric string from a newer build must not round-trip as a
            // bogus member (the io-bridge sidecar lesson).
            value = default(T);
            if (string.IsNullOrEmpty(text) || !char.IsLetter(text[0])) return false;
            return Enum.TryParse(text, ignoreCase: false, out value) && Enum.IsDefined(typeof(T), value);
        }

        private static string FirstNonEmpty(string preferred, string fallback) =>
            string.IsNullOrEmpty(preferred) ? fallback : preferred;
    }
}
