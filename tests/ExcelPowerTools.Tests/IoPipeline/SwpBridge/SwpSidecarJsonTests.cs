using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpSidecarJsonTests
    {
        private static SwpSidecarDocument FullDocument()
        {
            var document = new SwpSidecarDocument { GeneratedUtc = "2026-08-20T10:00:00.0000000Z" };
            document.Settings = new SwpBridgeSettings
            {
                PlantType = SwpPlantType.Ups,
                CurrentVersionFolder = "06_20-Aug-2026",
                SwpFileName = "SWP_01_Sorter_v2.xlsx",
                ConfigFileName = "CONFIG_PLANT_01_Sorter.xlsx",
                IoListPath = @"C:\io\8FVZ.xlsx",
                CardSheets = { "R1.S1.C1", "R1.S1.C2" },
                PlantInputsCsvPath = @"C:\csv\Plant_Inputs.csv",
                BusfieldCsvPath = @"C:\csv\BusField.csv",
                DestinationParamsCsvPath = @"C:\csv\Destination_Parameters.csv",
                SensorCsvPath = @"C:\csv\Sensor.csv",
                StrobeDeviceCsvPath = @"C:\csv\StrobeDevice.csv",
                FirstLineId = 3,
                GroupIndex = 2,
                IncludeRegex = "^MYDIAG",
                ExcludeRegex = "TEST$",
                ConfigMode = SwpConfigMode.KeepConfigText,
                ConfigSheetRulePattern = "BELT$",
                ConfigSheetRuleMode = SwpConfigMode.ReplaceTemplateTail,
                SwpRootFolder = @"C:\systems\sorter",
                MismatchMode = SwpMismatchMode.Overwrite,
                ApplyToVersionFolder = true,
                VersionFolderCopyPattern = "G1",
                SecondLanguage = "Português",
            };
            document.Overrides.Add(new SwpRowOverride
            {
                Mnemonic = "CABINETAL00100",
                SkipLabelCheck = true,
                SwpDbIndex = 7,
                ForceElectricalIds = true,
                Desc1Override = "MY \"quoted\" text",
                Desc2Override = "±°",
            });
            document.Overrides.Add(new SwpRowOverride { Mnemonic = "ALLB_1_001_50" });
            document.Dictionary1.Add(new SwpDictionaryEntry { Find = "CIRCUIT BREAKER", Replace = "CB" });
            document.Dictionary2.Add(new SwpDictionaryEntry { Find = "DISJONCTEUR", Replace = "DJ" });
            document.StrobeTemplates.Add(new SwpStrobeTemplate
            {
                TypeId = "1",
                TypeDescription = "Scale",
                Mnemonics = { "WARQ_%sys%_%cabinet%_01" },
                Descriptions = { "SCALE CERTIFICATION" },
            });
            document.Queue.Actions.Add(new SwpAction
            {
                ActionId = "a1",
                Sheet = "CABINET",
                CoordinateKey = "DB_CABINETS_DIAG.Alarm1>1>0",
                Field = SwpField.AddressG,
                Mnemonic = "CABINETAL00100",
                OldValue = "",
                NewValue = "80.3.1",
                Include = false,
                Status = SwpActionStatus.ConflictStale,
                StatusNote = "note",
                QueuedUtc = "t0",
                AppliedUtc = "",
                FldGroup = "CABINET|DB_CABINETS_DIAG.Alarm1>1>0",
            });
            return document;
        }

        [Fact]
        public void A_fresh_session_keeps_everything_but_the_history()
        {
            // "Save as new…": a second SWP driven by the same IO list starts from the first session's
            // settings and knowledge, never from its decisions — the two cards share their mnemonics.
            SwpSidecarDocument source = FullDocument();
            source.GlossaryTweaks.Add(new SwpGlossaryTweak
            {
                En = "CIRCUIT BREAKER TRIPPED",
                Translation = "DISJONCTEUR DÉCLENCHÉ",
                Enabled = false,
            });

            SwpSidecarDocument fresh = SwpSidecarJson.FreshSession(source);

            Assert.Empty(fresh.Queue.Actions);
            Assert.Empty(fresh.Overrides);
            Assert.Empty(fresh.LoadNotes);
            Assert.Equal(@"C:\io\8FVZ.xlsx", fresh.Settings.IoListPath);
            Assert.Equal(new List<string> { "R1.S1.C1", "R1.S1.C2" }, fresh.Settings.CardSheets);
            Assert.Equal("^MYDIAG", fresh.Settings.IncludeRegex);
            Assert.Equal(SwpConfigMode.KeepConfigText, fresh.Settings.ConfigMode);
            Assert.Equal("BELT$", fresh.Settings.ConfigSheetRulePattern);
            Assert.Equal(@"C:\systems\sorter", fresh.Settings.SwpRootFolder);
            Assert.Equal("G1", fresh.Settings.VersionFolderCopyPattern);
            Assert.Equal("CB", Assert.Single(fresh.Dictionary1).Replace);
            Assert.Equal("DJ", Assert.Single(fresh.Dictionary2).Replace);
            Assert.Equal("Scale", Assert.Single(fresh.StrobeTemplates).TypeDescription);
            SwpGlossaryTweak tweak = Assert.Single(fresh.GlossaryTweaks);
            Assert.Equal("DISJONCTEUR DÉCLENCHÉ", tweak.Translation);
            Assert.False(tweak.Enabled);

            // …and it shares nothing with the source: editing one never reaches the other.
            fresh.Settings.IncludeRegex = "CHANGED";
            fresh.Dictionary1[0].Replace = "CHANGED";
            Assert.Equal("^MYDIAG", source.Settings.IncludeRegex);
            Assert.Equal("CB", source.Dictionary1[0].Replace);
            Assert.Single(source.Queue.Actions);
            Assert.Equal(2, source.Overrides.Count);

            Assert.Throws<System.ArgumentNullException>(() => SwpSidecarJson.FreshSession(null));
        }

        [Fact]
        public void Round_trips_everything()
        {
            string json = SwpSidecarJson.Serialize(FullDocument());
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out string error), error);

            SwpBridgeSettings s = parsed.Settings;
            Assert.Equal(SwpPlantType.Ups, s.PlantType);
            Assert.Equal("06_20-Aug-2026", s.CurrentVersionFolder);
            Assert.Equal("SWP_01_Sorter_v2.xlsx", s.SwpFileName);
            Assert.Equal("CONFIG_PLANT_01_Sorter.xlsx", s.ConfigFileName);
            Assert.Equal(@"C:\io\8FVZ.xlsx", s.IoListPath);
            Assert.Equal(new List<string> { "R1.S1.C1", "R1.S1.C2" }, s.CardSheets);
            Assert.Equal(3, s.FirstLineId);
            Assert.Equal(2, s.GroupIndex);
            Assert.Equal("^MYDIAG", s.IncludeRegex);
            Assert.Equal(SwpConfigMode.KeepConfigText, s.ConfigMode);
            Assert.Equal("BELT$", s.ConfigSheetRulePattern);
            Assert.Equal(SwpConfigMode.ReplaceTemplateTail, s.ConfigSheetRuleMode);
            Assert.Equal(@"C:\systems\sorter", s.SwpRootFolder);
            // The superseded switches are still written, from the mode, so an older build reading
            // this file finds the setting it knows instead of a silent default.
            Assert.True(s.TrustConfigPlant);
            Assert.False(s.UpdateDescriptionFormulas);
            Assert.Equal(SwpMismatchMode.Overwrite, s.MismatchMode);
            Assert.True(s.ApplyToVersionFolder);
            Assert.Equal("G1", s.VersionFolderCopyPattern);
            Assert.Equal("Português", s.SecondLanguage);

            Assert.Equal(2, parsed.Overrides.Count);
            SwpRowOverride o = parsed.Overrides[0];
            Assert.Equal("CABINETAL00100", o.Mnemonic);
            Assert.True(o.SkipLabelCheck);
            Assert.Equal(7, o.SwpDbIndex);
            Assert.True(o.ForceElectricalIds);
            Assert.Equal("MY \"quoted\" text", o.Desc1Override);
            Assert.Equal("±°", o.Desc2Override);
            Assert.Null(parsed.Overrides[1].SwpDbIndex);

            Assert.Equal("CB", parsed.Dictionary1[0].Replace);
            Assert.Equal("DJ", parsed.Dictionary2[0].Replace);
            Assert.Equal("Scale", parsed.StrobeTemplates[0].TypeDescription);
            Assert.Equal("WARQ_%sys%_%cabinet%_01", parsed.StrobeTemplates[0].Mnemonics[0]);

            SwpAction a = Assert.Single(parsed.Queue.Actions);
            Assert.Equal("a1", a.ActionId);
            Assert.Equal(SwpField.AddressG, a.Field);
            Assert.Equal(SwpActionStatus.ConflictStale, a.Status);
            Assert.False(a.Include);
            Assert.Equal("CABINET|DB_CABINETS_DIAG.Alarm1>1>0", a.FldGroup);
            Assert.Empty(parsed.LoadNotes);
        }

        [Fact]
        public void A_foreign_schema_token_refuses_the_load()
        {
            Assert.False(SwpSidecarJson.TryParse("{\"schema\":\"swp-bridge-v99\"}", out _, out string error));
            Assert.Contains("swp-bridge-v99", error);
        }

        [Fact]
        public void A_missing_schema_token_refuses_the_load()
        {
            Assert.False(SwpSidecarJson.TryParse("{}", out _, out string error));
            Assert.Contains("no schema token", error);
        }

        [Fact]
        public void Unknown_action_enums_skip_into_load_notes()
        {
            string json =
                "{\"schema\":\"swp-bridge-v1\",\"actions\":[" +
                "{\"actionId\":\"a1\",\"field\":\"HologramHH\",\"status\":\"Pending\"}," +
                "{\"actionId\":\"a2\",\"field\":\"AddressG\",\"status\":\"Pending\"}]}";
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out string error), error);

            SwpAction kept = Assert.Single(parsed.Queue.Actions);
            Assert.Equal("a2", kept.ActionId);
            Assert.Contains(parsed.LoadNotes, n => n.Contains("HologramHH"));
        }

        [Fact]
        public void Numeric_enum_aliases_are_rejected()
        {
            string json =
                "{\"schema\":\"swp-bridge-v1\",\"actions\":[" +
                "{\"actionId\":\"a1\",\"field\":\"3\",\"status\":\"Pending\"}]}";
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out _));
            Assert.Empty(parsed.Queue.Actions);
            Assert.Single(parsed.LoadNotes);
        }

        [Fact]
        public void Unknown_properties_are_ignored()
        {
            string json =
                "{\"schema\":\"swp-bridge-v1\",\"futureThing\":{\"a\":1}," +
                "\"settings\":{\"plantType\":\"Msc\",\"laterAddition\":true}}";
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out string error), error);
            Assert.Equal(SwpPlantType.Msc, parsed.Settings.PlantType);
        }

        [Theory]
        // The old switches were independent; the exclusive mode reads them most-specific-first.
        [InlineData("\"trustConfigPlant\":true,\"updateDescriptionFormulas\":false", SwpConfigMode.KeepConfigText, false)]
        [InlineData("\"trustConfigPlant\":true,\"updateDescriptionFormulas\":true", SwpConfigMode.KeepConfigText, false)]
        [InlineData("\"trustConfigPlant\":false,\"updateDescriptionFormulas\":true", SwpConfigMode.ReplaceTemplateTail, false)]
        // Neither set meant the description was stacked AFTER the template's own tail — the shape
        // the mode exists to end. Those land on the rebuild, and are told so rather than moved
        // silently.
        [InlineData("\"trustConfigPlant\":false,\"updateDescriptionFormulas\":false", SwpConfigMode.ReplaceTemplateTail, true)]
        public void A_sidecar_predating_the_config_mode_migrates_from_the_old_switches(
            string switches, SwpConfigMode expected, bool expectNote)
        {
            string json = "{\"schema\":\"swp-bridge-v1\",\"settings\":{\"plantType\":\"Msc\"," + switches + "}}";

            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out string error), error);

            Assert.Equal(expected, parsed.Settings.ConfigMode);
            Assert.Equal(expectNote, parsed.LoadNotes.Exists(n => n.Contains("predates the CONFIG_PLANT mode")));
        }

        [Fact]
        public void An_absent_swp_root_means_the_sidecars_own_folder()
        {
            // Every sidecar written before named sessions carries no root, and must stay
            // self-locating — a system folder moved wholesale still resolves.
            string json = "{\"schema\":\"swp-bridge-v1\",\"settings\":{\"plantType\":\"Msc\"}}";
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out _));
            Assert.Equal("", parsed.Settings.SwpRootFolder);
        }

        [Fact]
        public void Include_defaults_to_true_when_absent()
        {
            string json =
                "{\"schema\":\"swp-bridge-v1\",\"actions\":[" +
                "{\"actionId\":\"a1\",\"field\":\"ExportB\",\"status\":\"Pending\"}]}";
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out _));
            Assert.True(parsed.Queue.Actions[0].Include);
        }

        [Fact]
        public void An_unknown_plant_type_falls_back_to_msc_with_a_note()
        {
            string json =
                "{\"schema\":\"swp-bridge-v1\",\"settings\":{\"plantType\":\"Hyperloop\"}}";
            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument parsed, out _));
            Assert.Equal(SwpPlantType.Msc, parsed.Settings.PlantType);
            Assert.Contains(parsed.LoadNotes, n => n.Contains("Hyperloop"));
        }
    }
}
