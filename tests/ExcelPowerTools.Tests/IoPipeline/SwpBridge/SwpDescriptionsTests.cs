using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpDescriptionsTests
    {
        private static DiagnosticEntry Entry(string d1a = "", string d1b = "", string d2a = "",
            string d2b = "", string pnName = "", string pnIp = "")
        {
            return new DiagnosticEntry
            {
                Mnemonic = "CABINETAL00100",
                Desc1A = d1a,
                Desc1B = d1b,
                Desc2A = d2a,
                Desc2B = d2b,
                ProfinetName = pnName,
                ProfinetIp = pnIp,
            };
        }

        private static List<SwpDictionaryEntry> Dict(params string[] pairs)
        {
            var dict = new List<SwpDictionaryEntry>();
            for (int i = 0; i < pairs.Length; i += 2)
                dict.Add(new SwpDictionaryEntry { Find = pairs[i], Replace = pairs[i + 1] });
            return dict;
        }

        private static SwpDescriptions.Result Compose(DiagnosticEntry entry,
            List<SwpDictionaryEntry> d1 = null, List<SwpDictionaryEntry> d2 = null,
            string lang1 = "FIELD BUS MODULE FAILURE", string lang2 = "",
            SwpRowOverride overrides = null)
        {
            return SwpDescriptions.Compose(entry, d1, d2, lang1, lang2, overrides);
        }

        [Fact]
        public void Concatenates_permanent_and_variable_parts_verbatim()
        {
            var r = Compose(Entry("CIRCUIT BREAKER", "COOLING FAN", "DISJONCTEUR", "VENTILATEUR"));
            Assert.Equal("CIRCUIT BREAKER COOLING FAN", r.Desc1);
            Assert.Equal("DISJONCTEUR VENTILATEUR", r.Desc2);
        }

        [Fact]
        public void Empty_variable_part_keeps_the_legacy_trailing_space()
        {
            Assert.Equal("CIRCUIT BREAKER ", Compose(Entry("CIRCUIT BREAKER")).Desc1);
        }

        [Fact]
        public void Dictionary_replaces_sequentially_in_order()
        {
            var dict = Dict("CIRCUIT BREAKER", "CB", "CB TRIPPED", "CB-T");
            // First entry turns the text into "CB TRIPPED X", the second then matches it.
            Assert.Equal("CB-T X", Compose(Entry("CIRCUIT BREAKER TRIPPED", "X"), dict).Desc1);
        }

        [Fact]
        public void Empty_find_rows_are_skipped()
        {
            Assert.Equal("A", SwpDescriptions.MultiReplace("A", Dict("", "ZAP")));
        }

        [Fact]
        public void A_bus_module_composes_the_failure_constant_with_its_identity()
        {
            var r = Compose(Entry("SOME MODULE", "", "QUALCOSA", "",
                pnName: "n001-panel", pnIp: "192.168.0.10"), lang2: "AVARIA DO MÓDULO BUS DE CAMPO");
            Assert.Equal("FIELD BUS MODULE FAILURE [n001-panel - 192.168.0.10]", r.Desc1);
            Assert.Equal("AVARIA DO MÓDULO BUS DE CAMPO [n001-panel - 192.168.0.10]", r.Desc2);
        }

        [Fact]
        public void A_missing_second_language_constant_falls_back_to_language_one()
        {
            var r = Compose(Entry("SOME MODULE", "", "", "", pnName: "n001-x", pnIp: "10.0.0.1"));
            Assert.Equal("FIELD BUS MODULE FAILURE [n001-x - 10.0.0.1]", r.Desc2);
        }

        [Fact]
        public void An_empty_language_one_constant_uses_the_built_in_default()
        {
            var r = Compose(Entry("M", "", "", "", pnName: "n001-x", pnIp: "1.2.3.4"), lang1: "");
            Assert.StartsWith("FIELD BUS MODULE FAILURE [", r.Desc1);
        }

        [Fact]
        public void Asi_master_rows_compose_like_any_fieldbus_module()
        {
            // The legacy blanked rows whose description says "MODULE AS-I" — in the field that
            // proposed ERASING a correct FIELD BUS MODULE FAILURE text on an AS-I gateway.
            var r = Compose(Entry("MODULE AS-I", "", "", "",
                pnName: "n115", pnIp: "192.168.1.115"), lang2: "MSG");
            Assert.Equal("FIELD BUS MODULE FAILURE [n115 - 192.168.1.115]", r.Desc1);
            Assert.Equal("MSG [n115 - 192.168.1.115]", r.Desc2);
        }

        // ---- the glossary as the bus-module + T.S. Ref authority ----------------------------

        private static SwpGlossaryPair Pair(string en, string translation, string tsRef = "") =>
            new SwpGlossaryPair { En = en, Translation = translation, TsRef = tsRef };

        private static SwpDescriptions.Context GlossaryContext(params SwpGlossaryPair[] pairs) =>
            new SwpDescriptions.Context
            {
                BusFailureLang1 = "GENERIC CONSTANT 1",
                BusFailureLang2 = "GENERIC CONSTANT 2",
                GlossaryPairs = new List<SwpGlossaryPair>(pairs),
            };

        [Fact]
        public void A_bus_module_keeps_its_description_when_it_is_an_official_phrase()
        {
            SwpDescriptions.Context ctx = GlossaryContext(
                Pair("FIELD BUS FAILURE - DRIVE", "AVARIA DO BUS - ACIONAMENTO", "TS_040"),
                Pair("FIELD BUS MODULE FAILURE", "AVARIA DO MÓDULO BUS DE CAMPO", "TS_041"));

            var r = SwpDescriptions.Compose(
                Entry("FIELD BUS FAILURE - DRIVE", "", "", "", pnName: "n0015-MC1", pnIp: "10.0.0.5"), ctx);

            Assert.Equal("FIELD BUS FAILURE - DRIVE [n0015-MC1 - 10.0.0.5]", r.Desc1);
            Assert.Equal("AVARIA DO BUS - ACIONAMENTO [n0015-MC1 - 10.0.0.5]", r.Desc2);
            Assert.False(r.BusDescMismatch);
            Assert.Equal("TS_040", r.TsRef);
        }

        [Fact]
        public void An_unofficial_bus_description_falls_to_the_dedicated_glossary_row()
        {
            SwpDescriptions.Context ctx = GlossaryContext(
                Pair("FIELD BUS MODULE FAILURE", "AVARIA DO MÓDULO BUS DE CAMPO", "TS_041"));

            var r = SwpDescriptions.Compose(
                Entry("SOME HOMEMADE TEXT", "", "", "", pnName: "n002-x", pnIp: "10.0.0.9"), ctx);

            // The glossary row serves BOTH languages — never the _Languages constants.
            Assert.Equal("FIELD BUS MODULE FAILURE [n002-x - 10.0.0.9]", r.Desc1);
            Assert.Equal("AVARIA DO MÓDULO BUS DE CAMPO [n002-x - 10.0.0.9]", r.Desc2);
            Assert.True(r.BusDescMismatch);
            Assert.Equal("SOME HOMEMADE TEXT", r.BusIoText);
            Assert.Equal("FIELD BUS MODULE FAILURE", r.BusChosenEn);
            Assert.Equal("TS_041", r.TsRef);
        }

        [Fact]
        public void A_glossary_without_the_dedicated_row_leaves_the_constants_in_charge()
        {
            SwpDescriptions.Context ctx = GlossaryContext(Pair("UNRELATED", "SEM RELAÇÃO"));

            var r = SwpDescriptions.Compose(
                Entry("SOMETHING", "", "", "", pnName: "n003-y", pnIp: "10.0.0.2"), ctx);
            Assert.Equal("GENERIC CONSTANT 1 [n003-y - 10.0.0.2]", r.Desc1);
            Assert.Equal("GENERIC CONSTANT 2 [n003-y - 10.0.0.2]", r.Desc2);
            Assert.False(r.BusDescMismatch);
        }

        [Fact]
        public void The_glossary_card_wins_over_a_differing_io_list_ts_ref()
        {
            SwpDescriptions.Context ctx = GlossaryContext(Pair("STATE OF EMERGENCY ", "ESTADO DE EMERGÊNCIA", "TS_001"));
            DiagnosticEntry entry = Entry("STATE OF EMERGENCY");
            entry.TsRef = "TS_999";

            var r = SwpDescriptions.Compose(entry, ctx);
            Assert.Equal("TS_001", r.TsRef);
            Assert.True(r.TsRefOverridden);
            Assert.Equal("TS_999", r.TsRefFromIo);
        }

        [Fact]
        public void An_empty_io_ts_ref_takes_the_glossary_card_silently()
        {
            SwpDescriptions.Context ctx = GlossaryContext(Pair("STATE OF EMERGENCY", "ESTADO DE EMERGÊNCIA", "TS_001"));

            var r = SwpDescriptions.Compose(Entry("STATE OF EMERGENCY"), ctx);
            Assert.Equal("TS_001", r.TsRef);
            Assert.False(r.TsRefOverridden);
        }

        [Fact]
        public void A_description_split_across_both_parts_still_matches_the_glossary()
        {
            // Field hygiene: a fully glossary-compliant text split into permanent + variable
            // parts. The FULL trimmed description matches → consumed whole: the translation
            // stands alone and the card comes with it.
            SwpDescriptions.Context ctx = GlossaryContext(
                Pair("EMERGENCY Circuit De-Energized (CUT) by SAFETY PLC",
                     "CIRCUITO DE EMERGÊNCIA DESENERGIZADO", "TS_012"));

            var r = SwpDescriptions.Compose(
                Entry("EMERGENCY Circuit", "De-Energized (CUT) by SAFETY PLC"), ctx);
            Assert.Equal("CIRCUITO DE EMERGÊNCIA DESENERGIZADO", r.Desc2);
            Assert.Equal("TS_012", r.TsRef);
        }

        [Fact]
        public void The_permanent_part_tier_wins_over_the_full_description_tier()
        {
            SwpDescriptions.Context ctx = GlossaryContext(
                Pair("RELAY FAULT", "AVARIA DO RELÉ"),
                Pair("RELAY FAULT +K1", "NUNCA USADO"));

            var r = SwpDescriptions.Compose(Entry("RELAY FAULT", "+K1"), ctx);
            Assert.Equal("AVARIA DO RELÉ +K1", r.Desc2);   // permanent match, variable rides
        }

        [Fact]
        public void A_bus_module_description_split_across_both_parts_matches_too()
        {
            SwpDescriptions.Context ctx = GlossaryContext(
                Pair("FIELD BUS FAILURE - DRIVE", "AVARIA DO BUS - ACIONAMENTO"));

            var r = SwpDescriptions.Compose(
                Entry("FIELD BUS FAILURE -", "DRIVE", pnName: "n007-x", pnIp: "10.0.0.7"), ctx);
            Assert.Equal("FIELD BUS FAILURE - DRIVE [n007-x - 10.0.0.7]", r.Desc1);
            Assert.False(r.BusDescMismatch);
        }

        [Fact]
        public void An_unofficial_phrase_keeps_the_io_list_ts_ref()
        {
            SwpDescriptions.Context ctx = GlossaryContext(Pair("STATE OF EMERGENCY", "ESTADO DE EMERGÊNCIA", "TS_001"));
            DiagnosticEntry entry = Entry("CIRCUIT BREAKER TRIPPED");
            entry.TsRef = "TS_777";

            var r = SwpDescriptions.Compose(entry, ctx);
            Assert.Equal("TS_777", r.TsRef);
            Assert.False(r.TsRefOverridden);
        }

        [Fact]
        public void Blank_language_two_falls_back_to_language_one()
        {
            Assert.Equal("ENGLISH TEXT PART", Compose(Entry("ENGLISH TEXT", "PART")).Desc2);
        }

        [Fact]
        public void Spaces_only_language_two_also_falls_back()
        {
            Assert.Equal("EN X", Compose(Entry("EN", "X", "  ", " ")).Desc2);
        }

        [Fact]
        public void Overrides_win_over_everything()
        {
            var overrides = new SwpRowOverride { Desc1Override = "MY TEXT", Desc2Override = "MON TEXTE" };
            var r = Compose(Entry("IGNORED", "", "IGNORÉ", "",
                pnName: "n001-x", pnIp: "1.2.3.4"), lang2: "MSG", overrides: overrides);
            Assert.Equal("MY TEXT", r.Desc1);
            Assert.Equal("MON TEXTE", r.Desc2);
        }

        [Fact]
        public void Language_two_dictionary_is_independent()
        {
            var r = Compose(Entry("ALPHA", "X", "BETA", "Y"), Dict("ALPHA", "A1"), Dict("BETA", "B2"));
            Assert.Equal("A1 X", r.Desc1);
            Assert.Equal("B2 Y", r.Desc2);
        }
    }
}
