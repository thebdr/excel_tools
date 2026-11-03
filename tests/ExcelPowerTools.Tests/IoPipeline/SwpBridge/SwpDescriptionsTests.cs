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
            var r = Compose(Entry("ENGLISH TEXT", "PART"));
            Assert.Equal("ENGLISH TEXT PART", r.Desc2);
            // Language-1 text in the language-2 cell is flagged: a single-language SCADA shows it.
            Assert.True(r.Desc2FromLanguage1);
        }

        [Fact]
        public void Spaces_only_language_two_also_falls_back()
        {
            var r = Compose(Entry("EN", "X", "  ", " "));
            Assert.Equal("EN X", r.Desc2);
            Assert.True(r.Desc2FromLanguage1);
        }

        [Fact]
        public void A_real_language_two_is_not_flagged()
        {
            var r = Compose(Entry("CIRCUIT BREAKER", "COOLING FAN", "DISJONCTEUR", "VENTILATEUR"));
            Assert.False(r.Desc2FromLanguage1);
            Assert.False(r.Desc2VariablePartFromLanguage1);
        }

        // ---- the variable part beside a glossary translation --------------------------------

        private static SwpDescriptions.Context FrenchGlossary(bool translateAlways) =>
            new SwpDescriptions.Context
            {
                GlossaryPairs = new List<SwpGlossaryPair>
                {
                    Pair("CIRCUIT BREAKER TRIPPED", "DISJONCTEUR MAGNÉTOTHERMIQUE DÉCLENCHÉ", "TS_020"),
                },
                GlossaryTranslateAlways = translateAlways,
            };

        [Fact]
        public void A_permanent_part_match_takes_the_io_lists_language_two_variable_part()
        {
            // FVZ induction, card R1.S2.C1 row 9: the glossary translates the permanent part and the
            // IO list translates the variable one — a French-only SCADA must read both in French.
            var r = SwpDescriptions.Compose(Entry(
                "CIRCUIT BREAKER TRIPPED", "24V IM 155 AND I/O MODULE",
                "DISJONCTEUR DÉCLENCHÉ", "24V IM 155 ET MODULES I/O"), FrenchGlossary(translateAlways: true));

            Assert.Equal("DISJONCTEUR MAGNÉTOTHERMIQUE DÉCLENCHÉ 24V IM 155 ET MODULES I/O", r.Desc2);
            Assert.False(r.Desc2VariablePartFromLanguage1);
            Assert.False(r.Desc2FromLanguage1);
        }

        [Fact]
        public void Without_a_language_two_variable_part_language_ones_rides_along_and_is_flagged()
        {
            var r = SwpDescriptions.Compose(Entry(
                "CIRCUIT BREAKER TRIPPED", "24V IM 155 AND I/O MODULE", "", ""), FrenchGlossary(translateAlways: false));

            Assert.Equal("DISJONCTEUR MAGNÉTOTHERMIQUE DÉCLENCHÉ 24V IM 155 AND I/O MODULE", r.Desc2);
            Assert.True(r.Desc2VariablePartFromLanguage1);
        }

        [Fact]
        public void The_language_two_dictionary_shortens_the_variable_part_it_contributes()
        {
            var r = SwpDescriptions.Compose(Entry(
                    "CIRCUIT BREAKER TRIPPED", "400V INDUCTION LINE 1", "DISJONCTEUR DÉCLENCHÉ", "400V LIGNE D'INDUCTION 1"),
                new SwpDescriptions.Context
                {
                    Dictionary2 = Dict("LIGNE D'INDUCTION", "L.IND."),
                    GlossaryPairs = FrenchGlossary(true).GlossaryPairs,
                    GlossaryTranslateAlways = true,
                });

            Assert.Equal("DISJONCTEUR MAGNÉTOTHERMIQUE DÉCLENCHÉ 400V L.IND. 1", r.Desc2);
        }

        private static SwpDescriptions.Context FvtSafetyContext(List<SwpDictionaryEntry> d1, List<SwpDictionaryEntry> d2) =>
            new SwpDescriptions.Context
            {
                Dictionary1 = d1,
                Dictionary2 = d2,
                GlossaryPairs = new List<SwpGlossaryPair>
                {
                    Pair("EMERGENCY PUSH BUTTON PRESSED", "ARRÊT D'URGENCE AVEC BOUTON COUP DE POING", "TS_002"),
                    Pair("SWITCH DISCONNECTOR IS OPEN", "LE SECTIONNEUR EST OUVERT", "TS_004"),
                },
                GlossaryTranslateAlways = true,
            };

        /// <summary>
        /// FVT safety list, 2026-09-22 (CABINETAL00427, row 2364): both variable parts read
        /// "CH1/CH2" and the _Dictionary deletes the channel markers on both sides. The language-2
        /// part came out empty and the code took that for "the IO list has no language-2 part",
        /// riding language 1's "CH1/CH2" in RAW — the one path with no dictionary at all — so
        /// Description 2 read "… COUP DE POING CH1/CH2" while Description 1 was clean. A part the
        /// IO list WROTE counts as present whatever the dictionary leaves of it.
        /// </summary>
        [Fact]
        public void A_language_two_variable_part_the_dictionary_empties_never_lets_language_ones_ride_raw()
        {
            List<SwpDictionaryEntry> channels = Dict("CH1/", "", "CH1\\", "", "CH1", "", "CH2", "");
            var r = SwpDescriptions.Compose(
                Entry("EMERGENCY PUSH BUTTON PRESSED", "CH1/CH2", "ARRÊT D’URGENCE AVEC BOUTON COUP DE POING", "CH1/CH2"),
                FvtSafetyContext(channels, channels));

            Assert.Equal("EMERGENCY PUSH BUTTON PRESSED ", r.Desc1);
            Assert.Equal("ARRÊT D'URGENCE AVEC BOUTON COUP DE POING ", r.Desc2);
            Assert.False(r.Desc2VariablePartFromLanguage1);
            Assert.Equal("", r.Desc2RidingText);

            // The 188 sibling rows with a device after the markers keep it, on both sides.
            r = SwpDescriptions.Compose(
                Entry("EMERGENCY PUSH BUTTON PRESSED", "CH1\\CH2 =ES-0035", "ARRÊT D’URGENCE AVEC BOUTON COUP DE POING", " CH1\\CH2=ES-0035"),
                FvtSafetyContext(channels, channels));
            Assert.Equal("EMERGENCY PUSH BUTTON PRESSED  =ES-0035", r.Desc1);
            Assert.Equal("ARRÊT D'URGENCE AVEC BOUTON COUP DE POING  =ES-0035", r.Desc2);
        }

        [Fact]
        public void Language_ones_riding_variable_part_passes_dictionary_one_then_dictionary_two()
        {
            // No language-2 variable part at all: language 1's rides — through ITS dictionary…
            var r = SwpDescriptions.Compose(
                Entry("SWITCH DISCONNECTOR IS OPEN", "CH1/CH2 =ES-2200", "", ""),
                FvtSafetyContext(Dict("CH1/", "", "CH2", ""), null));
            Assert.Equal("LE SECTIONNEUR EST OUVERT  =ES-2200", r.Desc2);
            Assert.True(r.Desc2VariablePartFromLanguage1);
            Assert.Equal("=ES-2200", r.Desc2RidingText);

            // …and through language 2's too, because it lands in Description 2.
            r = SwpDescriptions.Compose(
                Entry("SWITCH DISCONNECTOR IS OPEN", "CH1/CH2 =ES-2200", "", ""),
                FvtSafetyContext(null, Dict("CH1/", "", "CH2", "")));
            Assert.Equal("LE SECTIONNEUR EST OUVERT  =ES-2200", r.Desc2);
            Assert.Equal("=ES-2200", r.Desc2RidingText);

            // Rules that delete the whole part leave nothing to ride — and nothing to flag.
            r = SwpDescriptions.Compose(
                Entry("SWITCH DISCONNECTOR IS OPEN", "CH1/CH2", "", ""),
                FvtSafetyContext(Dict("CH1/", "", "CH2", ""), null));
            Assert.Equal("LE SECTIONNEUR EST OUVERT ", r.Desc2);
            Assert.False(r.Desc2VariablePartFromLanguage1);
        }

        [Fact]
        public void The_whole_description_fallback_to_language_one_passes_dictionary_two_except_on_bus_modules()
        {
            // No glossary, no language 2 at all: Description 2 repeats Description 1, cleaned by
            // the language-2 rules as well — it is what the single-language SCADA shows.
            var r = Compose(Entry("EMERGENCY PUSH BUTTON PRESSED", "CH1/CH2 =ES-0035"),
                d1: null, d2: Dict("CH1/CH2 ", ""));
            Assert.Equal("EMERGENCY PUSH BUTTON PRESSED CH1/CH2 =ES-0035", r.Desc1);
            Assert.Equal("EMERGENCY PUSH BUTTON PRESSED =ES-0035", r.Desc2);
            Assert.True(r.Desc2FromLanguage1);

            // A bus module's phrase + [name - ip] is identity: no dictionary touches it.
            r = Compose(Entry("PROFINET SWITCH", pnName: "n0015-tric-aec01-K28001", pnIp: "192.168.20.15"),
                d1: null, d2: Dict("K28001", "X"));
            Assert.Equal("FIELD BUS MODULE FAILURE [n0015-tric-aec01-K28001 - 192.168.20.15]", r.Desc2);
        }

        [Fact]
        public void A_permanent_part_with_no_variable_part_anywhere_is_not_flagged()
        {
            var r = SwpDescriptions.Compose(Entry("CIRCUIT BREAKER TRIPPED"), FrenchGlossary(true));
            Assert.Equal("DISJONCTEUR MAGNÉTOTHERMIQUE DÉCLENCHÉ ", r.Desc2);
            Assert.False(r.Desc2VariablePartFromLanguage1);
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
