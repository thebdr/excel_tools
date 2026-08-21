using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpGlossaryTests
    {
        // ---- synthetic glossary workbooks (both real layouts) -------------------------------

        private static Cell Inline(string reference, string text) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new XlsxText(text)),
        };

        /// <summary>The Troubleshooting layout: key + languages from column I onwards.</summary>
        private static byte[] TroubleshootingLayout()
        {
            return Build(("Glossario", new[]
            {
                new Row(Inline("I1", "Numero scheda "), Inline("J1", "EN"), Inline("K1", "IT"),
                    Inline("L1", "FR"), Inline("M1", "PT")) { RowIndex = 1 },
                new Row(Inline("I2", "TS_001"), Inline("J2", "EMERGENCY ACTIVE"),
                    Inline("K2", "EMERGENZA ATTIVA"), Inline("L2", "URGENCE ACTIVE"),
                    Inline("M2", "EMERGÊNCIA ATIVA")) { RowIndex = 2 },
                new Row(Inline("I3", "TS_003"), Inline("J3", "RELAY"),
                    Inline("K3", "RELE'"), Inline("L3", "RELAIS"), Inline("M3", "RELÉ")) { RowIndex = 3 },
                new Row(Inline("I4", "TS_003"), Inline("J4", "FEEDBACK OF SAFETY RELAY FAULT"),
                    Inline("K4", "GUASTO DEL RELÈ DI SICUREZZA"), Inline("L4", "DÉFAUT DU RELAIS DE SÉCURITÉ"),
                    Inline("M4", "AVARIA DO RELÉ DE SEGURANÇA")) { RowIndex = 4 },
                new Row(Inline("I5", "TS_009"), Inline("J5", "NO PORTUGUESE HERE"),
                    Inline("K5", "SENZA PORTOGHESE")) { RowIndex = 5 },
            }));
        }

        /// <summary>The ombra layout: languages from column B, no key column values.</summary>
        private static byte[] OmbraLayout()
        {
            return Build(("GlossarioOmbra", new[]
            {
                new Row(Inline("A1", "scheda"), Inline("B1", "EN"), Inline("C1", "IT"),
                    Inline("D1", "FR"), Inline("E1", "PT")) { RowIndex = 1 },
                new Row(Inline("B2", "SPARE STATE.00"), Inline("C2", "STATO DI SCORTA.00"),
                    Inline("D2", "ÉTAT DE RECHANGE.00"), Inline("E2", "ESTADO SOBRESSALENTE.00")) { RowIndex = 2 },
                new Row(Inline("B3", "RELAY"), Inline("C3", "RELÈ OMBRA"),
                    Inline("D3", "RELAIS OMBRE"), Inline("E3", "RELÉ SOMBRA")) { RowIndex = 3 },
            }));
        }

        private static byte[] Build(params (string Name, Row[] Rows)[] sheets)
        {
            using var stream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbook = document.AddWorkbookPart();
                workbook.Workbook = new Workbook();
                var sheetList = new Sheets();
                uint id = 1;
                foreach ((string name, Row[] rows) in sheets)
                {
                    var part = workbook.AddNewPart<WorksheetPart>();
                    var data = new SheetData();
                    foreach (Row row in rows) data.AppendChild(row);
                    part.Worksheet = new Worksheet(data);
                    sheetList.AppendChild(new Sheet { Id = workbook.GetIdOfPart(part), SheetId = id++, Name = name });
                }
                workbook.Workbook.AppendChild(sheetList);
            }
            return stream.ToArray();
        }

        // ---- parsing ------------------------------------------------------------------------

        [Fact]
        public void Parses_the_troubleshooting_layout_for_a_language()
        {
            var log = new List<SwpLogEntry>();
            List<SwpGlossaryPair> pairs = SwpGlossary.Parse(TroubleshootingLayout(), "fr", "standard", log);

            Assert.Equal(3, pairs.Count);   // the fourth row has no FR? all have FR except TS_009
            Assert.Contains(pairs, p => p.En == "RELAY" && p.Translation == "RELAIS");
            Assert.Contains(pairs, p => p.En == "FEEDBACK OF SAFETY RELAY FAULT" && p.Translation == "DÉFAUT DU RELAIS DE SÉCURITÉ");
            Assert.All(pairs, p => Assert.Equal("standard", p.Source));
            Assert.Empty(log);
        }

        [Fact]
        public void Parses_the_ombra_layout()
        {
            List<SwpGlossaryPair> pairs = SwpGlossary.Parse(OmbraLayout(), "PT", "shadow", new List<SwpLogEntry>());
            Assert.Equal(2, pairs.Count);
            Assert.Contains(pairs, p => p.En == "SPARE STATE.00" && p.Translation == "ESTADO SOBRESSALENTE.00");
        }

        [Fact]
        public void Rows_without_the_language_are_skipped()
        {
            List<SwpGlossaryPair> pairs = SwpGlossary.Parse(TroubleshootingLayout(), "PT", "standard", new List<SwpLogEntry>());
            Assert.DoesNotContain(pairs, p => p.En == "NO PORTUGUESE HERE");
        }

        [Fact]
        public void A_missing_language_column_warns()
        {
            var log = new List<SwpLogEntry>();
            Assert.Empty(SwpGlossary.Parse(TroubleshootingLayout(), "DE", "standard", log));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("DE"));
        }

        [Fact]
        public void Available_languages_come_from_the_header()
        {
            List<string> codes = SwpGlossary.AvailableLanguages(TroubleshootingLayout());
            Assert.Equal(new[] { "IT", "FR", "PT" }, codes);
        }

        [Fact]
        public void Ts_refs_come_from_the_numero_scheda_column()
        {
            List<SwpGlossaryPair> pairs = SwpGlossary.Parse(TroubleshootingLayout(), "FR", "standard", new List<SwpLogEntry>());
            Assert.Equal("TS_001", pairs.First(p => p.En == "EMERGENCY ACTIVE").TsRef);
            Assert.Equal("TS_003", pairs.First(p => p.En == "RELAY").TsRef);
        }

        [Fact]
        public void The_ombra_scheda_column_never_becomes_a_ts_ref()
        {
            // The ombra's bare "scheda" column holds template TEXTS ("SPARE STATE.00"), not
            // troubleshooting cards — only "Numero scheda"-style headers qualify.
            List<SwpGlossaryPair> pairs = SwpGlossary.Parse(OmbraLayout(), "PT", "shadow", new List<SwpLogEntry>());
            Assert.All(pairs, p => Assert.Equal("", p.TsRef));
        }

        [Fact]
        public void Merging_keeps_the_standard_card_when_the_shadow_rewords()
        {
            var standard = SwpGlossary.Parse(TroubleshootingLayout(), "PT", "standard", new List<SwpLogEntry>());
            var shadow = SwpGlossary.Parse(OmbraLayout(), "PT", "shadow", new List<SwpLogEntry>());

            List<SwpGlossaryPair> merged = SwpGlossary.Merge(standard, shadow, null);
            SwpGlossaryPair relay = merged.First(p => p.En == "RELAY");
            Assert.Equal("RELÉ SOMBRA", relay.Translation);   // the shadow's wording…
            Assert.Equal("TS_003", relay.TsRef);              // …with the standard's card intact
        }

        // ---- merge + tweaks -----------------------------------------------------------------

        [Fact]
        public void The_shadow_file_wins_on_a_shared_phrase()
        {
            var standard = SwpGlossary.Parse(TroubleshootingLayout(), "PT", "standard", new List<SwpLogEntry>());
            var shadow = SwpGlossary.Parse(OmbraLayout(), "PT", "shadow", new List<SwpLogEntry>());

            List<SwpGlossaryPair> merged = SwpGlossary.Merge(standard, shadow, null);
            SwpGlossaryPair relay = merged.First(p => p.En == "RELAY");
            Assert.Equal("RELÉ SOMBRA", relay.Translation);
            Assert.Equal("shadow", relay.Source);
        }

        [Fact]
        public void Tweaks_disable_correct_and_add_pairs()
        {
            var standard = new List<SwpGlossaryPair>
            {
                new SwpGlossaryPair { En = "RELAY", Translation = "RELÉ" },
                new SwpGlossaryPair { En = "MOTOR", Translation = "MOTOR-PT" },
            };
            var tweaks = new List<SwpGlossaryTweak>
            {
                new SwpGlossaryTweak { En = "RELAY", Enabled = false },
                new SwpGlossaryTweak { En = "MOTOR", Enabled = true, Translation = "MOTORE CORRETTO" },
                new SwpGlossaryTweak { En = "BELT", Enabled = true, Translation = "TAPETE" },
            };

            List<SwpGlossaryPair> merged = SwpGlossary.Merge(standard, null, tweaks);
            Assert.False(merged.First(p => p.En == "RELAY").Enabled);
            Assert.Equal("MOTORE CORRETTO", merged.First(p => p.En == "MOTOR").Translation);
            SwpGlossaryPair custom = merged.First(p => p.En == "BELT");
            Assert.Equal("custom", custom.Source);
            Assert.Equal("TAPETE", custom.Translation);
        }

        // ---- translate + relevance ----------------------------------------------------------

        private static List<SwpGlossaryPair> Pairs(params (string En, string T)[] pairs) =>
            pairs.Select(p => new SwpGlossaryPair { En = p.En, Translation = p.T }).ToList();

        [Fact]
        public void Only_a_full_permanent_part_match_translates()
        {
            // Substring splicing was retired: "RELAY FAULT POWER SUPPLY" contains the phrase
            // but IS not the phrase — no translation, language-1 fallback instead.
            var pairs = Pairs(("RELAY FAULT", "AVARIA DO RELÉ"));
            var partial = SwpDescriptions.Compose(Entry("RELAY FAULT POWER SUPPLY"), Ctx(pairs, always: true));
            Assert.Equal("RELAY FAULT POWER SUPPLY ", partial.Desc2);
            Assert.False(partial.GlossaryDiffers);

            // Case-insensitive on the whole phrase, and the VARIABLE part rides along.
            var exact = SwpDescriptions.Compose(
                new DiagnosticEntry { Mnemonic = "CABINETAL00100", Desc1A = "Relay Fault", Desc1B = "+MS1-K1" },
                Ctx(pairs, always: true));
            Assert.Equal("AVARIA DO RELÉ +MS1-K1", exact.Desc2);
        }

        [Fact]
        public void A_disabled_pair_does_not_translate()
        {
            var pairs = Pairs(("RELAY FAULT", "AVARIA DO RELÉ"));
            pairs[0].Enabled = false;
            var r = SwpDescriptions.Compose(Entry("RELAY FAULT"), Ctx(pairs, always: true));
            Assert.Equal("RELAY FAULT ", r.Desc2);   // untranslated → language-1 fallback
        }

        [Fact]
        public void Relevance_keeps_only_exactly_used_pairs_plus_customs()
        {
            var pairs = Pairs(("RELAY FAULT", "AVARIA"), ("NEVER SEEN", "NUNCA"));
            pairs.Add(new SwpGlossaryPair { En = "CUSTOM", Translation = "PERSONALIZADO", Source = "custom" });

            // Keys are rows' permanent parts: two rows carry the phrase, one merely contains it.
            List<SwpGlossaryPair> relevant = SwpGlossary.RelevantPairs(pairs,
                new[] { "RELAY FAULT", "relay fault ", "RELAY FAULT POWER SUPPLY" },
                out Dictionary<string, int> occurrences);

            Assert.Equal(2, relevant.Count);
            Assert.Equal(2, occurrences["RELAY FAULT"]);
            Assert.Contains(relevant, p => p.En == "CUSTOM");
            Assert.DoesNotContain(relevant, p => p.En == "NEVER SEEN");
        }

        // ---- compose modes ------------------------------------------------------------------

        private static DiagnosticEntry Entry(string en, string lang2 = "") => new DiagnosticEntry
        {
            Mnemonic = "CABINETAL00100",
            Desc1A = en,
            Desc2A = lang2,
        };

        private static SwpDescriptions.Context Ctx(List<SwpGlossaryPair> pairs, bool always) =>
            new SwpDescriptions.Context { GlossaryPairs = pairs, GlossaryTranslateAlways = always };

        [Fact]
        public void Fill_gaps_mode_translates_only_empty_language_two()
        {
            var pairs = Pairs(("RELAY FAULT", "AVARIA DO RELÉ"));

            var gap = SwpDescriptions.Compose(Entry("RELAY FAULT"), Ctx(pairs, always: false));
            Assert.Equal("AVARIA DO RELÉ ", gap.Desc2);   // legacy trailing space carried through

            var kept = SwpDescriptions.Compose(Entry("RELAY FAULT", "TEXTO DO IO"), Ctx(pairs, always: false));
            Assert.Equal("TEXTO DO IO ", kept.Desc2);
            Assert.True(kept.GlossaryDiffers);            // but the disagreement is flagged
            Assert.Equal("AVARIA DO RELÉ ", kept.GlossaryText);
        }

        [Fact]
        public void Always_mode_overrides_the_io_list()
        {
            var pairs = Pairs(("RELAY FAULT", "AVARIA DO RELÉ"));
            var r = SwpDescriptions.Compose(Entry("RELAY FAULT", "TEXTO DO IO"), Ctx(pairs, always: true));
            Assert.Equal("AVARIA DO RELÉ ", r.Desc2);
            Assert.True(r.GlossaryDiffers);
        }

        [Fact]
        public void No_hit_keeps_the_fallback_chain()
        {
            var pairs = Pairs(("SOMETHING ELSE", "OUTRA COISA"));
            var r = SwpDescriptions.Compose(Entry("RELAY FAULT"), Ctx(pairs, always: true));
            Assert.Equal("RELAY FAULT ", r.Desc2);   // untranslated → language-1 fallback
            Assert.False(r.GlossaryDiffers);
        }

        [Fact]
        public void Agreement_is_not_flagged()
        {
            var pairs = Pairs(("RELAY FAULT", "AVARIA DO RELÉ"));
            var r = SwpDescriptions.Compose(Entry("RELAY FAULT", "AVARIA DO RELÉ"), Ctx(pairs, always: false));
            Assert.False(r.GlossaryDiffers);
        }

        // ---- the real glossario files, when this machine has them ---------------------------

        private const string RealStandard =
            @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\3 Programmi\0_Fives" +
            @"\0vba_  ConfigPlantTranslator\Glossario Troubleshooting_rev-009.06.xlsx";
        private const string RealOmbra =
            @"C:\Users\bogdan.dragoi\OneDrive - FIVES\Work\3 Programmi\0_Fives" +
            @"\0vba_  ConfigPlantTranslator\Glossario_ombra_v42.xlsx";

        [Fact]
        public void Real_glossario_files_parse_for_french()
        {
            if (!File.Exists(RealStandard) || !File.Exists(RealOmbra)) return;

            var log = new List<SwpLogEntry>();
            var standard = SwpGlossary.Parse(File.ReadAllBytes(RealStandard), "FR", "standard", log);
            var shadow = SwpGlossary.Parse(File.ReadAllBytes(RealOmbra), "FR", "shadow", log);

            Assert.True(standard.Count > 200, "standard pairs: " + standard.Count);
            Assert.True(shadow.Count > 100, "shadow pairs: " + shadow.Count);
            Assert.Contains(standard, p => p.En == "EMERGENCY ACTIVE" && p.Translation == "URGENCE ACTIVE");
            Assert.Contains(shadow, p => p.En == "SPARE STATE.00" && p.Translation == "ÉTAT DE RECHANGE.00");

            // Merged (the shadow may re-word shared phrases): a full-permanent-part match
            // translates and the variable part rides along.
            List<SwpGlossaryPair> merged = SwpGlossary.Merge(standard, shadow, null);
            var entry = new DiagnosticEntry
            {
                Mnemonic = "CABINETAL00100",
                Desc1A = "EMERGENCY ACTIVE",
                Desc1B = "+MS1.CC1",
            };
            var translated = SwpDescriptions.Compose(entry, Ctx(merged, always: true));
            Assert.NotEqual("EMERGENCY ACTIVE +MS1.CC1", translated.Desc2);
            Assert.EndsWith("+MS1.CC1", translated.Desc2);

            var standardOnly = SwpDescriptions.Compose(entry, Ctx(standard, always: true));
            Assert.StartsWith("URGENCE ACTIVE", standardOnly.Desc2);
        }

        [Fact]
        public void Real_glossario_carries_the_bus_failure_family_and_its_cards()
        {
            if (!File.Exists(RealStandard)) return;

            var standard = SwpGlossary.Parse(File.ReadAllBytes(RealStandard), "FR", "standard", new List<SwpLogEntry>());

            // The dedicated row every unofficial bus-module description falls back to.
            SwpGlossaryPair dedicated = standard.First(p => p.En == "FIELD BUS MODULE FAILURE");
            Assert.True(dedicated.Translation.Trim().Length > 0);

            // The official variant family exists and the cards flow from "Numero scheda".
            Assert.Contains(standard, p => p.En == "FIELD BUS DEVICE FAILURE");
            Assert.Contains(standard, p => p.En.StartsWith("FIELD BUS FAILURE -"));
            Assert.Contains(standard, p => p.En == "STATE OF EMERGENCY" && p.TsRef == "TS_001");
        }
    }
}
