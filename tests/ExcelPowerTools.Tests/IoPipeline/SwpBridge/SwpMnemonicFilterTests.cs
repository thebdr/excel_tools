using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpMnemonicFilterTests
    {
        private static IoRowData Row(string mnemonic, string address = "I.80.3.1", string id = "",
            string type = "1")
        {
            return new IoRowData
            {
                CardSheet = "R1.S1.C1",
                SourceRow = 10,
                MnemonicCell = mnemonic,
                AddressCell = address,
                IdCell = id,
                TypeCell = type,
            };
        }

        private static List<DiagnosticEntry> Run(IoRowData row, SwpPlantType plant,
            List<SwpLogEntry> log = null, Regex include = null, Regex exclude = null)
        {
            return SwpMnemonicFilter.Run(new[] { row },
                new SwpFilterOptions { PlantType = plant, IncludeMnemonics = include, ExcludeMnemonics = exclude },
                log ?? new List<SwpLogEntry>());
        }

        [Fact]
        public void Accepts_a_sorter_cabinet_alarm()
        {
            var entries = Run(Row("CABINETAL00100"), SwpPlantType.Msc);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.Equal("CABINETAL00100", e.Mnemonic);
            Assert.Equal("80", e.Node);
            Assert.Equal("3", e.Offset);
            Assert.Equal("1", e.Bit);
            Assert.False(e.IsBusModule);
            Assert.Equal("80.3.1", e.SwpAddress());
        }

        [Fact]
        public void Splits_multi_mnemonic_cells_into_separate_entries()
        {
            var entries = Run(Row("CABINETAL00100/CABINETWR00101"), SwpPlantType.Msc);
            Assert.Equal(2, entries.Count);
            Assert.Equal("CABINETAL00100", entries[0].Mnemonic);
            Assert.Equal("CABINETWR00101", entries[1].Mnemonic);
        }

        [Fact]
        public void Strips_the_bus_prefix_and_keeps_the_raw_text()
        {
            var entries = Run(Row("BUS_CABINETAL00161", address: "", id: "236"), SwpPlantType.Msc);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.Equal("CABINETAL00161", e.Mnemonic);
            Assert.Equal("BUS_CABINETAL00161", e.RawMnemonic);
            Assert.True(e.IsBusModule);
            Assert.Equal("236", e.Node);
            Assert.Equal("236.0.0", e.SwpAddress());
        }

        [Fact]
        public void Skips_non_diagnostic_mnemonics_silently()
        {
            var log = new List<SwpLogEntry>();
            Assert.Empty(Run(Row("MTRSTART_00001"), SwpPlantType.Msc, log));
            Assert.Empty(log);
        }

        [Fact]
        public void Skips_short_names_even_when_a_pattern_substring_matches()
        {
            Assert.Empty(Run(Row("CABINET1"), SwpPlantType.Msc));
        }

        [Theory]
        [InlineData("XEMERGENCY_ON001")]
        [InlineData("EMERGENCY_RST001")]
        public void Emergency_commands_are_never_diagnostics(string mnemonic)
        {
            Assert.Empty(Run(Row(mnemonic), SwpPlantType.Msc));
        }

        [Fact]
        public void Emergency_rewrites_to_cabinet_alarm()
        {
            var entries = Run(Row("EMERGENCY00100"), SwpPlantType.Msc);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.Equal("CABINETAL00100", e.Mnemonic);
            Assert.Equal("EMERGENCY00100", e.RawMnemonic);
        }

        [Fact]
        public void Emeri_rewrites_to_allq()
        {
            var entries = Run(Row("EMERI_1_01_07"), SwpPlantType.Ups);
            Assert.Equal("ALLQ_1_01_07", Assert.Single(entries).Mnemonic);
        }

        [Fact]
        public void Allq_and_collapses_and_chops_two_characters()
        {
            // Legacy verbatim: replace ALLQ_AND_ with ALLQ_, then drop the last two characters.
            var log = new List<SwpLogEntry>();
            var entries = Run(Row("ALLQ_AND_1_01_00"), SwpPlantType.Ups, log);
            Assert.Equal("ALLQ_1_01_", Assert.Single(entries).Mnemonic);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning); // unusual length, kept
        }

        [Fact]
        public void Msc_rejects_wrong_lengths_with_an_error_log()
        {
            var log = new List<SwpLogEntry>();
            Assert.Empty(Run(Row("CABINETAL001000"), SwpPlantType.Msc, log));
            SwpLogEntry entry = Assert.Single(log);
            Assert.Equal(SwpLogSeverity.Error, entry.Severity);
            Assert.Equal("CABINETAL001000", entry.Mnemonic);
        }

        [Theory]
        [InlineData("CABINETAL000162", true)]
        [InlineData("BELTAL10162", true)]
        [InlineData("ZONEAL1162", true)]
        [InlineData("BELTAL101620", false)]
        public void Ispc_length_rules(string mnemonic, bool accepted)
        {
            Assert.Equal(accepted, Run(Row(mnemonic), SwpPlantType.Ispc).Count == 1);
        }

        [Fact]
        public void Output_addresses_never_reach_the_swp()
        {
            Assert.Empty(Run(Row("CABINETAL00100", address: "Q.80.3.1"), SwpPlantType.Msc));
        }

        [Fact]
        public void Address_components_normalize_numerically()
        {
            var entries = Run(Row("CABINETAL00100", address: "I.080.03.01"), SwpPlantType.Msc);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.Equal("80.3.1", e.SwpAddress());
        }

        [Fact]
        public void Unreadable_dotted_address_rejects_with_error()
        {
            var log = new List<SwpLogEntry>();
            Assert.Empty(Run(Row("CABINETAL00100", address: "I.80.3"), SwpPlantType.Msc, log));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("address"));
        }

        [Theory]
        [InlineData("1A", "1A")]
        [InlineData("01a", "1A")]
        [InlineData("10B", "10B")]
        public void Asi_node_ids_are_real_bus_identities(string id, string expectedNode)
        {
            var entries = Run(Row("CLASSDIAGAL00620030130", address: "", id: id), SwpPlantType.Msc);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.True(e.IsBusModule);
            Assert.Equal(expectedNode, e.Node);
            Assert.Equal("1M." + expectedNode, e.AsiNode);        // no Slot column → line 1
            Assert.Equal("1M." + expectedNode, e.SwpAddress());   // G gets line.slave, no ".0.0"
        }

        [Fact]
        public void Module_row_with_a_non_node_id_stays_but_carries_no_address()
        {
            var log = new List<SwpLogEntry>();
            var entries = Run(Row("CABINETAL00161", address: "", id: "n/a"), SwpPlantType.Msc, log);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.False(e.IsBusModule);
            Assert.Equal("", e.SwpAddress());
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info && l.Message.Contains("not a bus node"));
        }

        [Fact]
        public void Include_regex_rescues_a_foreign_shape()
        {
            var entries = Run(Row("MYDIAG_1_01_07"), SwpPlantType.Ups,
                include: new Regex("^MYDIAG"));
            Assert.Equal("MYDIAG_1_01_07", Assert.Single(entries).Mnemonic);
        }

        [Fact]
        public void Include_regex_bypasses_every_length_rule()
        {
            // Too long for the base gate AND not a sorter length — the rescue must win anyway.
            var entries = Run(Row("CLASSDIAGAL006200301305X"), SwpPlantType.Msc,
                include: new Regex("^CLASSDIAG"));
            Assert.Equal("CLASSDIAGAL006200301305X", Assert.Single(entries).Mnemonic);
        }

        [Fact]
        public void Exclude_regex_drops_an_accepted_mnemonic_with_a_note()
        {
            var log = new List<SwpLogEntry>();
            Assert.Empty(Run(Row("CABINETAL00100"), SwpPlantType.Msc, log,
                exclude: new Regex("^CABINETAL001")));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info && l.Message.Contains("regex"));
        }

        [Fact]
        public void Row_data_carries_through_to_the_entry()
        {
            var row = Row("CABINETAL00100");
            row.Desc1A = "CIRCUIT BREAKER TRIPPED";
            row.Desc1B = "POWER SUPPLY";
            row.Desc2A = "DISJONCTEUR";
            row.Desc2B = "ALIMENTATION";
            row.FUnit = "=S1";
            row.Location = "+MC1.CC1";
            row.Device = "-F22001";
            row.TsRef = "TS_001";
            row.ProfinetName = "n001-dev";
            row.ProfinetIp = "192.168.0.10";
            DiagnosticEntry e = Assert.Single(Run(row, SwpPlantType.Msc));
            Assert.Equal("CIRCUIT BREAKER TRIPPED", e.Desc1A);
            Assert.Equal("POWER SUPPLY", e.Desc1B);
            Assert.Equal("DISJONCTEUR", e.Desc2A);
            Assert.Equal("ALIMENTATION", e.Desc2B);
            Assert.Equal("=S1", e.FUnit);
            Assert.Equal("+MC1.CC1", e.Location);
            Assert.Equal("-F22001", e.Device);
            Assert.Equal("TS_001", e.TsRef);
            Assert.Equal("n001-dev", e.ProfinetName);
            Assert.Equal("192.168.0.10", e.ProfinetIp);
            Assert.Equal("R1.S1.C1", e.CardSheet);
            Assert.Equal(10, e.SourceRow);
        }

        [Fact]
        public void Empty_and_whitespace_parts_are_skipped()
        {
            Assert.Empty(SwpMnemonicFilter.Run(
                new[] { Row("  / "), Row(""), null },
                new SwpFilterOptions { PlantType = SwpPlantType.Msc },
                new List<SwpLogEntry>()));
        }

        [Fact]
        public void Modifier_rewrites_only_fire_at_the_start_of_the_mnemonic()
        {
            // Field case: A_RESET_EMERGENCY is a COMMAND — the old anywhere-Replace turned it
            // into A_RESET_CABINETAL and spammed a length error. Mid-string modifiers neither
            // rewrite nor qualify as diagnostics.
            var log = new List<SwpLogEntry>();
            Assert.Empty(SwpMnemonicFilter.Run(
                new[] { Row("A_RESET_EMERGENCY"), Row("X_EMERI_1_01_05") },
                new SwpFilterOptions { PlantType = SwpPlantType.Msc }, log));
            Assert.Empty(log);

            // Anchored occurrences keep the legacy behavior.
            Assert.Equal("CABINETAL00105", SwpMnemonicFilter.ApplyRewrites("EMERGENCY00105"));
            Assert.Equal("ALLQ_1_01_05", SwpMnemonicFilter.ApplyRewrites("EMERI_1_01_05"));
            Assert.Equal("A_RESET_EMERGENCY", SwpMnemonicFilter.ApplyRewrites("A_RESET_EMERGENCY"));
        }

        // ---- stray-character hygiene --------------------------------------------------------

        [Fact]
        public void The_denylist_cleans_known_debris_and_keeps_everything_else()
        {
            Assert.Equal("-A1001", SwpMnemonicFilter.CleanStray("'-A1001"));
            Assert.Equal("n001", SwpMnemonicFilter.CleanStray("“n001”"));
            Assert.Equal("+MC1 .CC1", SwpMnemonicFilter.CleanStray("’+MC1 .CC1"));
            Assert.Equal("x y", SwpMnemonicFilter.CleanStray("x\ty"));
            Assert.Equal("ab", SwpMnemonicFilter.CleanStray("a​﻿b`"));
            // Unknown characters are NOT touched — that is the point of the denylist.
            Assert.Equal("=S1+CH§é#", SwpMnemonicFilter.CleanStray("=S1+CH§é#"));
        }

        [Fact]
        public void Identity_fields_are_cleaned_with_one_log_per_distinct_value()
        {
            IoRowData Dirty() => new IoRowData
            {
                CardSheet = "R1.S1.C1",
                SourceRow = 10,
                MnemonicCell = "CABINETAL00100",
                AddressCell = "'I.80.3.0",       // the stray apostrophe must not break parsing
                FUnit = "'=S1",
                Location = "'+MC1.CC1",
                Device = "`-F1",
                ProfinetName = "‘n001’",
                ProfinetIp = "'192.168.1.1",
            };

            var log = new List<SwpLogEntry>();
            List<DiagnosticEntry> entries = SwpMnemonicFilter.Run(new[] { Dirty(), Dirty() },
                new SwpFilterOptions { PlantType = SwpPlantType.Msc }, log);

            Assert.Equal(2, entries.Count);
            DiagnosticEntry e = entries[0];
            Assert.Equal("=S1", e.FUnit);
            Assert.Equal("+MC1.CC1", e.Location);
            Assert.Equal("-F1", e.Device);
            Assert.Equal("n001", e.ProfinetName);
            Assert.Equal("192.168.1.1", e.ProfinetIp);
            Assert.Equal("80", e.Node);   // the address parsed after cleaning

            // One Info line per DISTINCT field+value, not per row.
            Assert.Equal(6, log.Count(l => l.Message.Contains("stray characters cleaned")));
        }
    }
}
