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

        /// <summary>
        /// SAFETY is the sorter's cabinet family and nothing else: the safety tooling writes its
        /// mnemonic column in sorter syntax (the FVT list carries CABINETAL00000…00935), and a
        /// CLASSDIAG shape — a sorter object with no sheet on a cabinet-only SWP — is not a
        /// diagnostic there at all.
        /// </summary>
        [Fact]
        public void Safety_takes_the_cabinet_shapes_and_leaves_classdiag_unclaimed()
        {
            var log = new List<SwpLogEntry>();
            Assert.Equal("CABINETAL00800", Assert.Single(Run(Row("CABINETAL00800"), SwpPlantType.Safety, log)).Mnemonic);
            Assert.Equal("CABINETWR00705", Assert.Single(Run(Row("CABINETWR00705"), SwpPlantType.Safety, log)).Mnemonic);
            Assert.Equal("EMERGENCY00100", Assert.Single(Run(Row("EMERGENCY00100"), SwpPlantType.Safety, log)).Mnemonic);
            Assert.Empty(Run(Row("EMERGENCY_ON"), SwpPlantType.Safety, log));            // a command, recognized and dropped
            Assert.Empty(Run(Row("CLASSDIAGAL00620000705"), SwpPlantType.Safety, log));  // a chute: sorter only
            Assert.Single(Run(Row("CLASSDIAGAL00620000705"), SwpPlantType.Msc, log));
            Assert.Empty(log);
        }

        /// <summary>
        /// The field gap of 2026-09-22: 454 CABINETAL rows in the FVT safety list, 70 in the
        /// workspace — the 384 signal rows carry Siemens "I.byte.bit" addresses (no node) and
        /// were dropped as unreadable. SAFETY reads them and hands the SWP the TIA form; every
        /// other plant type keeps the error, because on a sorter list that shape is a missing bit.
        /// </summary>
        [Fact]
        public void Safety_reads_siemens_plc_addresses_and_writes_the_tia_form()
        {
            var log = new List<SwpLogEntry>();
            DiagnosticEntry fvt = Assert.Single(Run(Row("CABINETAL00104", address: "I.0.4"), SwpPlantType.Safety, log));
            Assert.Equal("I0.4", fvt.PlcAddress);
            Assert.Equal("I0.4", fvt.SwpAddress());
            Assert.Equal("", fvt.Node);
            Assert.False(fvt.IsBusModule);

            DiagnosticEntry fvz = Assert.Single(Run(Row("CABINETAL00206", address: "I12.6"), SwpPlantType.Safety, log));
            Assert.Equal("I12.6", fvz.SwpAddress());
            Assert.Empty(log);

            Assert.Empty(Run(Row("CABINETAL00300", address: "Q.0.1"), SwpPlantType.Safety, log));   // an output
            Assert.Empty(log);

            // A sorter list with the same shape has lost its bit — still an error, still dropped.
            Assert.Empty(Run(Row("CABINETAL00104", address: "I.0.4"), SwpPlantType.Msc, log));
            SwpLogEntry error = Assert.Single(log);
            Assert.Equal(SwpLogSeverity.Error, error.Severity);
            Assert.Contains("unreadable I/O address \"I.0.4\"", error.Message);

            // A module row on the safety list is unchanged: node from the ID, no bit address.
            DiagnosticEntry module = Assert.Single(Run(Row("CABINETAL00800", address: "", id: "23", type: "PA"), SwpPlantType.Safety));
            Assert.True(module.IsBusModule);
            Assert.Equal("23.0.0", module.SwpAddress());
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
        public void Emergency_keeps_its_own_name()
        {
            // It used to be RENAMED to CABINETAL so it would ride the cabinet branch — the two
            // prefixes are both nine characters, which is the only reason that worked. The rule
            // table gives it a destination instead, so the name Plant_Inputs knows it by survives.
            var entries = Run(Row("EMERGENCY00100"), SwpPlantType.Msc);
            DiagnosticEntry e = Assert.Single(entries);
            Assert.Equal("EMERGENCY00100", e.Mnemonic);
            Assert.Equal("EMERGENCY00100", e.RawMnemonic);
        }

        [Fact]
        public void Emeri_keeps_its_own_name()
        {
            var entries = Run(Row("EMERI_1_01_07"), SwpPlantType.Ups);
            Assert.Equal("EMERI_1_01_07", Assert.Single(entries).Mnemonic);
        }

        [Fact]
        public void Only_the_zero_member_of_an_and_group_is_kept()
        {
            // Many digital inputs drive ONE diagnostic bit; the string is built from the _0 member
            // and the rest are recognized and dropped. The legacy tool got here by rewriting the
            // name and chopping two characters off it.
            Assert.Equal("ALLQ_AND_1_01_00_0",
                Assert.Single(Run(Row("ALLQ_AND_1_01_00_0"), SwpPlantType.Ups)).Mnemonic);

            Assert.Empty(Run(Row("ALLQ_AND_1_01_00_1"), SwpPlantType.Ups));
            Assert.Empty(Run(Row("ALLQ_OR_1_01_00_7"), SwpPlantType.Ups));
        }

        [Fact]
        public void A_shape_no_rule_claims_is_dropped_quietly()
        {
            // The per-row length ERROR is gone with the length gate that produced it: an unclaimed
            // mnemonic is now just one of the many non-diagnostic rows, and those never logged.
            var log = new List<SwpLogEntry>();
            Assert.Empty(Run(Row("CABINETAL001000"), SwpPlantType.Msc, log));
            Assert.DoesNotContain(log, l => l.Severity == SwpLogSeverity.Error);
        }

        [Fact]
        public void Claiming_almost_nothing_warns_about_the_plant_type()
        {
            // What replaces the per-row errors: the RATIO. A sorter IO list read as an upstream
            // plant matches almost nothing, and that must not pass for a clean build.
            var log = new List<SwpLogEntry>();
            var rows = new List<IoRowData>();
            for (int i = 0; i < 40; i++) rows.Add(Row("CABINETAL" + i.ToString("000") + "00"));

            Assert.Empty(SwpMnemonicFilter.Run(rows,
                new SwpFilterOptions { PlantType = SwpPlantType.Ups }, log));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("check the plant type"));
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

            // The rules are ANCHORED, which is what keeps the field case fixed: EMERGENCY at the
            // start is a cabinet alarm, EMERGENCY in the middle of a command is nothing.
            Assert.Single(Run(Row("EMERGENCY00105"), SwpPlantType.Msc));
            Assert.Single(Run(Row("EMERI_1_01_05"), SwpPlantType.Ups));
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
