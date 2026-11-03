using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The index-rules table against the official "SWP and diagnostic configuration rules" v93
    /// (§4.9 / §5.2 / §7.10), the SAFETY plant type it enables, and the content checks the same
    /// document defines.
    /// </summary>
    public class SwpIndexRulesTests
    {
        // ---- the official table -------------------------------------------------------------

        [Theory]
        [InlineData(SwpPlantType.Msc, "CABINET", 0, 0, 1, true)]
        [InlineData(SwpPlantType.Msc, "CHUTE", 0, 0, 0, false)]
        [InlineData(SwpPlantType.Msc, "SCANNER", 0, 0, 1, false)]
        [InlineData(SwpPlantType.Msc, "AIR_BLOWER", 0, 0, 0, false)]
        [InlineData(SwpPlantType.Msc, "ITEM_LOCATOR", 0, 0, 1, false)]
        [InlineData(SwpPlantType.Ispc, "CABINET", 1, 0, 1, true)]
        [InlineData(SwpPlantType.Ispc, "GROUP_CABINET", 1, 0, 0, true)]
        [InlineData(SwpPlantType.Ispc, "BELT", 1, 0, 1, false)]
        [InlineData(SwpPlantType.Ups, "CABINET", 1, 0, 0, true)]
        [InlineData(SwpPlantType.Ups, "BELT", 1, 0, 1, true)]
        [InlineData(SwpPlantType.Ups, "ZONE", 1, 0, 1, true)]
        [InlineData(SwpPlantType.Safety, "CABINET", 0, 0, 0, true)]
        public void The_official_values_match_the_document(SwpPlantType plant, string objectKey,
            int mnemonicStart, int entryStart, int opcStart, bool check)
        {
            SwpIndexRule rule = SwpIndexRules.Official().RuleFor(plant, objectKey);
            Assert.Equal(mnemonicStart, rule.MnemonicStart);
            Assert.Equal(entryStart, rule.SwpEntryStart);
            Assert.Equal(opcStart, rule.OpcTagStart);
            Assert.Equal(check, rule.AlarmCheckEnabled);
        }

        [Fact]
        public void Derived_offsets_reproduce_the_v49_matching_semantics()
        {
            SwpIndexRules rules = SwpIndexRules.Official();
            Assert.Equal(1, rules.AlarmOffset(SwpPlantType.Msc, "CABINET"));            // AS = IL + 1
            Assert.Equal(0, rules.AlarmOffset(SwpPlantType.Ispc, "CABINET"));           // 1:1
            Assert.Equal(-1, rules.AlarmOffset(SwpPlantType.Ispc, "GROUP_CABINET"));    // AS = IL − 1
            Assert.Equal(-1, rules.AlarmOffset(SwpPlantType.Ups, "CABINET"));
            Assert.Equal(0, rules.AlarmOffset(SwpPlantType.Ups, "BELT"));
            Assert.Equal(0, rules.AlarmOffset(SwpPlantType.Safety, "CABINET"));         // 1:1, 0-based
            Assert.Equal(0, rules.EntryOffset(SwpPlantType.Msc, "CABINET"));
            Assert.Equal(-1, rules.EntryOffset(SwpPlantType.Ispc, "CABINET"));
            Assert.Equal(-1, rules.EntryOffset(SwpPlantType.Ups, "ZONE"));
        }

        [Fact]
        public void Unknown_objects_fall_back_to_the_plant_family_defaults()
        {
            SwpIndexRules rules = SwpIndexRules.Official();
            SwpIndexRule mscish = rules.RuleFor(SwpPlantType.Msc, "SOMETHING_NEW");
            SwpIndexRule upsish = rules.RuleFor(SwpPlantType.Ups, "SOMETHING_NEW");
            Assert.Equal(0, mscish.EntryOffset);
            Assert.Equal(-1, upsish.EntryOffset);
            Assert.False(mscish.AlarmCheckEnabled);
        }

        [Fact]
        public void The_table_round_trips_through_lines_and_tolerates_junk()
        {
            SwpIndexRules original = SwpIndexRules.Official();
            List<string> lines = original.ToLines();
            lines.Add("garbage");
            lines.Add("NotAPlant|X|0|0|0|x");
            lines.Add("Msc|CABINET|9|9|9|x");   // duplicate — the first row must win

            SwpIndexRules reread = SwpIndexRules.FromLines(lines);
            Assert.NotNull(reread);
            Assert.Equal(original.Rules.Count, reread!.Rules.Count);
            Assert.Equal(1, reread.AlarmOffset(SwpPlantType.Msc, "CABINET"));
            Assert.True(reread.AlarmCheckEnabled(SwpPlantType.Safety, "CABINET"));

            Assert.Null(SwpIndexRules.FromLines(new[] { "junk", "" }));
        }

        [Fact]
        public void An_edited_rule_redirects_the_join_and_the_alarm_match()
        {
            SwpIndexRules edited = SwpIndexRules.FromLines(new[] { "Msc|CABINET|0|1|0|x" })!;

            // Entry shift +1: CABINETAL00215 now lands on table entry 3.
            SwpMappedMnemonic mapped = SwpAddressMap.Map("CABINETAL00215", SwpPlantType.Msc,
                rules: edited);
            Assert.Equal(3, mapped.Coordinate.TableEntry);

            // Alarm shift 0: the export matches 1:1 instead of the official +1.
            var alarms = SwpAlarmStrings.Parse(new[] { "x,S1.CABINETAL00215,1,a,b,c,d,e,f,TEXT," },
                SwpPlantType.Msc, "t", new List<SwpLogEntry>());
            Assert.NotNull(SwpAlarmStrings.Match(mapped, SwpPlantType.Msc, 1, 1, alarms, edited));
            Assert.Null(SwpAlarmStrings.Match(mapped, SwpPlantType.Msc, 1, 1, alarms));   // official +1
        }

        // ---- SAFETY -------------------------------------------------------------------------

        [Fact]
        public void Safety_mnemonics_filter_map_and_match_one_to_one_zero_based()
        {
            // Filter: MSC-shaped acceptance under the SAFETY plant type.
            var rows = new List<IoRowData>
            {
                new IoRowData { MnemonicCell = "CABINETAL00215", AddressCell = "I.80.3.0", CardSheet = "R1.S1.C1" },
            };
            List<DiagnosticEntry> entries = SwpMnemonicFilter.Run(rows,
                new SwpFilterOptions { PlantType = SwpPlantType.Safety }, new List<SwpLogEntry>());
            DiagnosticEntry entry = Assert.Single(entries);

            // Map: the CABINET sheet, table entry = raw digits (no shift).
            SwpMappedMnemonic mapped = SwpAddressMap.Map(entry.Mnemonic, SwpPlantType.Safety);
            Assert.Equal("CABINET", mapped.Coordinate.SheetName);
            Assert.Equal(2, mapped.Coordinate.TableEntry);
            Assert.Equal("DB_CABINETS_DIAG.Alarm1", mapped.Coordinate.TagType);

            // Alarm strings: cabinet 2 exports as cabinet 2 (1:1, zero-based) — not the MSC +1.
            var alarms = SwpAlarmStrings.Parse(new[] { "x,S1.CABINETAL00215,1,a,b,c,d,e,f,TEXT," },
                SwpPlantType.Safety, "t", new List<SwpLogEntry>());
            Assert.Single(alarms);
            Assert.NotNull(SwpAlarmStrings.Match(mapped, SwpPlantType.Safety, 1, 1, alarms));
            SwpMappedMnemonic neighbour = SwpAddressMap.Map("CABINETAL00115", SwpPlantType.Safety);
            Assert.Null(SwpAlarmStrings.Match(neighbour, SwpPlantType.Safety, 1, 1, alarms));
        }

        [Fact]
        public void The_new_classdiag_types_map_to_their_verified_sheets()
        {
            SwpMappedMnemonic airBlower = SwpAddressMap.Map("CLASSDIAGAL01120000005", SwpPlantType.Msc);
            Assert.Equal("AIR_BLOWER", airBlower.Coordinate.SheetName);
            Assert.Equal("DB_AIR_BLOWER.Alarm1", airBlower.Coordinate.TagType);
            Assert.Equal(0, airBlower.Coordinate.TableEntry);

            SwpMappedMnemonic itemLocator = SwpAddressMap.Map("CLASSDIAGWR01130000133", SwpPlantType.Msc);
            Assert.Equal("ITEM_LOCATOR", itemLocator.Coordinate.SheetName);
            Assert.Equal("DB_ITEM_LOCATOR.Warning2", itemLocator.Coordinate.TagType);
            Assert.Equal(1, itemLocator.Coordinate.TableEntry);
        }

        // ---- the per-object alarm-check gate ------------------------------------------------

        [Fact]
        public void Classdiag_objects_stay_out_of_the_alarm_check()
        {
            // A chute whose numbers collide with an exported cabinet alarm: without the gate it
            // would false-match (or read NotFound); the rules table keeps it NotChecked.
            var alarms = SwpAlarmStrings.Parse(new[] { "x,S1.CABINETAL00215,1,a,b,c,d,e,f,TEXT," },
                SwpPlantType.Msc, "t", new List<SwpLogEntry>());
            var row = new SwpSpineRow
            {
                Entry = new DiagnosticEntry { Mnemonic = "CLASSDIAGAL00620000115", CardSheet = "R1.S1.C1" },
                Mapped = SwpAddressMap.Map("CLASSDIAGAL00620000115", SwpPlantType.Msc),
            };
            var model = new SwpWorkspaceModel();
            model.Rows.Add(row);

            SwpCoherence.Annotate(model, null, null, SwpPlantType.Msc, new List<SwpLogEntry>(), alarms);
            Assert.Equal(SwpAlarmVerdict.NotChecked, row.AlarmVerdict);
        }

        // ---- alarm-strings file hygiene -----------------------------------------------------

        [Fact]
        public void Duplicated_export_tags_warn_about_an_appended_file()
        {
            var log = new List<SwpLogEntry>();
            SwpAlarmStrings.Parse(new[]
            {
                "x,S1.CABINETAL00215,1,a,b,c,d,e,f,TEXT,",
                "x,S1.CABINETAL00315,1,a,b,c,d,e,f,OTHER,",
                "x,S1.CABINETAL00215,1,a,b,c,d,e,f,TEXT,",
            }, SwpPlantType.Msc, "Alarms_MSC.txt", log);

            SwpLogEntry warning = Assert.Single(log,
                l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("appended"));
            Assert.Contains("1 duplicated tag", warning.Message);
        }

        // ---- busfield rules (§3.2 / §3.7 / §4.8) --------------------------------------------

        private static CsvDocument Csv(params string[] lines) =>
            CsvFile.Load(Encoding.UTF8.GetBytes(string.Join("\r\n", lines)));

        private static SwpSpineRow BusModule(string mnemonic, string node)
        {
            return new SwpSpineRow
            {
                Entry = new DiagnosticEntry
                {
                    Mnemonic = mnemonic,
                    IsBusModule = true,
                    Node = node,
                    CardSheet = "R1.S1.C1",
                },
                Mapped = SwpAddressMap.Map(mnemonic, SwpPlantType.Msc),
            };
        }

        [Fact]
        public void A_modern_zeroed_busfield_row_is_accepted_for_type_11_inputs()
        {
            // Element 0 / bit 0 is the OFFICIAL configuration for inputs of type BUS (§3.7) —
            // the linkage lives in Plant_Inputs, so nothing here can disagree.
            CsvDocument busfield = Csv(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;236;32880;0;0;0;0;0;0;0;0;#;x");
            SwpSpineRow row = BusModule("CABINETAL00161", "236");
            var model = new SwpWorkspaceModel();
            model.Rows.Add(row);

            SwpCoherence.Annotate(model, null, busfield, SwpPlantType.Msc, new List<SwpLogEntry>());
            Assert.Equal(SwpBusfieldVerdict.Ok, row.BusfieldVerdict);
        }

        [Fact]
        public void A_busfield_without_node_zero_warns_about_the_controller()
        {
            CsvDocument busfield = Csv(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;236;32880;0;0;0;0;0;0;0;0;#;x");
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(new SwpWorkspaceModel(), null, busfield, SwpPlantType.Msc, log);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("node 0"));

            CsvDocument withZero = Csv(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;0;32880;0;0;0;0;0;0;0;0;#;controller",
                "@;1;236;32880;0;0;0;0;0;0;0;0;#;x");
            var quiet = new List<SwpLogEntry>();
            SwpCoherence.Annotate(new SwpWorkspaceModel(), null, withZero, SwpPlantType.Msc, quiet);
            Assert.DoesNotContain(quiet, l => l.Message.Contains("node 0"));
        }

        [Fact]
        public void An_asi_master_without_the_ignore_error_flag_warns()
        {
            // Node 1703 names node 17 as its master (linked address, field 5); the master's
            // warning-on-fault is 0 — the ignore-error bit (2 or 3) is missing.
            CsvDocument busfield = Csv(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;17;3;0;0;0;0;0;0;0;0;#;master",
                "@;1;1703;3;0;17;1;3;0;0;0;0;#;slave");
            var log = new List<SwpLogEntry>();
            SwpCoherence.Annotate(new SwpWorkspaceModel(), null, busfield, SwpPlantType.Msc, log);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("AS-I master") && l.Message.Contains("17"));

            CsvDocument healthy = Csv(
                "#;h;h;h;h;h;h;h;h;h;h;h;#;h",
                "@;1;17;3;2;0;0;0;0;0;0;0;#;master",
                "@;1;1703;3;0;17;1;3;0;0;0;0;#;slave");
            var quiet = new List<SwpLogEntry>();
            SwpCoherence.Annotate(new SwpWorkspaceModel(), null, healthy, SwpPlantType.Msc, quiet);
            Assert.DoesNotContain(quiet, l => l.Message.Contains("AS-I master"));
        }

        // ---- IO-side rules in the diff (§3.3 / §3.8) ----------------------------------------

        [Fact]
        public void Forbidden_bits_and_dashed_locations_are_logged_by_the_diff()
        {
            var entries = new List<DiagnosticEntry>
            {
                new DiagnosticEntry { Mnemonic = "CABINETAL00131", Location = "+IO-LINK", CardSheet = "R1.S1.C1" },
                new DiagnosticEntry { Mnemonic = "CABINETAL00263", Location = "+IO-LINK", CardSheet = "R1.S1.C1" },
                new DiagnosticEntry { Mnemonic = "CABINETAL00330", Location = "+MC1.CC1", CardSheet = "R1.S1.C1" },
            };
            var log = new List<SwpLogEntry>();
            SwpDiff.Build(entries, new SwpWorkbookIndex(), new SwpDiffOptions(), log);

            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("bit 31"));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("bit 63"));
            // The dashed location is reported once, not once per row.
            Assert.Single(log, l => l.Message.Contains("contains '-'"));
        }

        // ---- SWP content checks (§3.1 / §3.3 / §3.8) ----------------------------------------

        private static SwpRowSnapshot Bit(int row, string tagType, int entry, int bit,
            string export = "x", string desc = "", string priority = "2", bool priorityIsText = false,
            string device = "")
        {
            return new SwpRowSnapshot
            {
                RowNumber = row,
                TagType = tagType,
                TableEntry = entry,
                Bit = bit,
                ExportB = export,
                Desc1T = desc,
                PriorityE = priority,
                PriorityIsText = priorityIsText,
                DeviceJ = device,
                BlockHeaderRow = 15,
            };
        }

        private static SwpWorkbookIndex Index(params SwpRowSnapshot[] rows)
        {
            var sheet = new SwpSheetIndex { SheetName = "CABINET", IsDiagnosticSheet = true };
            sheet.BitRows.AddRange(rows);
            var index = new SwpWorkbookIndex();
            index.Sheets["CABINET"] = sheet;
            return index;
        }

        [Fact]
        public void The_same_identity_twice_in_one_category_is_a_duplicate()
        {
            var log = new List<SwpLogEntry>();
            SwpContentChecks.Run(Index(
                Bit(17, "DB_CABINETS_DIAG.Alarm1", 0, 0, desc: "MOTOR FAULT", device: "-B1"),
                Bit(18, "DB_CABINETS_DIAG.Alarm1", 0, 1, desc: "MOTOR FAULT", device: "-B1"),
                Bit(51, "DB_CABINETS_DIAG.Warning1", 0, 6, desc: "OTHER", priority: "3"),
                Bit(60, "DB_CABINETS_DIAG.Alarm1", 0, 7, export: "0", desc: "MOTOR FAULT", device: "-B1")), log);

            SwpLogEntry finding = Assert.Single(log, l => l.Message.Contains("are the same"));
            Assert.Contains("17", finding.Message);
            Assert.Contains("18", finding.Message);
            Assert.Contains("alarm", finding.Message);
        }

        [Fact]
        public void Sixteen_safety_doors_sharing_one_sentence_are_not_duplicates()
        {
            // The field case: one block, sixteen doors, all reading "SAFETY DOOR OPEN" and told
            // apart by their DEVICE. The description alone was never the identity.
            var rows = new List<SwpRowSnapshot>();
            for (int i = 0; i < 16; i++)
                rows.Add(Bit(3205 + i, "DB_CABINETS_DIAG.Alarm2", 0, 32 + i,
                    desc: "SAFETY DOOR OPEN", device: "-B" + (i + 1)));

            var log = new List<SwpLogEntry>();
            SwpContentChecks.Run(Index(rows.ToArray()), log);

            Assert.DoesNotContain(log, l => l.Message.Contains("are the same"));
        }

        [Fact]
        public void The_same_string_as_an_alarm_and_as_a_warning_is_not_a_duplicate()
        {
            // The normal shape of a signal DECLASSED from alarm to warning. Reporting it buried
            // the real duplicates.
            var log = new List<SwpLogEntry>();
            SwpContentChecks.Run(Index(
                Bit(17, "DB_CABINETS_DIAG.Alarm1", 0, 0, desc: "CIRCUIT BREAKER TRIPPED 400V"),
                Bit(50, "DB_CABINETS_DIAG.Warning1", 0, 5,
                    desc: "CIRCUIT BREAKER TRIPPED 400V", priority: "3")), log);

            Assert.DoesNotContain(log, l => l.Message.Contains("both export"));
        }

        [Fact]
        public void A_forbidden_bit_is_reported_only_where_a_mnemonic_lands()
        {
            // An export tick on bit 31/63 is ignored by the plant; a MNEMONIC landing there is a
            // signal nobody will ever see. Only the second is worth a finding.
            SwpWorkbookIndex index = Index(
                Bit(48, "DB_CABINETS_DIAG.Alarm1", 0, 31, desc: "A"),
                Bit(49, "DB_CABINETS_DIAG.Alarm2", 0, 63, desc: "B"),
                Bit(50, "DB_CABINETS_DIAG.Alarm2", 0, 62, desc: "C"));

            var targeted = new List<SwpLogEntry>();
            SwpContentChecks.Run(index, targeted,
                new HashSet<string> { SwpContentChecks.RowKeyOf("CABINET", 48) });
            SwpLogEntry finding = Assert.Single(targeted, l => l.Message.Contains("forbidden"));
            Assert.Contains("row 48", finding.Message);
            Assert.Contains("takes a mnemonic", finding.Message);

            // Row 49 sits on bit 63 with its export ticked and no mnemonic — silent.
            var untargeted = new List<SwpLogEntry>();
            SwpContentChecks.Run(index, untargeted, new HashSet<string>());
            Assert.DoesNotContain(untargeted, l => l.Message.Contains("forbidden"));
        }

        [Theory]
        [InlineData("", false, "must be a number")]
        [InlineData("0", false, "must be a number")]
        [InlineData("abc", false, "must be a number")]
        [InlineData("2", true, "as TEXT")]
        public void Broken_priorities_on_exported_rows_warn(string priority, bool isText, string expected)
        {
            var log = new List<SwpLogEntry>();
            SwpContentChecks.Run(Index(
                Bit(17, "DB_CABINETS_DIAG.Alarm1", 0, 0, desc: "A", priority: priority, priorityIsText: isText)), log);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains(expected));
        }

        [Fact]
        public void Unconventional_priorities_are_informational_and_unexported_rows_are_ignored()
        {
            var log = new List<SwpLogEntry>();
            SwpContentChecks.Run(Index(
                Bit(17, "DB_CABINETS_DIAG.Alarm1", 0, 0, desc: "A", priority: "3"),
                Bit(18, "DB_CABINETS_DIAG.Warning1", 0, 1, export: "0", desc: "B", priority: "")), log);

            SwpLogEntry info = Assert.Single(log);
            Assert.Equal(SwpLogSeverity.Info, info.Severity);
            Assert.Contains("convention is 2", info.Message);
        }

        [Fact]
        public void Header_locations_with_a_dash_are_flagged()
        {
            var sheet = new SwpSheetIndex { SheetName = "CABINET", IsDiagnosticSheet = true };
            sheet.BlockHeaders.Add(new SwpBlockHeader { RowNumber = 15, LocationI = "+IO-LINK" });
            sheet.BlockHeaders.Add(new SwpBlockHeader { RowNumber = 181, LocationI = "+MC1.CC1" });
            var index = new SwpWorkbookIndex();
            index.Sheets["CABINET"] = sheet;

            var log = new List<SwpLogEntry>();
            SwpContentChecks.Run(index, log);
            SwpLogEntry finding = Assert.Single(log);
            Assert.Contains("+IO-LINK", finding.Message);
        }
    }
}
