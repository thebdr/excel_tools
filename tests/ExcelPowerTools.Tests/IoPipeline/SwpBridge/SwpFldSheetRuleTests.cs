using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The destination-sheet FLD rule: a regex on the SWP sheet name that overrides the workspace
    /// mismatch default, sitting BETWEEN the default and the per-row [Append FLD] / [Overwrite FLD]
    /// tick. It generalizes what was hardcoded as EndsWith("BELT"): whole sheets often want the
    /// other treatment, and ticking every row on them by hand is not a workflow.
    ///
    /// The FLD rule itself is untouched — a matched sheet still carries the functional unit,
    /// location and device the IO list holds; only the place they land changes.
    /// </summary>
    public class SwpFldSheetRuleTests
    {
        // ---- fixtures -----------------------------------------------------------------------
        // The same SAFETY DOOR row the FLD tests use: its mnemonic maps to the CABINET sheet, so
        // "CABINET" is the pattern that HITS and "BELT$" the one that MISSES.

        private static DiagnosticEntry Door() => new DiagnosticEntry
        {
            CardSheet = "R1.S1.C1",
            SourceRow = 143,
            Mnemonic = "CABINETAL01914",
            RawMnemonic = "CABINETAL01914",
            Node = "235",
            Offset = "32",
            Bit = "5",
            FUnit = "=S1",
            Location = "+SG14",
            Device = "-B1",
            TsRef = "TS_009",
            Desc1A = "SAFETY DOOR OPEN",
            Desc2A = "LA PORTE DE SECURITE EST OUVERTE",
        };

        private static SwpRowSnapshot DoorRow() => new SwpRowSnapshot
        {
            RowNumber = 2300,
            TagType = "DB_CABINETS_DIAG.Alarm1",
            TableEntry = 19,
            Bit = 14,
            ExportB = "x",
            BlockFUnit = "=S1",
            BlockLocation = "+SAFETYDOORS",
            BlockLabel = "=S1+SAFETYDOORS",
            BlockHeaderRow = 2251,
            HasCellH = true,
            CachedH = "=S1",
            FormulaH = "H2251",
            HasCellI = true,
            CachedI = "+SAFETYDOORS",
            FormulaI = "I2251",
            Desc1T = "AL14",
            Desc2U = "AL14",
            FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57&T2300,\"%1\",F2251)",
            FormulaR = "SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57&U2300,\"%1\",F2251)",
        };

        private static SwpWorkbookIndex Index()
        {
            SwpRowSnapshot row = DoorRow();
            var index = new SwpWorkbookIndex
            {
                Sheets = new Dictionary<string, SwpSheetIndex>(System.StringComparer.OrdinalIgnoreCase),
            };
            var sheet = new SwpSheetIndex { SheetName = "CABINET" };
            sheet.RowsByKey[row.RowKey] = row;
            sheet.BitRows.Add(row);
            index.Sheets["CABINET"] = sheet;
            return index;
        }

        private static SwpSpineRow Build(SwpMismatchMode workspaceDefault, string pattern,
            SwpMismatchMode ruleMode, SwpMismatchMode? rowMode = null, List<SwpLogEntry> log = null)
        {
            var options = new SwpDiffOptions
            {
                PlantType = SwpPlantType.Msc,
                MismatchMode = workspaceDefault,
                FldSheetRule = SwpFldSheetRule.Create(pattern, ruleMode),
            };
            if (rowMode.HasValue)
                options.Overrides = new Dictionary<string, SwpRowOverride>
                {
                    ["CABINETAL01914"] = new SwpRowOverride
                    {
                        Mnemonic = "CABINETAL01914",
                        FldMode = rowMode,
                    },
                };
            return SwpDiff.Build(new[] { Door() }, Index(), options,
                log ?? new List<SwpLogEntry>()).Rows[0];
        }

        private static SwpProposedEdit Edit(SwpSpineRow row, SwpField field) =>
            row.Edits.FirstOrDefault(e => e.Field == field);

        // ---- the rule in isolation ----------------------------------------------------------

        [Fact]
        public void An_empty_pattern_is_no_exception_at_all()
        {
            SwpFldSheetRule rule = SwpFldSheetRule.Create("   ", SwpMismatchMode.Overwrite);

            Assert.False(rule.IsActive);
            Assert.Equal("", rule.Error);
            Assert.Null(rule.ModeFor("CABINET"));
        }

        [Fact]
        public void A_matching_sheet_takes_the_rules_mode_and_a_missed_one_takes_none()
        {
            SwpFldSheetRule rule = SwpFldSheetRule.Create("BELT$", SwpMismatchMode.Overwrite);

            Assert.True(rule.IsActive);
            Assert.Equal(SwpMismatchMode.Overwrite, rule.ModeFor("UP01_BELT"));
            Assert.Null(rule.ModeFor("CABINET"));
        }

        [Fact]
        public void The_pattern_matches_case_insensitively()
        {
            // SWP sheet names are upper-case by convention; a rule that silently missed on case
            // would hand a whole sheet the treatment the workspace did not ask for.
            SwpFldSheetRule rule = SwpFldSheetRule.Create("belt$", SwpMismatchMode.Overwrite);

            Assert.Equal(SwpMismatchMode.Overwrite, rule.ModeFor("UP01_BELT"));
        }

        [Fact]
        public void A_pattern_that_does_not_parse_is_inert_and_says_why()
        {
            SwpFldSheetRule rule = SwpFldSheetRule.Create("UP0[12", SwpMismatchMode.Overwrite);

            Assert.False(rule.IsActive);
            Assert.NotEqual("", rule.Error);
            Assert.Null(rule.ModeFor("UP01_BELT"));   // inert, never throwing
            Assert.Equal("UP0[12", rule.Pattern);     // kept, so the form can show it back
        }

        [Fact]
        public void A_null_or_empty_sheet_name_never_matches()
        {
            SwpFldSheetRule rule = SwpFldSheetRule.Create(".*", SwpMismatchMode.Overwrite);

            Assert.Null(rule.ModeFor(null));
            Assert.Null(rule.ModeFor(""));
        }

        // ---- the cascade: row tick > sheet rule > workspace default --------------------------

        [Fact]
        public void A_matched_sheet_leaves_the_workspace_default_and_writes_its_own_cells()
        {
            // Default appends; the rule names CABINET for the overwrite, so the identity must
            // land in H/I rather than in the descriptions.
            SwpSpineRow row = Build(SwpMismatchMode.AppendToDescription, "CABINET",
                SwpMismatchMode.Overwrite);

            Assert.True(row.FldOverwriteMode);
            Assert.True(row.FldModeFromSheetRule);
            // Only the location differs in this fixture (both sides say "=S1"), so I carries the
            // identity and the descriptions stay bare — the overwrite signature.
            Assert.Equal("+SG14", Edit(row, SwpField.LocationI).Proposed);
            Assert.Equal("-B1", Edit(row, SwpField.DeviceJ).Proposed);
            Assert.Equal("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed.Trim());
            Assert.Null(Edit(row, SwpField.FUnitH));   // H already agrees
        }

        [Fact]
        public void A_sheet_the_rule_misses_still_follows_the_workspace_default()
        {
            SwpSpineRow row = Build(SwpMismatchMode.AppendToDescription, "BELT$",
                SwpMismatchMode.Overwrite);

            Assert.False(row.FldOverwriteMode);
            Assert.False(row.FldModeFromSheetRule);
            // The append signature: the row's own cells are untouched and the identity rides the
            // descriptions instead.
            Assert.Null(Edit(row, SwpField.LocationI));
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void A_per_row_tick_beats_the_sheet_rule()
        {
            // The rule would overwrite this sheet; the row was ticked back to the append. The
            // most specific decision wins, and the row no longer counts as the rule doing it.
            SwpSpineRow row = Build(SwpMismatchMode.AppendToDescription, "CABINET",
                SwpMismatchMode.Overwrite, rowMode: SwpMismatchMode.AppendToDescription);

            Assert.False(row.FldOverwriteMode);
            Assert.False(row.FldModeFromSheetRule);
            Assert.Null(Edit(row, SwpField.LocationI));
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void A_per_row_tick_can_also_pull_a_row_onto_the_rules_mode()
        {
            SwpSpineRow row = Build(SwpMismatchMode.AppendToDescription, "BELT$",
                SwpMismatchMode.Overwrite, rowMode: SwpMismatchMode.Overwrite);

            Assert.True(row.FldOverwriteMode);
            Assert.False(row.FldModeFromSheetRule);   // the tick decided it, not the rule
            Assert.Equal("+SG14", Edit(row, SwpField.LocationI).Proposed);
        }

        [Fact]
        public void The_rule_can_also_pull_a_sheet_back_to_the_append()
        {
            // The mirror case: an overwrite workspace with an append exception.
            SwpSpineRow row = Build(SwpMismatchMode.Overwrite, "CABINET",
                SwpMismatchMode.AppendToDescription);

            Assert.False(row.FldOverwriteMode);
            Assert.True(row.FldModeFromSheetRule);
            Assert.Null(Edit(row, SwpField.LocationI));
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);
        }

        // ---- reporting ----------------------------------------------------------------------

        [Fact]
        public void An_active_rule_reports_how_many_rows_it_moved()
        {
            var log = new List<SwpLogEntry>();
            Build(SwpMismatchMode.AppendToDescription, "CABINET", SwpMismatchMode.Overwrite, log: log);

            SwpLogEntry line = log.FirstOrDefault(e => e.Message.Contains("destination-sheet FLD rule"));
            Assert.NotNull(line);
            Assert.Equal(SwpLogSeverity.Info, line.Severity);
            Assert.Contains("1 row(s)", line.Message);
            Assert.Contains("overwrites", line.Message);
        }

        [Fact]
        public void A_rule_that_does_not_parse_warns_and_falls_back_to_the_default()
        {
            var log = new List<SwpLogEntry>();
            SwpSpineRow row = Build(SwpMismatchMode.AppendToDescription, "UP0[12",
                SwpMismatchMode.Overwrite, log: log);

            Assert.False(row.FldOverwriteMode);          // the default stood
            Assert.False(row.FldModeFromSheetRule);
            SwpLogEntry warning = log.FirstOrDefault(e =>
                e.Severity == SwpLogSeverity.Warning && e.Message.Contains("does not parse"));
            Assert.NotNull(warning);
        }

        [Fact]
        public void No_rule_configured_says_nothing_at_all()
        {
            var log = new List<SwpLogEntry>();
            Build(SwpMismatchMode.AppendToDescription, "", SwpMismatchMode.Overwrite, log: log);

            Assert.DoesNotContain(log, e => e.Message.Contains("destination-sheet FLD rule"));
        }

        // ---- persistence --------------------------------------------------------------------

        [Fact]
        public void The_rule_survives_a_sidecar_round_trip_as_the_mode_it_chose()
        {
            var document = new SwpSidecarDocument();
            document.Settings.MismatchMode = SwpMismatchMode.Overwrite;
            document.Settings.FldSheetRulePattern = "BELT$";
            document.Settings.FldSheetRuleMode = SwpMismatchMode.AppendToDescription;

            Assert.True(SwpSidecarJson.TryParse(SwpSidecarJson.Serialize(document),
                out SwpSidecarDocument read, out string error), error);

            Assert.Equal("BELT$", read.Settings.FldSheetRulePattern);
            // Stored as the mode, not as a flip: changing MismatchMode later cannot reverse it.
            Assert.Equal(SwpMismatchMode.AppendToDescription, read.Settings.FldSheetRuleMode);
        }

        [Fact]
        public void A_sidecar_written_before_the_rule_existed_carries_no_exception()
        {
            // A faithful old sidecar: everything a current one carries, minus the two keys the
            // rule added. Hand-rolling a minimal document would test the parser, not the upgrade.
            var document = new SwpSidecarDocument();
            document.Settings.MismatchMode = SwpMismatchMode.Overwrite;
            string json = System.Text.RegularExpressions.Regex.Replace(
                SwpSidecarJson.Serialize(document),
                ",\\s*\"fldSheetRule(Pattern|Mode)\"\\s*:\\s*\"[^\"]*\"", "");
            Assert.DoesNotContain("fldSheetRule", json);

            Assert.True(SwpSidecarJson.TryParse(json, out SwpSidecarDocument read, out string error), error);

            Assert.Equal("", read.Settings.FldSheetRulePattern);
            Assert.False(SwpFldSheetRule.Create(read.Settings.FldSheetRulePattern,
                read.Settings.FldSheetRuleMode).IsActive);
        }
    }
}
