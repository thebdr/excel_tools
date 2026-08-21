using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpDiffTests
    {
        // ---- fixture helpers ----------------------------------------------------------------

        private static DiagnosticEntry CabinetEntry(string mnemonic = "CABINETAL00100")
        {
            return new DiagnosticEntry
            {
                CardSheet = "R1.S1.C1",
                SourceRow = 10,
                Mnemonic = mnemonic,
                RawMnemonic = mnemonic,
                Node = "80",
                Offset = "3",
                Bit = "1",
                FUnit = "=S1",
                Location = "+MC1.CC1",
                Device = "-F22001",
                TsRef = "TS_001",
                Desc1A = "CIRCUIT BREAKER",
                Desc1B = "TRIPPED",
                Desc2A = "DISJONCTEUR",
                Desc2B = "OUVERT",
            };
        }

        /// <summary>The CABINET row CABINETAL00100 maps to: DB_CABINETS_DIAG.Alarm1&gt;1&gt;0.</summary>
        private static SwpRowSnapshot CabinetSnapshot()
        {
            return new SwpRowSnapshot
            {
                RowNumber = 216,
                TagType = "DB_CABINETS_DIAG.Alarm1",
                TableEntry = 1,
                Bit = 0,
                ExportB = "0",
                BlockFUnit = "=S1",
                BlockLocation = "+MC1.CC1",
                BlockLabel = "=S1+MC1.CC1",
                BlockHeaderRow = 181,
                FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57&T216,\"%1\",F181)",
                FormulaR = "SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57&U216,\"%1\",F181)",
                FormulaHShared = true,   // healthy FU/Location cells, the real rows' default
                FormulaIShared = true,
            };
        }

        private static SwpWorkbookIndex Index(params (string Sheet, SwpRowSnapshot Row)[] rows)
        {
            var index = new SwpWorkbookIndex
            {
                Sheets = new Dictionary<string, SwpSheetIndex>(System.StringComparer.OrdinalIgnoreCase),
            };
            foreach (var (sheetName, row) in rows)
            {
                if (!index.Sheets.TryGetValue(sheetName, out SwpSheetIndex sheet))
                    index.Sheets[sheetName] = sheet = new SwpSheetIndex { SheetName = sheetName };
                sheet.RowsByKey[row.RowKey] = row;
                sheet.BitRows.Add(row);
            }
            return index;
        }

        private static SwpWorkspaceModel Build(DiagnosticEntry entry, SwpWorkbookIndex index,
            SwpDiffOptions options = null, List<SwpLogEntry> log = null)
        {
            return SwpDiff.Build(new[] { entry }, index,
                options ?? new SwpDiffOptions { PlantType = SwpPlantType.Msc },
                log ?? new List<SwpLogEntry>());
        }

        private static SwpProposedEdit EditFor(SwpSpineRow row, SwpField field) =>
            row.Edits.FirstOrDefault(e => e.Field == field);

        // ---- the preview's recalc simulation ------------------------------------------------

        [Fact]
        public void The_simulated_recalc_swaps_the_trailing_description()
        {
            // The exact field case of 2026-08-21: Q = template & T; the cached final ends with
            // the CURRENT T, and the simulation must replace it — never concatenate old + new.
            string simulated = SwpDiff.SimulateRecalculatedFinal(
                "CABINET =S1+MC1.CC1 - FIELD BUS MODULE FAILURE [n001 - 192.168.1.1]",
                "FIELD BUS MODULE FAILURE [n001 - 192.168.1.1]",
                "FIELD BUS MODULE FAILURE [n0001-mc1-cc1-a1001 - 192.168.1.1]",
                formulaReferencesDesc: true, formulaAppendPending: false, wholesalePending: false,
                SwpPlantType.Msc, "CABINET", "Cabinet", "=S1+MC1.CC1");
            Assert.Equal(
                "CABINET =S1+MC1.CC1 - FIELD BUS MODULE FAILURE [n0001-mc1-cc1-a1001 - 192.168.1.1]",
                simulated);
        }

        [Fact]
        public void An_empty_current_description_appends_to_the_bare_prefix()
        {
            Assert.Equal("CABINET =S1+X - AL05",
                SwpDiff.SimulateRecalculatedFinal("CABINET =S1+X - ", "", "AL05",
                    true, false, false, SwpPlantType.Msc, "CABINET", "Cabinet", "=S1+X"));
        }

        [Fact]
        public void A_template_only_formula_shows_the_description_only_through_the_append()
        {
            // The CHUTE field case of 2026-08-21: Q is template-only (no T reference), so the
            // pending &" "&T append adds the description EXACTLY ONCE …
            Assert.Equal("CHUTE =S1+CH117 - AS-i SLAVE FAULT  =S1+CH116-K1",
                SwpDiff.SimulateRecalculatedFinal("CHUTE =S1+CH117 - AS-i SLAVE FAULT", "",
                    " =S1+CH116-K1",
                    formulaReferencesDesc: false, formulaAppendPending: true, wholesalePending: false,
                    SwpPlantType.Msc, "CHUTE", "Chute", "=S1+CH117"));

            // … and WITHOUT the append, a description edit does not change the final text.
            Assert.Equal("CHUTE =S1+CH117 - AS-i SLAVE FAULT",
                SwpDiff.SimulateRecalculatedFinal("CHUTE =S1+CH117 - AS-i SLAVE FAULT", "",
                    " =S1+CH116-K1",
                    formulaReferencesDesc: false, formulaAppendPending: false, wholesalePending: false,
                    SwpPlantType.Msc, "CHUTE", "Chute", "=S1+CH117"));
        }

        [Fact]
        public void A_pending_wholesale_rewrite_composes_from_scratch()
        {
            Assert.Equal("CABINET =S1+MC1.CC1 - NEW TEXT",
                SwpDiff.SimulateRecalculatedFinal("whatever was cached", "OLD", "NEW TEXT",
                    true, false, wholesalePending: true,
                    SwpPlantType.Msc, "CABINET", "Cabinet", "=S1+MC1.CC1"));
            Assert.Equal("UPSTREAM Upstream 3 - CABINET =S3+Q1 - NEW TEXT",
                SwpDiff.SimulateRecalculatedFinal("", "", "NEW TEXT",
                    true, false, wholesalePending: true,
                    SwpPlantType.Ups, "UP03_CABINET", "Upstream 3", "=S3+Q1"));
        }

        // ---- matching and proposals ---------------------------------------------------------

        [Fact]
        public void The_glossary_ts_ref_reaches_the_L_proposal_with_a_warning()
        {
            // The IO row's Description PERMANENT part is "CIRCUIT BREAKER" (the variable part
            // "TRIPPED" does not key the glossary); the glossary owns that phrase with card
            // TS_050 while the IO list says TS_001 — the glossary wins on L.
            var log = new List<SwpLogEntry>();
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    GlossaryPairs = new List<SwpGlossaryPair>
                    {
                        new SwpGlossaryPair { En = "CIRCUIT BREAKER", Translation = "DISJUNTOR", TsRef = "TS_050" },
                    },
                }, log);

            SwpSpineRow row = Assert.Single(model.Rows);
            Assert.Equal("TS_050", EditFor(row, SwpField.TsRefL).Proposed);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("TS_001") && l.Message.Contains("TS_050") &&
                l.Message.Contains("glossary wins"));
        }

        [Fact]
        public void An_unofficial_bus_description_is_flagged_in_the_log()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.ProfinetName = "n004-z";
            entry.ProfinetIp = "10.0.0.4";
            entry.Desc1A = "HOMEMADE MODULE TEXT";
            entry.Desc1B = "";

            var log = new List<SwpLogEntry>();
            Build(entry, Index(("CABINET", CabinetSnapshot())), new SwpDiffOptions
            {
                PlantType = SwpPlantType.Msc,
                GlossaryPairs = new List<SwpGlossaryPair>
                {
                    new SwpGlossaryPair { En = "FIELD BUS MODULE FAILURE", Translation = "AVARIA DO MÓDULO" },
                },
            }, log);

            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("HOMEMADE MODULE TEXT") &&
                l.Message.Contains("FIELD BUS MODULE FAILURE"));
        }

        [Fact]
        public void A_fresh_row_proposes_every_field()
        {
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())));

            SwpSpineRow row = Assert.Single(model.Rows);
            Assert.Equal(SwpRowVerdict.Pending, row.Verdict);
            Assert.Equal("X", EditFor(row, SwpField.ExportB).Proposed);   // rules §3.1 format
            Assert.Equal("80.3.1", EditFor(row, SwpField.AddressG).Proposed);
            Assert.Equal("-F22001", EditFor(row, SwpField.DeviceJ).Proposed);
            Assert.Equal("TS_001", EditFor(row, SwpField.TsRefL).Proposed);
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc1T).Proposed);
            Assert.Equal("DISJONCTEUR OUVERT", EditFor(row, SwpField.Desc2U).Proposed);
            Assert.Null(EditFor(row, SwpField.FormulaQ));   // no mismatch → formulas untouched
            Assert.Equal(216, row.Edits[0].RowNumber);
        }

        [Fact]
        public void An_up_to_date_row_proposes_nothing()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.ExportB = "X";
            snapshot.AddressG = "80.3.1";
            snapshot.DeviceJ = "-F22001";
            snapshot.TsRefL = "TS_001";
            snapshot.Desc1T = "CIRCUIT BREAKER TRIPPED";
            snapshot.Desc2U = "DISJONCTEUR OUVERT";

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)));

            SwpSpineRow row = Assert.Single(model.Rows);
            Assert.Equal(SwpRowVerdict.UpToDate, row.Verdict);
            Assert.Empty(row.Edits);
        }

        [Fact]
        public void Trailing_space_differences_do_not_churn()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.ExportB = "x";                       // legacy lowercase — must not churn vs "X"
            snapshot.AddressG = "80.3.1";
            snapshot.DeviceJ = "-F22001";
            snapshot.TsRefL = "TS_001";
            snapshot.Desc1T = "CIRCUIT BREAKER TRIPPED ";  // legacy files carry trailing spaces
            snapshot.Desc2U = "DISJONCTEUR OUVERT";

            DiagnosticEntry entry = CabinetEntry();
            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)));
            Assert.Equal(SwpRowVerdict.UpToDate, model.Rows[0].Verdict);
        }

        // ---- mismatch machinery -------------------------------------------------------------

        [Fact]
        public void A_label_mismatch_appends_blanks_and_extends()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.DeviceJ = "-F22001";   // previously written — the mismatch must blank it
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)), log: log);

            SwpSpineRow row = model.Rows[0];
            Assert.Equal(SwpRowVerdict.LabelMismatch, row.Verdict);
            Assert.False(row.FUnitMismatch);
            Assert.True(row.LocationMismatch);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("location differ") && l.Message.Contains("[Mismatch OK]"));

            Assert.Equal("CIRCUIT BREAKER TRIPPED =S1+MC1.CC1-F22001",
                EditFor(row, SwpField.Desc1T).Proposed);
            Assert.Equal("", EditFor(row, SwpField.DeviceJ).Proposed);

            // The standard formulas already reference T216/U216 — the label addition flows
            // through the description cell, so NO extension may be proposed (extending would
            // print the description twice; seen in the field on +EMERGENCY rows).
            Assert.Null(EditFor(row, SwpField.FormulaQ));
            Assert.Null(EditFor(row, SwpField.FormulaR));
        }

        [Fact]
        public void A_fieldbus_module_gets_the_regular_mismatch_treatment()
        {
            // User ruling: the profinet specialness ends at the description composition —
            // from there the row is a regular alarm, so a location mismatch blanks the device
            // and appends the compact label exactly like any other row.
            DiagnosticEntry entry = CabinetEntry();
            entry.IsBusModule = true;
            entry.Node = "66";
            entry.Offset = "";
            entry.Bit = "";
            entry.ProfinetName = "n066";
            entry.ProfinetIp = "192.168.1.66";
            entry.Location = "+SOMEWHERE.ELSE";
            entry.Device = "-XN1";

            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.DeviceJ = "-XN1";

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)));

            SwpSpineRow row = model.Rows[0];
            Assert.True(row.LocationMismatch);
            Assert.Equal(SwpRowVerdict.LabelMismatch, row.Verdict);
            Assert.Equal("", EditFor(row, SwpField.DeviceJ).Proposed);
            Assert.Equal("FIELD BUS MODULE FAILURE [n066 - 192.168.1.66] =S1+SOMEWHERE.ELSE-XN1",
                EditFor(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void A_template_only_formula_still_gets_the_mismatch_extension()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",F181)";   // no T reference
            snapshot.FormulaR = "SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57,\"%1\",F181)";

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)));

            SwpSpineRow row = model.Rows[0];
            Assert.EndsWith("&\" \"&T216", EditFor(row, SwpField.FormulaQ).Proposed);
            Assert.EndsWith("&\" \"&U216", EditFor(row, SwpField.FormulaR).Proposed);
        }

        [Fact]
        public void Overwrite_mode_still_warns_but_writes_the_io_values_plainly()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockFUnit = "=S9";
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    MismatchMode = SwpMismatchMode.Overwrite,
                }, log);

            SwpSpineRow row = model.Rows[0];
            Assert.Equal(SwpRowVerdict.LabelMismatch, row.Verdict);   // the warning always stands
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("functional unit") && l.Message.Contains("[Mismatch OK]"));
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc1T).Proposed);
            Assert.Equal("-F22001", EditFor(row, SwpField.DeviceJ).Proposed);
            Assert.Null(EditFor(row, SwpField.FormulaQ));
        }

        [Fact]
        public void Mismatch_ok_silences_the_warning_but_the_treatment_still_applies()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockFUnit = "=S9";
            snapshot.BlockLocation = "+ELSEWHERE";
            snapshot.DeviceJ = "-F22001";
            var overrides = new Dictionary<string, SwpRowOverride>
            {
                { "CABINETAL00100", new SwpRowOverride { Mnemonic = "CABINETAL00100", SkipLabelCheck = true } },
            };
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, Overrides = overrides }, log);

            SwpSpineRow row = model.Rows[0];
            Assert.True(row.FUnitMismatch && row.LocationMismatch);   // the difference is real…
            Assert.True(row.MismatchSuppressed);                      // …and marked deliberate
            Assert.NotEqual(SwpRowVerdict.LabelMismatch, row.Verdict);
            Assert.DoesNotContain(log, l => l.Message.Contains("differ from the block"));
            // Append mode (the default) still treats the row: label appended, device blanked.
            Assert.EndsWith("=S1+MC1.CC1-F22001", EditFor(row, SwpField.Desc1T).Proposed);
            Assert.Equal("", EditFor(row, SwpField.DeviceJ).Proposed);
        }

        [Fact]
        public void A_broken_fu_location_formula_reworks_the_warning()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockFUnit = "=S9";
            snapshot.FormulaH = "";          // pasted value — the cleaner queues its restoration
            snapshot.FormulaHShared = false;
            snapshot.FormulaI = "I181";
            var log = new List<SwpLogEntry>();

            Build(CabinetEntry(), Index(("CABINET", snapshot)), log: log);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("restoration queued") &&
                l.Message.Contains("restored block values"));
        }

        // ---- the T.S. Ref rules -------------------------------------------------------------

        [Fact]
        public void An_empty_ts_ref_never_blanks_an_existing_one()
        {
            // The SCADA opens the troubleshooting manual from this code — a wrong section
            // beats no action, so an existing value survives an empty composition.
            DiagnosticEntry entry = CabinetEntry();
            entry.TsRef = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.TsRefL = "TS_009";

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)));
            SwpSpineRow row = model.Rows[0];
            Assert.Null(EditFor(row, SwpField.TsRefL));
            Assert.False(row.TsRefMissing);   // the SWP holds one — nothing is missing
        }

        [Fact]
        public void An_io_alarm_without_any_ts_ref_anywhere_goes_red()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.TsRef = "";
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", CabinetSnapshot())), log: log);
            SwpSpineRow row = model.Rows[0];
            Assert.True(row.TsRefMissing);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("mandatory"));

            // A generated row is not an IO-list alarm — never flagged.
            DiagnosticEntry generated = CabinetEntry();
            generated.TsRef = "";
            generated.GeneratedBy = "AP";
            SwpWorkspaceModel generatedModel = Build(generated, Index(("CABINET", CabinetSnapshot())));
            Assert.False(generatedModel.Rows[0].TsRefMissing);
        }

        [Fact]
        public void A_glossary_card_satisfies_the_mandatory_ts_ref()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.TsRef = "";

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    GlossaryPairs = new List<SwpGlossaryPair>
                    {
                        new SwpGlossaryPair { En = "CIRCUIT BREAKER", Translation = "DISJUNTOR", TsRef = "TS_050" },
                    },
                });

            SwpSpineRow row = model.Rows[0];
            Assert.False(row.TsRefMissing);
            Assert.Equal("TS_050", EditFor(row, SwpField.TsRefL).Proposed);
        }

        // ---- belt rules ---------------------------------------------------------------------

        private static DiagnosticEntry BeltEntry() => new DiagnosticEntry
        {
            CardSheet = "R1.S1.C1",
            SourceRow = 3,
            Mnemonic = "ALLB_1_001_50",
            Node = "112",
            Offset = "7",
            Bit = "0",
            Device = "-B1",
            Location = "+UL1101.01",
            FUnit = "=S1",
            Desc1A = "Sensor Fault #1",
        };

        private static SwpRowSnapshot BeltSnapshot() => new SwpRowSnapshot
        {
            RowNumber = 120,
            TagType = "DB_UP01_BELT.Alarm2",
            TableEntry = 0,
            Bit = 50,
            BlockFUnit = "=S1",
            BlockLocation = "+SOMEWHERE.ELSE",   // belts legitimately differ
        };

        [Fact]
        public void Belt_alarms_keep_their_device_and_warn_like_everyone()
        {
            var log = new List<SwpLogEntry>();
            SwpWorkspaceModel model = Build(BeltEntry(), Index(("UP01_BELT", BeltSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Ups }, log);

            SwpSpineRow row = model.Rows[0];
            Assert.Null(EditFor(row, SwpField.DeviceJ));   // the belt device rule is untouched
            Assert.True(row.LocationMismatch);
            // The policy redesign: EVERY unsuppressed mismatch warns (belts included) — the
            // per-row [Mismatch OK] flag is the way to silence a deliberate one.
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("location differ"));
            // The mismatch still shapes the descriptions, exactly like the VBA.
            Assert.Contains("=S1+UL1101.01-B1", EditFor(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void Force_electrical_ids_writes_the_device_even_on_belts()
        {
            var overrides = new Dictionary<string, SwpRowOverride>
            {
                { "ALLB_1_001_50", new SwpRowOverride { Mnemonic = "ALLB_1_001_50", ForceElectricalIds = true } },
            };
            SwpRowSnapshot snapshot = BeltSnapshot();
            snapshot.BlockLocation = "+UL1101.01";   // no mismatch → the device value flows

            SwpWorkspaceModel model = Build(BeltEntry(), Index(("UP01_BELT", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Ups, Overrides = overrides });

            Assert.Equal("-B1", EditFor(model.Rows[0], SwpField.DeviceJ).Proposed);
        }

        // ---- trust-config-plant -------------------------------------------------------------

        [Fact]
        public void Generated_rows_write_only_the_addition_when_templates_are_trusted()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.GeneratedBy = "AP";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER";

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, TrustConfigPlant = true });

            SwpSpineRow row = model.Rows[0];
            Assert.Equal(" =S1+MC1.CC1-F22001", EditFor(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void Module_typed_rows_are_trusted_too()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.TypeCell = "PA";
            SwpRowSnapshot snapshot = CabinetSnapshot();

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, TrustConfigPlant = true });

            // No mismatch → the addition is empty → nothing to write into T at all.
            Assert.Null(EditFor(model.Rows[0], SwpField.Desc1T));
        }

        // ---- wholesale formula rewrite ------------------------------------------------------

        [Fact]
        public void Wholesale_mode_rewrites_both_description_formulas()
        {
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, UpdateDescriptionFormulas = true });

            SwpSpineRow row = model.Rows[0];
            Assert.Equal("CONCAT(UPPER($B$6),\" \",$F181,\" - \",T216)",
                EditFor(row, SwpField.FormulaQ).Proposed);
            Assert.Equal("CONCAT(UPPER($B$6),\" \",$F181,\" - \",U216)",
                EditFor(row, SwpField.FormulaR).Proposed);
        }

        [Fact]
        public void Upstream_wholesale_formula_carries_the_sheet_type()
        {
            Assert.Equal("CONCAT(\"UPSTREAM \",$B$6,\" - BELT \",$F100,\" - \",T120)",
                SwpDiff.WholesaleDescriptionFormula(SwpPlantType.Ups, "UP01_BELT", 100, 120, "T"));
        }

        [Fact]
        public void Shared_formulas_are_refused_with_a_warning()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            // Template-only formulas (no T/U reference) — the shapes the extension applies to.
            snapshot.FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",F181)";
            snapshot.FormulaR = "SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57,\"%1\",F181)";
            snapshot.FormulaQShared = true;
            snapshot.BlockLocation = "+OTHER";
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)), log: log);

            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.NotNull(EditFor(model.Rows[0], SwpField.FormulaR));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("shared"));
        }

        [Fact]
        public void The_formula_append_is_idempotent()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER";
            snapshot.FormulaQ += "&\" \"&T216";   // a previous run already appended
            snapshot.FormulaR += "&\" \"&U216";

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)));

            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaR));
        }

        // ---- verdicts, duplicates, tallies --------------------------------------------------

        [Fact]
        public void Missing_rows_sheets_and_invalid_mnemonics_get_their_verdicts()
        {
            var log = new List<SwpLogEntry>();
            SwpWorkspaceModel model = SwpDiff.Build(new[]
            {
                CabinetEntry(),                              // row exists → Pending
                CabinetEntry("CABINETAL09900"),              // no such row → RowMissing
                new DiagnosticEntry { Mnemonic = "CLASSDIAGAL00620000107" },  // CHUTE sheet absent
                new DiagnosticEntry { Mnemonic = "GARBAGE" },
            }, Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc }, log);

            Assert.Equal(SwpRowVerdict.Pending, model.Rows[0].Verdict);
            Assert.Equal(SwpRowVerdict.RowMissing, model.Rows[1].Verdict);
            Assert.Equal(SwpRowVerdict.SheetMissing, model.Rows[2].Verdict);
            Assert.Equal(SwpRowVerdict.Invalid, model.Rows[3].Verdict);
            Assert.Contains(log, l => l.Message.Contains("tag not found"));
        }

        [Fact]
        public void Duplicate_targets_are_demoted_and_queue_nothing()
        {
            var log = new List<SwpLogEntry>();
            SwpWorkspaceModel model = SwpDiff.Build(new[]
            {
                CabinetEntry(),
                CabinetEntry(),   // the same coordinate twice
            }, Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc }, log);

            Assert.All(model.Rows, r => Assert.Equal(SwpRowVerdict.DuplicateTarget, r.Verdict));
            Assert.Empty(model.Edits);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("2 rows target"));
        }

        [Fact]
        public void Tallies_group_by_target_sheet()
        {
            SwpWorkspaceModel model = SwpDiff.Build(new[]
            {
                CabinetEntry(),
                new DiagnosticEntry { Mnemonic = "GARBAGE" },
            }, Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc }, new List<SwpLogEntry>());

            SwpSheetTally unmapped = model.Tallies.First(t => t.SheetName == "(unmapped)");
            SwpSheetTally cabinet = model.Tallies.First(t => t.SheetName == "CABINET");
            Assert.Equal(1, unmapped.Invalid);
            Assert.Equal(1, cabinet.Pending);
            Assert.Equal(1, cabinet.Total);
        }

        [Fact]
        public void The_swp_db_index_override_reroutes_the_coordinate()
        {
            var overrides = new Dictionary<string, SwpRowOverride>
            {
                { "CABINETAL00100", new SwpRowOverride { Mnemonic = "CABINETAL00100", SwpDbIndex = 7 } },
            };
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.TableEntry = 7;

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, Overrides = overrides });

            Assert.Equal(SwpRowVerdict.Pending, model.Rows[0].Verdict);
            Assert.Equal(7, model.Rows[0].Mapped.Coordinate.TableEntry);
        }
    }
}
