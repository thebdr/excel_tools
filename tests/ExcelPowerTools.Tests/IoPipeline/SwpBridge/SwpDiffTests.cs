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
                formulaReferencesDesc: true, formulaAppendPending: false, rebuildPending: false,
                "Cabinet", "=S1+MC1.CC1");
            Assert.Equal(
                "CABINET =S1+MC1.CC1 - FIELD BUS MODULE FAILURE [n0001-mc1-cc1-a1001 - 192.168.1.1]",
                simulated);
        }

        [Fact]
        public void An_empty_current_description_appends_to_the_bare_prefix()
        {
            Assert.Equal("CABINET =S1+X - AL05",
                SwpDiff.SimulateRecalculatedFinal("CABINET =S1+X - ", "", "AL05",
                    true, false, false, "Cabinet", "=S1+X"));
        }

        [Fact]
        public void A_template_only_formula_shows_the_description_only_through_the_append()
        {
            // The CHUTE field case of 2026-08-21: Q is template-only (no T reference), so the
            // pending &" "&T append adds the description EXACTLY ONCE …
            Assert.Equal("CHUTE =S1+CH117 - AS-i SLAVE FAULT  =S1+CH116-K1",
                SwpDiff.SimulateRecalculatedFinal("CHUTE =S1+CH117 - AS-i SLAVE FAULT", "",
                    " =S1+CH116-K1",
                    formulaReferencesDesc: false, formulaAppendPending: true, rebuildPending: false,
                    "Chute", "=S1+CH117"));

            // … and WITHOUT the append, a description edit does not change the final text.
            Assert.Equal("CHUTE =S1+CH117 - AS-i SLAVE FAULT",
                SwpDiff.SimulateRecalculatedFinal("CHUTE =S1+CH117 - AS-i SLAVE FAULT", "",
                    " =S1+CH116-K1",
                    formulaReferencesDesc: false, formulaAppendPending: false, rebuildPending: false,
                    "Chute", "=S1+CH117"));
        }

        [Fact]
        public void A_pending_rebuild_cuts_the_template_tail_off_first()
        {
            // The real IL03_BELT row 778: B6 = "=S1+IN22", F679 = "=S1+IN22.05", and the template
            // "INDUCTION %1 - BELT %2 - SPARE ALARM.63" rendered as the cached text. The rebuild
            // keeps everything up to the block label and puts the description where the tail was.
            Assert.Equal(
                "INDUCTION =S1+IN22 - BELT =S1+IN22.05 - FIELD BUS FAILURE - DRIVE [n0115-in22-05-t1 - 192.168.3.115]",
                SwpDiff.SimulateRecalculatedFinal(
                    "INDUCTION =S1+IN22 - BELT =S1+IN22.05 - SPARE ALARM.63", "",
                    "FIELD BUS FAILURE - DRIVE [n0115-in22-05-t1 - 192.168.3.115]",
                    formulaReferencesDesc: false, formulaAppendPending: false, rebuildPending: true,
                    "=S1+IN22", "=S1+IN22.05"));

            // A tail containing " - " itself is exactly what a rendered-text cut would break on.
            Assert.Equal("INDUCTION =S1+IN22 - BELT =S1+IN22.05 - NEW TEXT",
                SwpDiff.SimulateRecalculatedFinal(
                    "INDUCTION =S1+IN22 - BELT =S1+IN22.05 - FIELD BUS FAILURE - DRIVE", "OLD",
                    "NEW TEXT", false, false, rebuildPending: true, "=S1+IN22", "=S1+IN22.05"));

            // A row already carrying the &" "&T extension rebuilds from the same cached text:
            // the anchors sit in the head, so the old appended description goes with the tail.
            Assert.Equal("CABINET =S1+MC1.CC1 - NEW TEXT",
                SwpDiff.SimulateRecalculatedFinal(
                    "CABINET =S1+MC1.CC1 - AS-i SLAVE FAULT OLD DESC", "OLD DESC", "NEW TEXT",
                    true, false, rebuildPending: true, "Cabinet", "=S1+MC1.CC1"));
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
            // The default replaces the template's trailing text, so a fresh row's formula is
            // part of what it proposes — the description reaches SCADA only through that shape.
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(EditFor(row, SwpField.FormulaQ).Proposed));
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

            // KeepConfigText: this is about the IO-side fields, and the other mode would queue a
            // formula rebuild on every row regardless of them.
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.KeepConfigText });

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
            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.KeepConfigText });
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

            // The standard formulas already reference T216/U216, so the label addition flows
            // through the description cell by itself: no APPEND may be proposed (appending would
            // print the description twice; seen in the field on +EMERGENCY rows). What the
            // default does propose is the rebuild, which references T216 exactly once.
            Assert.DoesNotContain("&\" \"&T216", EditFor(row, SwpField.FormulaQ).Proposed);
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(EditFor(row, SwpField.FormulaQ).Proposed));
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(EditFor(row, SwpField.FormulaR).Proposed));
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

            // The FLD reach is what matters: a template-only formula must end up displaying its
            // description cell, so the label appended into it reaches SCADA. This row's T/U takes
            // a real description, so it is the rebuilt shape that carries it.
            SwpSpineRow row = model.Rows[0];
            Assert.EndsWith("&\" - \"&T216)", EditFor(row, SwpField.FormulaQ).Proposed);
            Assert.EndsWith("&\" - \"&U216)", EditFor(row, SwpField.FormulaR).Proposed);
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(EditFor(row, SwpField.FormulaQ).Proposed));
        }

        [Fact]
        public void Overwrite_mode_puts_the_io_identity_in_the_rows_own_cells()
        {
            // The FLD rule: brutal overwrite means the IO identity lands in H/I/J — the
            // descriptions stay bare, the device is written, the differing cell gets the value.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockFUnit = "=S9";
            snapshot.HasCellH = true;
            snapshot.CachedH = "=S9";
            snapshot.FormulaH = "H181";       // the real rows: a plain per-cell formula, writable
            snapshot.FormulaHShared = false;
            snapshot.HasCellI = true;
            snapshot.CachedI = "+MC1.CC1";
            snapshot.FormulaI = "I181";
            snapshot.FormulaIShared = false;
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
            Assert.Equal("=S1", EditFor(row, SwpField.FUnitH).Proposed);      // the FU lands in H
            Assert.Null(EditFor(row, SwpField.LocationI));                   // the location agreed
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc1T).Proposed);
            Assert.Equal("-F22001", EditFor(row, SwpField.DeviceJ).Proposed); // …and the device in J
            // Where the FLD lands and what happens to CONFIG's text are separate questions: the
            // row still takes the workspace's CONFIG_PLANT mode, here the rebuild.
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(EditFor(row, SwpField.FormulaQ).Proposed));
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
        public void Keep_config_text_writes_only_the_fld_addition()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER";

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.KeepConfigText });

            SwpSpineRow row = model.Rows[0];
            Assert.Equal(" =S1+MC1.CC1-F22001", EditFor(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void Keep_config_text_with_no_mismatch_writes_nothing_at_all()
        {
            // No mismatch → the addition is empty → an existing template-fed cell is never blanked.
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.KeepConfigText });

            Assert.Null(EditFor(model.Rows[0], SwpField.Desc1T));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
        }

        [Fact]
        public void The_destination_sheet_rule_excepts_the_sheets_it_names()
        {
            // The CONFIG_PLANT rule has the FLD rule's shape: one regex on the DESTINATION sheet
            // naming the sheets that take the other mode, reported so its reach is never silent.
            var log = new List<SwpLogEntry>();
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    ConfigMode = SwpConfigMode.ReplaceTemplateTail,
                    ConfigSheetRule = SwpConfigSheetRule.Create("^CABINET", SwpConfigMode.KeepConfigText),
                }, log);

            SwpSpineRow row = model.Rows[0];
            Assert.True(row.ConfigModeFromSheetRule);
            Assert.True(row.ConfigFormulaKept);
            Assert.Null(EditFor(row, SwpField.FormulaQ));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info &&
                l.Message.Contains("1 row(s) follow the destination-sheet CONFIG_PLANT rule /^CABINET/"));
        }

        [Fact]
        public void A_destination_sheet_rule_that_does_not_parse_is_inert_with_a_warning()
        {
            var log = new List<SwpLogEntry>();
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    ConfigSheetRule = SwpConfigSheetRule.Create("([unclosed", SwpConfigMode.KeepConfigText),
                }, log);

            // Inert, never fatal: every sheet follows the default and the log says so.
            Assert.False(model.Rows[0].ConfigModeFromSheetRule);
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(EditFor(model.Rows[0], SwpField.FormulaQ).Proposed));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("CONFIG_PLANT rule") && l.Message.Contains("does not parse"));
        }

        // ---- the Q/R formula rebuild --------------------------------------------------------

        [Fact]
        public void Rebuild_mode_keeps_each_rows_own_template_reference()
        {
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.ReplaceTemplateTail });

            SwpSpineRow row = model.Rows[0];
            // The fixture's formula concatenates the template with T216 inside the SUBSTITUTE;
            // the rebuild expresses that suffix itself, so the template survives exactly once.
            Assert.Equal(
                "_xlfn.LET(_xlpm.t,[1]DB_CABINETS_DIAG!$K$57,_xlpm.p,FIND(CHAR(1),SUBSTITUTE(_xlpm.t,\"%\",CHAR(1)," +
                "LEN(_xlpm.t)-LEN(SUBSTITUTE(_xlpm.t,\"%\",\"\")))),SUBSTITUTE(LEFT(_xlpm.t,IFERROR(_xlpm.p,LEN(_xlpm.t))+1),\"%1\",F181)" +
                "&\" - \"&T216)",
                EditFor(row, SwpField.FormulaQ).Proposed);
            Assert.Equal(
                "_xlfn.LET(_xlpm.t,[1]DB_CABINETS_DIAG!$L$57,_xlpm.p,FIND(CHAR(1),SUBSTITUTE(_xlpm.t,\"%\",CHAR(1)," +
                "LEN(_xlpm.t)-LEN(SUBSTITUTE(_xlpm.t,\"%\",\"\")))),SUBSTITUTE(LEFT(_xlpm.t,IFERROR(_xlpm.p,LEN(_xlpm.t))+1),\"%1\",F181)" +
                "&\" - \"&U216)",
                EditFor(row, SwpField.FormulaR).Proposed);
        }

        [Fact]
        public void Rebuild_mode_leaves_a_row_whose_template_reference_is_gone()
        {
            // What the old wholesale CONCAT left behind: no template to truncate. The row is
            // named in the log rather than guessed at — the reference-SWP audit restores it.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.FormulaQ = "_xlfn.CONCAT(UPPER($B$6),\" \",$F181,\" - \",T216)";
            snapshot.FormulaR = "_xlfn.CONCAT(UPPER($B$6),\" \",$F181,\" - \",U216)";
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.ReplaceTemplateTail }, log);

            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaR));
            Assert.Equal(2, log.Count(l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("no CONFIG_PLANT template to rebuild from")));
        }

        [Fact]
        public void A_formula_excel_already_removed_is_named_not_silently_skipped()
        {
            // What the unloadable first build left once Excel repaired the file: the cell keeps its
            // (correct) cached text as a plain value and no longer shows T/U. Nothing can rebuild it
            // from itself — but a silent skip would leave it frozen forever.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.FormulaQ = "";
            snapshot.CachedQ = "CABINET =S1+MC1.CC1 - CIRCUIT BREAKER TRIPPED";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.ReplaceTemplateTail }, log).Rows[0];

            Assert.Null(EditFor(row, SwpField.FormulaQ));
            Assert.Single(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("final-text cell at") && l.Message.Contains("holds a plain value") &&
                l.Message.Contains("T216 never reaches the displayed text"));
            // R still carries its template, so it is rebuilt — and says nothing.
            Assert.True(SwpDescriptionFormula.IsCurrentForm(EditFor(row, SwpField.FormulaR).Proposed));
        }

        [Fact]
        public void A_cell_still_holding_the_unloadable_first_build_gets_the_loadable_one()
        {
            // The first build wrote LET with bare parameter names, and Excel deletes that formula
            // on open. A file patched but not yet opened still holds it; treating it as "already
            // rebuilt" would leave the loss armed, so the next build re-proposes the file spelling.
            const string unloadable =
                "_xlfn.LET(t,[1]DB_CABINETS_DIAG!$K$57,p,FIND(CHAR(1),SUBSTITUTE(t,\"%\",CHAR(1)," +
                "LEN(t)-LEN(SUBSTITUTE(t,\"%\",\"\")))),SUBSTITUTE(LEFT(t,IFERROR(p,LEN(t))+1),\"%1\",F181)" +
                "&\" - \"&T216)";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.FormulaQ = unloadable;

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.ReplaceTemplateTail }).Rows[0];

            SwpProposedEdit edit = EditFor(row, SwpField.FormulaQ);
            Assert.Equal(unloadable, edit.Current);
            Assert.Equal(
                "_xlfn.LET(_xlpm.t,[1]DB_CABINETS_DIAG!$K$57,_xlpm.p,FIND(CHAR(1),SUBSTITUTE(_xlpm.t,\"%\",CHAR(1)," +
                "LEN(_xlpm.t)-LEN(SUBSTITUTE(_xlpm.t,\"%\",\"\")))),SUBSTITUTE(LEFT(_xlpm.t,IFERROR(_xlpm.p,LEN(_xlpm.t))+1),\"%1\",F181)" +
                "&\" - \"&T216)",
                edit.Proposed);
        }

        [Fact]
        public void An_already_rebuilt_row_is_left_alone_and_says_nothing()
        {
            // Idempotence: the second build of the same workspace must queue nothing AND stay
            // silent — a rebuilt formula has no template reference of its own, but that is the
            // shape working correctly, not the damage the warning is about.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.FormulaQ = SwpDescriptionFormula.Build(
                Parse("SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",F181)"), "T216");
            snapshot.FormulaR = SwpDescriptionFormula.Build(
                Parse("SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57,\"%1\",F181)"), "U216");
            snapshot.BlockLocation = "+OTHER.PLACE";   // a mismatch too: neither path may fire
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.ReplaceTemplateTail }, log);

            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaR));
            Assert.DoesNotContain(log, l => l.Message.Contains("no CONFIG_PLANT template"));
        }

        private static SwpDescriptionFormula.Parsed Parse(string formula)
        {
            Assert.True(SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed parsed));
            return parsed;
        }

        [Fact]
        public void Rebuild_mode_leaves_a_row_whose_suffix_cell_would_hold_the_label_alone()
        {
            // Rule: the truncation only happens where T/U names the bit. This entry carries no
            // description at all, so T would take the FLD label — cutting the template's tail
            // off would leave the alarm nameless.
            DiagnosticEntry entry = CabinetEntry();
            entry.Desc1A = "";
            entry.Desc1B = "";
            entry.Desc2A = "";
            entry.Desc2B = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",F181)";
            snapshot.FormulaR = "SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57,\"%1\",F181)";

            SwpWorkspaceModel model = Build(entry, Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.ReplaceTemplateTail });

            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaR));
        }

        [Fact]
        public void A_mismatch_row_rebuilds_instead_of_appending_when_its_suffix_describes()
        {
            // The BELT case: the append extension is what makes the FLD label reach SCADA, but a
            // template-only formula plus a real description belongs in the rebuilt shape — the
            // description replaces the template's generic tail rather than stacking after it.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.FormulaQ = "SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%1\",$B$6),\"%2\",F181)";
            snapshot.FormulaR = "SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$L$122,\"%1\",$B$6),\"%2\",F181)";

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot))).Rows[0];

            Assert.Equal(
                "_xlfn.LET(_xlpm.t,[1]DB_IL_BELT!$K$122,_xlpm.p,FIND(CHAR(1),SUBSTITUTE(_xlpm.t,\"%\",CHAR(1)," +
                "LEN(_xlpm.t)-LEN(SUBSTITUTE(_xlpm.t,\"%\",\"\")))),SUBSTITUTE(SUBSTITUTE(LEFT(_xlpm.t,IFERROR(_xlpm.p,LEN(_xlpm.t))+1)," +
                "\"%1\",$B$6),\"%2\",F181)&\" - \"&T216)",
                EditFor(row, SwpField.FormulaQ).Proposed);
        }

        [Fact]
        public void A_trusted_mismatch_row_keeps_the_template_tail_and_only_appends_the_label()
        {
            // T holds the FLD label alone here, so the tail — the alarm's own name — must stay.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.FormulaQ = "SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%1\",$B$6),\"%2\",F181)";

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    ConfigMode = SwpConfigMode.KeepConfigText,
                }).Rows[0];

            Assert.Equal(
                "SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%1\",$B$6),\"%2\",F181)&\" \"&T216",
                EditFor(row, SwpField.FormulaQ).Proposed);
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

            // KeepConfigText: the other mode supersedes the append with the rebuilt shape, which
            // is a different question. Here the point is that appending TWICE never happens.
            SwpWorkspaceModel model = Build(CabinetEntry(), Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.KeepConfigText });

            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaR));
        }

        // ---- a kept formula that shows its description ---------------------------------------

        /// <summary>The FVZ CABINET shape: CONFIG_PLANT's template ends at "…%2 - " (or, with a tail,
        /// names the alarm itself) and the formula concatenates the row's own T/U after it — which
        /// still hold the English an older run left there.</summary>
        private static SwpWorkbookIndex CabinetKeptIndex(string templateTail = "")
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.FormulaQ = "SUBSTITUTE(SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",$B$6),\"%2\",F181)&\" \"&T216";
            snapshot.FormulaR = "SUBSTITUTE(SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57,\"%1\",$B$6),\"%2\",F181)&\" \"&U216";
            snapshot.Desc1T = "CIRCUIT BREAKER OPEN";
            snapshot.Desc2U = "CIRCUIT BREAKER OPEN";
            snapshot.CachedQ = "INTRODUCTION GRP01 - CABINET =S1+MC1.CC1 - " + templateTail + " CIRCUIT BREAKER OPEN";
            snapshot.CachedR = "INTRODUCTION GRP01 - ARMOIRE DE GROUPE =S1+MC1.CC1 - " + templateTail + " CIRCUIT BREAKER OPEN";
            SwpWorkbookIndex index = Index(("CABINET", snapshot));
            index.Sheets["CABINET"].Description = "INTRODUCTION GRP01";
            return index;
        }

        [Fact]
        public void A_kept_formula_showing_its_description_after_a_bare_template_still_receives_the_description()
        {
            // FVZ 2026-09-10: the CONFIG_PLANT rule kept every CABINET formula — right — and with it
            // withheld the descriptions, so a French-only SCADA went on reading the English an older
            // run had left in U. Keeping CONFIG's text keeps its WORDS; this template has none to keep.
            var log = new List<SwpLogEntry>();
            SwpSpineRow row = Build(CabinetEntry(), CabinetKeptIndex(),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    ConfigMode = SwpConfigMode.ReplaceTemplateTail,
                    ConfigSheetRule = SwpConfigSheetRule.Create("CABINET", SwpConfigMode.KeepConfigText),
                }, log).Rows[0];

            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc1T).Proposed);
            Assert.Equal("DISJONCTEUR OUVERT", EditFor(row, SwpField.Desc2U).Proposed);
            Assert.Null(EditFor(row, SwpField.FormulaQ));   // the formula itself is kept
            Assert.Null(EditFor(row, SwpField.FormulaR));
            Assert.True(row.ConfigFormulaKept);
            Assert.True(row.KeptFormulaShowsDescription);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info &&
                l.Message.StartsWith("1 row(s) keep CONFIG_PLANT's formula and still receive their description"));
        }

        [Fact]
        public void A_kept_formula_whose_template_names_the_alarm_keeps_the_label_only_treatment()
        {
            SwpSpineRow row = Build(CabinetEntry(), CabinetKeptIndex(templateTail: "SPARE ALARM.63"),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, ConfigMode = SwpConfigMode.KeepConfigText }).Rows[0];

            // The template's own sentence stays whole and nothing stacks onto it.
            Assert.Null(EditFor(row, SwpField.Desc1T));
            Assert.Null(EditFor(row, SwpField.Desc2U));
            Assert.False(row.KeptFormulaShowsDescription);
        }

        // ---- the second language ------------------------------------------------------------

        [Fact]
        public void Language_one_text_in_description_two_is_reported_once_a_second_language_is_configured()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.Desc2A = "";
            entry.Desc2B = "";

            var withLanguage = new List<SwpLogEntry>();
            SwpSpineRow row = Build(entry, Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, SecondLanguage = "FR" }, withLanguage).Rows[0];
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc2U).Proposed);
            Assert.Contains(withLanguage, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message == "no FR text for Description 2 — neither the IO list's language 2 nor the glossary " +
                             "provides one, so it repeats language 1: \"CIRCUIT BREAKER TRIPPED\"");

            // No second language configured, or English — language 1 itself: the fallback is the
            // design there, and stays quiet.
            foreach (string quiet in new[] { "", "EN", "en" })
            {
                var withoutLanguage = new List<SwpLogEntry>();
                Build(entry, Index(("CABINET", CabinetSnapshot())),
                    new SwpDiffOptions { PlantType = SwpPlantType.Msc, SecondLanguage = quiet }, withoutLanguage);
                Assert.DoesNotContain(withoutLanguage, l => l.Message.Contains("for Description 2"));
            }
        }

        [Fact]
        public void A_language_one_variable_part_beside_a_glossary_translation_is_reported()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.Desc1A = "CIRCUIT BREAKER TRIPPED";
            entry.Desc1B = "COOLING FAN";
            entry.Desc2A = "";
            entry.Desc2B = "";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(entry, Index(("CABINET", CabinetSnapshot())),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    SecondLanguage = "FR",
                    GlossaryPairs = new List<SwpGlossaryPair>
                    {
                        new SwpGlossaryPair { En = "CIRCUIT BREAKER TRIPPED", Translation = "DISJONCTEUR DÉCLENCHÉ" },
                    },
                }, log).Rows[0];

            Assert.Equal("DISJONCTEUR DÉCLENCHÉ COOLING FAN", EditFor(row, SwpField.Desc2U).Proposed);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("no FR variable part — \"COOLING FAN\" stays in language 1"));
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
            // A chute takes its SWP entry from Destination_Parameters, which this build never
            // loaded — so the mnemonic cannot be placed at all and is Invalid. It used to read
            // SheetMissing because the legacy mapper quietly fell back to the raw index, which
            // pointed at whatever row that number happened to hit.
            Assert.Equal(SwpRowVerdict.Invalid, model.Rows[2].Verdict);
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

        // ---- the empty-value guards: an empty IO-side value never erases an SWP one ----------

        private static SwpRowSnapshot OverwritableSnapshot()
        {
            // The real rows' writable shape: plain per-cell H/I formulas (not shared).
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.HasCellH = true;
            snapshot.CachedH = "=S1";
            snapshot.FormulaH = "H181";
            snapshot.FormulaHShared = false;
            snapshot.HasCellI = true;
            snapshot.CachedI = "+MC1.CC1";
            snapshot.FormulaI = "I181";
            snapshot.FormulaIShared = false;
            return snapshot;
        }

        [Fact]
        public void An_empty_io_device_never_blanks_the_swp_device()
        {
            // The review's C1: with no mismatch and a blank IO-list Device, J "-F22001" was
            // erased by an empty inline string — no log line, the device carried nowhere.
            DiagnosticEntry entry = CabinetEntry();
            entry.Device = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.DeviceJ = "-F22001";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(entry, Index(("CABINET", snapshot)), log: log).Rows[0];

            Assert.Null(EditFor(row, SwpField.DeviceJ));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("no device") && l.Message.Contains("-F22001") && l.Message.Contains("kept"));
        }

        [Fact]
        public void In_append_mode_an_empty_io_device_keeps_j_and_the_label_carries_the_place()
        {
            // The blank is the device MOVING into the label; with no device to move, J stays.
            DiagnosticEntry entry = CabinetEntry();
            entry.Device = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.DeviceJ = "-F22001";

            SwpSpineRow row = Build(entry, Index(("CABINET", snapshot))).Rows[0];

            Assert.True(row.LocationMismatch);
            Assert.Null(EditFor(row, SwpField.DeviceJ));
            Assert.Equal("CIRCUIT BREAKER TRIPPED =S1+MC1.CC1", EditFor(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void In_overwrite_mode_an_empty_io_device_keeps_j_while_the_place_is_written()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.Device = "";
            SwpRowSnapshot snapshot = OverwritableSnapshot();
            snapshot.CachedI = "+OTHER.PLACE";
            snapshot.DeviceJ = "-F22001";

            SwpSpineRow row = Build(entry, Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, MismatchMode = SwpMismatchMode.Overwrite }).Rows[0];

            Assert.Equal("+MC1.CC1", EditFor(row, SwpField.LocationI).Proposed);
            Assert.Null(EditFor(row, SwpField.DeviceJ));
        }

        [Fact]
        public void An_identity_less_generated_row_never_touches_the_swp_identity()
        {
            // The review's C2: strobe entries (and a photocell whose PHT row is missing) carry
            // empty FU / Location / Device, and GeneratedBy marks the mismatch deliberate — the
            // empty strings were treated as authoritative: H/I blanked in overwrite mode, J
            // blanked and a bare " " appended in append mode. Nothing may move.
            DiagnosticEntry entry = CabinetEntry();
            entry.GeneratedBy = "ASD";
            entry.FUnit = "";
            entry.Location = "";
            entry.Device = "";
            SwpRowSnapshot snapshot = OverwritableSnapshot();
            snapshot.DeviceJ = "-F22001";
            var log = new List<SwpLogEntry>();

            SwpSpineRow overwrite = Build(entry, Index(("CABINET", snapshot)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, MismatchMode = SwpMismatchMode.Overwrite }, log).Rows[0];
            Assert.False(overwrite.FUnitMismatch);
            Assert.False(overwrite.LocationMismatch);
            Assert.Null(EditFor(overwrite, SwpField.FUnitH));
            Assert.Null(EditFor(overwrite, SwpField.LocationI));
            Assert.Null(EditFor(overwrite, SwpField.DeviceJ));

            SwpSpineRow append = Build(entry, Index(("CABINET", snapshot)), log: log).Rows[0];
            Assert.Null(EditFor(append, SwpField.DeviceJ));
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(append, SwpField.Desc1T).Proposed.Trim());
            // No identity means no FLD append; the formula still takes the CONFIG_PLANT mode.
            Assert.DoesNotContain("&\" \"&T", EditFor(append, SwpField.FormulaQ).Proposed);
            Assert.DoesNotContain(log, l => l.Message.Contains("no functional unit"));   // generated: by design, silent
        }

        [Fact]
        public void An_io_list_row_without_any_identity_is_reported_and_leaves_the_row_alone()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.FUnit = "";
            entry.Location = "";
            entry.Device = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+ELSEWHERE";
            snapshot.DeviceJ = "-F22001";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(entry, Index(("CABINET", snapshot)), log: log).Rows[0];

            Assert.False(row.FUnitMismatch);
            Assert.False(row.LocationMismatch);
            Assert.NotEqual(SwpRowVerdict.LabelMismatch, row.Verdict);
            Assert.Null(EditFor(row, SwpField.DeviceJ));
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc1T).Proposed.Trim());
            Assert.Single(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("no functional unit, location or device"));
            Assert.DoesNotContain(log, l => l.Message.Contains("no device for this row"));   // one line, not two
        }

        [Fact]
        public void A_partial_identity_compares_only_its_present_members()
        {
            // Location missing from the IO list: "" against "+MC1.CC1" is missing data, not a
            // configuration difference — no mismatch, no label, the device still written.
            DiagnosticEntry entry = CabinetEntry();
            entry.Location = "";

            SwpSpineRow row = Build(entry, Index(("CABINET", CabinetSnapshot()))).Rows[0];

            Assert.False(row.LocationMismatch);
            Assert.Equal("-F22001", EditFor(row, SwpField.DeviceJ).Proposed);
            Assert.Equal("CIRCUIT BREAKER TRIPPED", EditFor(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void An_empty_composed_description_never_blanks_an_existing_one()
        {
            // The review's H1 — the T.S. Ref rule, applied to T/U.
            DiagnosticEntry entry = CabinetEntry();
            entry.Desc1A = "";
            entry.Desc1B = "";
            entry.Desc2A = "";
            entry.Desc2B = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.Desc1T = "OLD TEXT";
            snapshot.Desc2U = "ANCIEN TEXTE";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(entry, Index(("CABINET", snapshot)), log: log).Rows[0];

            Assert.Null(EditFor(row, SwpField.Desc1T));
            Assert.Null(EditFor(row, SwpField.Desc2U));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("composed description is empty"));
        }

        [Fact]
        public void An_empty_composed_description_still_carries_the_fld_label()
        {
            DiagnosticEntry entry = CabinetEntry();
            entry.Desc1A = "";
            entry.Desc1B = "";
            entry.Desc2A = "";
            entry.Desc2B = "";
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.Desc1T = "OLD TEXT";

            SwpSpineRow row = Build(entry, Index(("CABINET", snapshot))).Rows[0];

            Assert.Equal("=S1+MC1.CC1-F22001", EditFor(row, SwpField.Desc1T).Proposed.Trim());
            Assert.Equal("=S1+MC1.CC1-F22001", EditFor(row, SwpField.Desc2U).Proposed.Trim());
        }

        // ---- a kept row beside a rebuilt one --------------------------------------------------

        [Fact]
        public void A_row_ticked_to_keep_config_text_is_not_rebuilt_beside_one_that_is()
        {
            // The review's H2 in its current shape: a row excepted from the workspace's rebuild
            // must keep its CONFIG-linked formula, or it ends up with T/U holding only the label
            // and a formula whose template text was cut off — no text at all.
            DiagnosticEntry trusted = CabinetEntry();
            DiagnosticEntry plain = CabinetEntry("CABINETAL00101");
            SwpRowSnapshot other = CabinetSnapshot();
            other.Bit = 1;
            other.RowNumber = 217;
            var log = new List<SwpLogEntry>();

            SwpWorkspaceModel model = SwpDiff.Build(new[] { trusted, plain },
                Index(("CABINET", CabinetSnapshot()), ("CABINET", other)),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    ConfigMode = SwpConfigMode.ReplaceTemplateTail,
                    Overrides = new Dictionary<string, SwpRowOverride>
                    {
                        ["CABINETAL00100"] = new SwpRowOverride
                        {
                            Mnemonic = "CABINETAL00100",
                            TrustConfigDescriptions = true,
                        },
                    },
                }, log);

            Assert.True(model.Rows[0].ConfigFormulaKept);
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaQ));
            Assert.Null(EditFor(model.Rows[0], SwpField.FormulaR));
            Assert.False(model.Rows[1].ConfigFormulaKept);
            Assert.StartsWith("_xlfn.LET(", EditFor(model.Rows[1], SwpField.FormulaQ).Proposed);
            Assert.Single(log, l => l.Severity == SwpLogSeverity.Info &&
                l.Message.Contains("1 row(s) keep their CONFIG_PLANT description formulas"));
        }

        // ---- the append extension's reach -----------------------------------------------------

        [Fact]
        public void A_value_only_description_cell_becomes_a_formula_that_shows_the_label()
        {
            // The review's M2: a pasted-over literal in Q got no extension and no log line — J
            // blanked, the label written into T, and never displayed anywhere.
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.FormulaQ = "";
            snapshot.CachedQ = "CABINET =S1+MC1.CC1 - CIRCUIT \"BREAKER\"";
            snapshot.FormulaR = "";
            snapshot.CachedR = "";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot)), log: log).Rows[0];

            Assert.Equal("\"CABINET =S1+MC1.CC1 - CIRCUIT \"\"BREAKER\"\"\"&\" \"&T216",
                EditFor(row, SwpField.FormulaQ).Proposed);
            Assert.Equal("U216", EditFor(row, SwpField.FormulaR).Proposed);   // an empty cell: the description alone
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning && l.Message.Contains("plain value"));
        }

        [Fact]
        public void A_shared_formula_child_is_a_shared_formula_not_a_pasted_value()
        {
            // A shared child carries no formula text: it must get the shared-formula warning,
            // never the literal-to-formula rewrite (which would break the shared range).
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.FormulaQ = "";
            snapshot.FormulaQShared = true;
            snapshot.CachedQ = "CABINET =S1+MC1.CC1 - CIRCUIT BREAKER";
            var log = new List<SwpLogEntry>();

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot)), log: log).Rows[0];

            Assert.Null(EditFor(row, SwpField.FormulaQ));
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("Q final-text formula") && l.Message.Contains("shared formula"));
            Assert.DoesNotContain(log, l => l.Message.Contains("plain value"));
        }

        [Fact]
        public void A_formula_of_any_shape_gets_the_extension()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.FormulaQ = "$B$6&\" \"&$F$181";   // not ending in ")" — used to be skipped silently

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot))).Rows[0];

            Assert.Equal("$B$6&\" \"&$F$181&\" \"&T216", EditFor(row, SwpField.FormulaQ).Proposed);
        }

        // ---- the FLD group ---------------------------------------------------------------------

        [Fact]
        public void In_append_mode_the_device_blank_and_its_label_writes_form_one_group()
        {
            SwpRowSnapshot snapshot = CabinetSnapshot();
            snapshot.BlockLocation = "+OTHER.PLACE";
            snapshot.DeviceJ = "-F22001";
            snapshot.FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",F181)";   // template-only → extension queued

            SwpSpineRow row = Build(CabinetEntry(), Index(("CABINET", snapshot))).Rows[0];

            string group = EditFor(row, SwpField.DeviceJ).FldGroup;
            Assert.Equal("CABINET|DB_CABINETS_DIAG.Alarm1>1>0", group);
            Assert.Equal(group, EditFor(row, SwpField.Desc1T).FldGroup);
            Assert.Equal(group, EditFor(row, SwpField.Desc2U).FldGroup);
            Assert.Equal(group, EditFor(row, SwpField.FormulaQ).FldGroup);
            Assert.Equal("", EditFor(row, SwpField.ExportB).FldGroup);    // the row's other writes stay independent
            Assert.Equal("", EditFor(row, SwpField.AddressG).FldGroup);
        }

        [Fact]
        public void Without_a_device_blank_there_is_no_group()
        {
            // Overwrite mode: each cell carries its own member — independent writes.
            SwpRowSnapshot overwritable = OverwritableSnapshot();
            overwritable.CachedI = "+OTHER.PLACE";
            SwpSpineRow overwrite = Build(CabinetEntry(), Index(("CABINET", overwritable)),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, MismatchMode = SwpMismatchMode.Overwrite }).Rows[0];
            Assert.NotNull(EditFor(overwrite, SwpField.LocationI));
            Assert.All(overwrite.Edits, e => Assert.Equal("", e.FldGroup));

            // Append mode with J already blank: the label writes are harmless on their own.
            SwpRowSnapshot blankJ = CabinetSnapshot();
            blankJ.BlockLocation = "+OTHER.PLACE";
            SwpSpineRow append = Build(CabinetEntry(), Index(("CABINET", blankJ))).Rows[0];
            Assert.Null(EditFor(append, SwpField.DeviceJ));
            Assert.All(append.Edits, e => Assert.Equal("", e.FldGroup));
        }
    }
}
