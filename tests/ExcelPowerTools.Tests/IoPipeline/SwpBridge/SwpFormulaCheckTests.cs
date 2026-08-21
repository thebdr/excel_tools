using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpFormulaCheckTests
    {
        // ---- fixtures (model-level; the reader capture is tested separately) ----------------

        private static SwpRowSnapshot BitRow(int row, int headerRow = 15,
            string formulaH = "", bool sharedH = false, bool hasH = true, string cachedH = "",
            string formulaI = "", bool sharedI = false, bool hasI = true, string cachedI = "")
        {
            return new SwpRowSnapshot
            {
                RowNumber = row,
                TagType = "DB_TEST.Alarm1",
                TableEntry = 0,
                Bit = row - 17,
                BlockHeaderRow = headerRow,
                BlockFUnit = "=S1",
                BlockLocation = "+X1",
                BlockLabel = "=S1+X1",
                FormulaH = formulaH,
                FormulaHShared = sharedH,
                HasCellH = hasH,
                CachedH = cachedH,
                FormulaI = formulaI,
                FormulaIShared = sharedI,
                HasCellI = hasI,
                CachedI = cachedI,
            };
        }

        private static SwpSheetIndex Sheet(string name, bool diagnostic = true,
            params SwpRowSnapshot[] rows)
        {
            var sheet = new SwpSheetIndex { SheetName = name, IsDiagnosticSheet = diagnostic };
            foreach (SwpRowSnapshot row in rows)
            {
                sheet.BitRows.Add(row);
                sheet.RowsByKey[row.RowKey] = row;
            }
            var header = new SwpBlockHeader
            {
                RowNumber = 15,
                LabelF = "=S1+X1",
                FUnitH = "=S1",
                LocationI = "+X1",
            };
            sheet.BlockHeaders.Add(header);
            sheet.BlockHeaderByRow[15] = header;
            return sheet;
        }

        private static SwpWorkbookIndex Index(params SwpSheetIndex[] sheets)
        {
            var index = new SwpWorkbookIndex();
            foreach (SwpSheetIndex sheet in sheets) index.Sheets[sheet.SheetName] = sheet;
            return index;
        }

        // ---- self-contained repairs ---------------------------------------------------------

        [Fact]
        public void Shared_formulas_are_the_intact_template()
        {
            var index = Index(Sheet("CABINET", rows: BitRow(17, sharedH: true, sharedI: true)));
            Assert.Empty(SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>()));
        }

        [Fact]
        public void A_virtual_object_with_no_header_location_is_healthy_as_it_is()
        {
            // SORTERSTATE/TRAIN shape from the real FVZ file: header Location empty, rows with
            // an absent or blank I cell. A repair formula would only display a spurious 0.
            SwpRowSnapshot absent = BitRow(17, sharedH: true, formulaI: "", hasI: false);
            SwpRowSnapshot blank = BitRow(18, sharedH: true, formulaI: "", hasI: true, cachedI: "");
            absent.BlockLocation = "";
            blank.BlockLocation = "";
            SwpSheetIndex sheet = Sheet("SORTERSTATE", rows: new[] { absent, blank });
            sheet.BlockHeaders[0].LocationI = "";
            sheet.BlockHeaders[0].LabelF = "=S1";   // no '+' to rebuild from either

            var log = new List<SwpLogEntry>();
            Assert.Empty(SwpFormulaCheck.ProposeRepairs(Index(sheet), log));
            Assert.Empty(log);   // and no "lost its '+'" noise about the empty header
        }

        [Fact]
        public void An_empty_header_funit_also_silences_the_H_repair()
        {
            SwpRowSnapshot row = BitRow(17, formulaH: "", hasH: false, sharedI: true);
            row.BlockFUnit = "";
            SwpSheetIndex sheet = Sheet("VIRTUAL", rows: row);
            sheet.BlockHeaders[0].FUnitH = "";

            Assert.Empty(SwpFormulaCheck.ProposeRepairs(Index(sheet), new List<SwpLogEntry>()));
        }

        [Fact]
        public void A_pasted_value_over_H_queues_the_repair()
        {
            var log = new List<SwpLogEntry>();
            var index = Index(Sheet("CABINET",
                rows: BitRow(17, formulaH: "", hasH: true, cachedH: "=S9", sharedI: true)));

            List<SwpProposedEdit> repairs = SwpFormulaCheck.ProposeRepairs(index, log);

            SwpProposedEdit repair = Assert.Single(repairs);
            Assert.Equal(SwpField.FormulaH, repair.Field);
            Assert.Equal("H15", repair.Proposed);
            Assert.Equal("", repair.Current);
            Assert.Equal("DB_TEST.Alarm1>0>0", repair.CoordinateKey);
            Assert.Contains(log, l => l.Message.Contains("pasted value") && l.Message.Contains("=S9"));
        }

        [Fact]
        public void A_formula_pointing_elsewhere_queues_the_repair()
        {
            var index = Index(Sheet("CABINET",
                rows: BitRow(17, sharedH: true, formulaI: "I99")));
            SwpProposedEdit repair = Assert.Single(SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>()));
            Assert.Equal(SwpField.FormulaI, repair.Field);
            Assert.Equal("I15", repair.Proposed);
            Assert.Equal("I99", repair.Current);
        }

        [Fact]
        public void A_correct_unshared_formula_is_left_alone()
        {
            var index = Index(Sheet("CABINET",
                rows: BitRow(17, formulaH: "H15", formulaI: "$I$15")));
            Assert.Empty(SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>()));
        }

        [Fact]
        public void A_header_location_without_its_plus_is_rebuilt_from_the_label()
        {
            SwpSheetIndex sheet = Sheet("CABINET", rows: BitRow(17, sharedH: true, sharedI: true));
            sheet.BlockHeaders[0].LocationI = "X1 broken";

            List<SwpProposedEdit> repairs = SwpFormulaCheck.ProposeRepairs(Index(sheet), new List<SwpLogEntry>());

            SwpProposedEdit repair = Assert.Single(repairs);
            Assert.Equal(SwpField.HeaderLocationI, repair.Field);
            Assert.Equal("header>15", repair.CoordinateKey);
            Assert.Equal("X1 broken", repair.Current);
            Assert.Equal("+X1", repair.Proposed);
        }

        [Fact]
        public void A_label_without_a_plus_only_warns()
        {
            SwpSheetIndex sheet = Sheet("CABINET", rows: BitRow(17, sharedH: true, sharedI: true));
            sheet.BlockHeaders[0].LocationI = "broken";
            sheet.BlockHeaders[0].LabelF = "NO PLUS HERE";
            var log = new List<SwpLogEntry>();

            Assert.Empty(SwpFormulaCheck.ProposeRepairs(Index(sheet), log));
            Assert.Contains(log, l => l.Message.Contains("offers none"));
        }

        [Fact]
        public void A_strobe_sheet_first_header_wants_the_mid_formula()
        {
            SwpSheetIndex sheet = Sheet("UP01_STROBE", rows: BitRow(17, sharedH: true, sharedI: true));
            sheet.BlockHeaders[0].LocationIFormula = "";
            sheet.BlockHeaders[0].LocationI = "whatever";   // the '+' rule must NOT fire here

            List<SwpProposedEdit> repairs = SwpFormulaCheck.ProposeRepairs(Index(sheet), new List<SwpLogEntry>());

            SwpProposedEdit repair = Assert.Single(repairs);
            Assert.Equal(SwpField.HeaderFormulaI, repair.Field);
            Assert.Equal("MID(F15,4,12)", repair.Proposed);
        }

        [Fact]
        public void A_correct_strobe_formula_normalizes_and_passes()
        {
            SwpSheetIndex sheet = Sheet("UP01_STROBE", rows: BitRow(17, sharedH: true, sharedI: true));
            sheet.BlockHeaders[0].LocationIFormula = "mid($F$15, 4, 12)";
            Assert.Empty(SwpFormulaCheck.ProposeRepairs(Index(sheet), new List<SwpLogEntry>()));
        }

        [Fact]
        public void Non_diagnostic_sheets_are_skipped()
        {
            var index = Index(Sheet("Cover", diagnostic: false,
                rows: BitRow(17, formulaH: "", cachedH: "junk")));
            Assert.Empty(SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>()));
        }

        // ---- the base audit -----------------------------------------------------------------

        private static SwpRowSnapshot AuditRow(int row, uint formulaMask,
            (int Col, string Formula)[] externals = null)
        {
            var snapshot = BitRow(row, sharedH: true, sharedI: true);
            snapshot.FormulaMask = formulaMask;
            if (externals != null)
            {
                snapshot.ExternalFormulas = new Dictionary<int, string>();
                foreach ((int col, string formula) in externals)
                    snapshot.ExternalFormulas[col] = formula;
            }
            return snapshot;
        }

        [Fact]
        public void The_audit_reports_lost_formulas_by_block_offset()
        {
            // Base block: rows 15..180; its row 17 has formulas in Q (bit 16).
            SwpSheetIndex baseSheet = Sheet("UP01_BELT", rows: AuditRow(17, formulaMask: 1u << 16));
            // Target's row 183 folds to offset 2 → base row 17; it lost the Q formula.
            SwpSheetIndex target = Sheet("UP03_BELT", rows: AuditRow(183, formulaMask: 0));
            target.BlockHeaders[0].RowNumber = 181;
            target.BlockHeaderByRow.Clear();
            target.BlockHeaderByRow[181] = target.BlockHeaders[0];
            target.BitRows[0].RowNumber = 183;

            var log = new List<SwpLogEntry>();
            SwpFormulaCheck.AuditAgainstBase(Index(target), Index(baseSheet), log);

            Assert.Contains(log, l => l.Message.Contains("row 183") && l.Message.Contains("col Q") &&
                                      l.Message.Contains("formula was lost"));
        }

        [Fact]
        public void The_audit_flags_diverging_config_plant_references()
        {
            SwpSheetIndex baseSheet = Sheet("CABINET", rows: AuditRow(17, 1u << 16,
                new[] { (17, "SUBSTITUTE([BASE_CONFIG_PLANT.xlsx]DB_X!$K$57&T17,\"%1\",F15)") }));
            SwpSheetIndex target = Sheet("CABINET", rows: AuditRow(17, 1u << 16,
                new[] { (17, "SUBSTITUTE([CONFIG_PLANT.xlsx]DB_X!$K$99&T17,\"%1\",F15)") }));

            var log = new List<SwpLogEntry>();
            SwpFormulaCheck.AuditAgainstBase(Index(target), Index(baseSheet), log);
            Assert.Contains(log, l => l.Message.Contains("CONFIG_PLANT reference differs"));
        }

        [Fact]
        public void Equivalent_references_normalize_silently()
        {
            SwpSheetIndex baseSheet = Sheet("CABINET", rows: AuditRow(17, 1u << 16,
                new[] { (17, "SUBSTITUTE([BASE_CONFIG_PLANT.xlsx]DB_X!$K$57&T17,\"%1\",F15)") }));
            SwpSheetIndex target = Sheet("CABINET", rows: AuditRow(17, 1u << 16,
                new[] { (17, "SUBSTITUTE([CONFIG_PLANT.xlsx]DB_X!K57&T17,\"%1\",F15)") }));

            var log = new List<SwpLogEntry>();
            SwpFormulaCheck.AuditAgainstBase(Index(target), Index(baseSheet), log);
            Assert.DoesNotContain(log, l => l.Message.Contains("differs"));
        }

        [Fact]
        public void Sheets_map_by_suffix_when_no_exact_name_matches()
        {
            SwpSheetIndex baseSheet = Sheet("UP01_ZONE", rows: AuditRow(17, 1u << 16));
            SwpSheetIndex target = Sheet("UP07_ZONE", rows: AuditRow(17, 0));

            var log = new List<SwpLogEntry>();
            SwpFormulaCheck.AuditAgainstBase(Index(target), Index(baseSheet), log);
            Assert.Contains(log, l => l.Message.Contains("UP07_ZONE") && l.Message.Contains("lost"));
        }

        [Fact]
        public void An_unmapped_sheet_is_noted_not_failed()
        {
            SwpSheetIndex target = Sheet("SOMETHING", rows: AuditRow(17, 0));
            var log = new List<SwpLogEntry>();
            SwpFormulaCheck.AuditAgainstBase(Index(target), Index(), log);
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Info && l.Message.Contains("no reference sheet"));
        }
    }
}
