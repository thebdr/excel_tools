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
    /// <summary>
    /// The FLD rule: the IO list's Functional unit / Location / Device must ALWAYS be visible
    /// in the SWP — in the row's own cells (brutal overwrite) or appended to the descriptions.
    /// The field case: CABINETAL01914 "SAFETY DOOR OPEN", IO =S1 +SG14 -B1 against a block that
    /// says =S1 +SAFETYDOORS for sixteen doors; in overwrite mode the first implementation wrote
    /// the descriptions bare and never touched H/I, so "+SG14" existed nowhere.
    /// </summary>
    public class SwpFldRuleTests
    {
        // ---- fixtures -----------------------------------------------------------------------

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
            Desc2A = "LA PORTE DE SÉCURITÉ EST OUVERTE",
        };

        /// <summary>CABINETAL01914 → cabinet 19, bit 14 → DB_CABINETS_DIAG.Alarm1&gt;19&gt;14 (bits 0-31
        /// live in Alarm1), on a block "=S1 +SAFETYDOORS".</summary>
        private static SwpRowSnapshot DoorRow(bool sharedFormulas = false) => new SwpRowSnapshot
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
            FormulaH = sharedFormulas ? "" : "H2251",
            FormulaHShared = sharedFormulas,
            HasCellI = true,
            CachedI = "+SAFETYDOORS",
            FormulaI = sharedFormulas ? "" : "I2251",
            FormulaIShared = sharedFormulas,
            Desc1T = "AL14",
            Desc2U = "AL14",
            FormulaQ = "SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57&T2300,\"%1\",F2251)",
            FormulaR = "SUBSTITUTE([1]DB_CABINETS_DIAG!$L$57&U2300,\"%1\",F2251)",
        };

        private static SwpWorkbookIndex Index(SwpRowSnapshot row)
        {
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

        private static SwpSpineRow Build(SwpRowSnapshot row, SwpMismatchMode mode, List<SwpLogEntry> log = null) =>
            SwpDiff.Build(new[] { Door() }, Index(row),
                new SwpDiffOptions { PlantType = SwpPlantType.Msc, MismatchMode = mode },
                log ?? new List<SwpLogEntry>()).Rows[0];

        private static SwpProposedEdit Edit(SwpSpineRow row, SwpField field) =>
            row.Edits.FirstOrDefault(e => e.Field == field);

        /// <summary>The same build with a per-row [Overwrite FLD] / [Append FLD] flip in play.</summary>
        private static SwpSpineRow BuildWithFldOverride(SwpRowSnapshot row, SwpMismatchMode workspaceDefault,
            SwpMismatchMode? rowMode) =>
            SwpDiff.Build(new[] { Door() }, Index(row),
                new SwpDiffOptions
                {
                    PlantType = SwpPlantType.Msc,
                    MismatchMode = workspaceDefault,
                    Overrides = new Dictionary<string, SwpRowOverride>
                    {
                        ["CABINETAL01914"] = new SwpRowOverride
                        {
                            Mnemonic = "CABINETAL01914",
                            FldMode = rowMode,
                        },
                    },
                },
                new List<SwpLogEntry>()).Rows[0];

        // ---- the per-row FLD flip -----------------------------------------------------------

        [Fact]
        public void A_row_flipped_to_overwrite_writes_its_cells_under_an_append_default()
        {
            SwpSpineRow row = BuildWithFldOverride(DoorRow(),
                workspaceDefault: SwpMismatchMode.AppendToDescription,
                rowMode: SwpMismatchMode.Overwrite);

            Assert.True(row.FldOverwriteMode);
            Assert.Equal("+SG14", Edit(row, SwpField.LocationI).Proposed);          // the cell, not the text
            Assert.Equal("-B1", Edit(row, SwpField.DeviceJ).Proposed);
            Assert.Equal("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed.Trim());
        }

        [Fact]
        public void A_row_flipped_to_append_writes_the_descriptions_under_an_overwrite_default()
        {
            SwpSpineRow row = BuildWithFldOverride(DoorRow(),
                workspaceDefault: SwpMismatchMode.Overwrite,
                rowMode: SwpMismatchMode.AppendToDescription);

            Assert.False(row.FldOverwriteMode);
            Assert.Null(Edit(row, SwpField.LocationI));                              // H/I untouched
            Assert.Null(Edit(row, SwpField.FUnitH));
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);    // the FLD still lands
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc2U).Proposed);
        }

        [Fact]
        public void An_override_without_a_mode_follows_the_workspace_default()
        {
            SwpSpineRow appended = BuildWithFldOverride(DoorRow(),
                SwpMismatchMode.AppendToDescription, rowMode: null);
            SwpSpineRow overwritten = BuildWithFldOverride(DoorRow(),
                SwpMismatchMode.Overwrite, rowMode: null);

            Assert.False(appended.FldOverwriteMode);
            Assert.EndsWith(" =S1+SG14-B1", Edit(appended, SwpField.Desc1T).Proposed);
            Assert.True(overwritten.FldOverwriteMode);
            Assert.Equal("+SG14", Edit(overwritten, SwpField.LocationI).Proposed);
        }

        // ---- the per-row CONFIG-text answer --------------------------------------------------

        private static SwpSpineRow BuildWithTrust(bool workspaceDefault, bool? rowAnswer)
        {
            var options = new SwpDiffOptions
            {
                PlantType = SwpPlantType.Msc,
                MismatchMode = SwpMismatchMode.AppendToDescription,
                TrustConfigPlant = workspaceDefault,
            };
            if (rowAnswer.HasValue)
                options.Overrides = new Dictionary<string, SwpRowOverride>
                {
                    ["CABINETAL01914"] = new SwpRowOverride
                    {
                        Mnemonic = "CABINETAL01914",
                        TrustConfigDescriptions = rowAnswer,
                    },
                };
            return SwpDiff.Build(new[] { Door() }, Index(DoorRow()), options, new List<SwpLogEntry>()).Rows[0];
        }

        [Fact]
        public void Keep_config_text_on_an_ordinary_row_writes_the_fld_label_alone()
        {
            // The workspace default would compose the IO text here (the row is neither generated
            // nor "P*"-typed): an explicit tick must still take effect, or the gesture does nothing.
            SwpSpineRow row = BuildWithTrust(workspaceDefault: false, rowAnswer: true);

            Assert.Equal("=S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed.Trim());
            Assert.DoesNotContain("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void Write_io_text_on_a_row_beats_the_workspace_trust_setting()
        {
            SwpSpineRow row = BuildWithTrust(workspaceDefault: true, rowAnswer: false);

            Assert.StartsWith("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed);
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void Without_a_row_answer_the_default_keeps_its_automatic_scope()
        {
            // Default ON, but this row is neither generated nor "P*"-typed → it still composes.
            SwpSpineRow row = BuildWithTrust(workspaceDefault: true, rowAnswer: null);

            Assert.StartsWith("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed);
        }

        [Fact]
        public void The_fld_flip_survives_a_sidecar_round_trip_and_an_empty_one_is_purged()
        {
            var document = new SwpSidecarDocument();
            document.Overrides.Add(new SwpRowOverride
            {
                Mnemonic = "CABINETAL01914",
                FldMode = SwpMismatchMode.Overwrite,
                TrustConfigDescriptions = true,
            });
            document.Overrides.Add(new SwpRowOverride { Mnemonic = "OTHER" });   // nothing set

            Assert.True(SwpSidecarJson.TryParse(SwpSidecarJson.Serialize(document),
                out SwpSidecarDocument read, out string error), error);

            Assert.Equal(SwpMismatchMode.Overwrite, read.Overrides.Single(o => o.Mnemonic == "CABINETAL01914").FldMode);
            Assert.True(read.Overrides.Single(o => o.Mnemonic == "CABINETAL01914").TrustConfigDescriptions);
            // A mode-less override carries nothing and is dropped on save (IsEmpty).
            Assert.Null(read.Overrides.Single(o => o.Mnemonic == "OTHER").FldMode);
            Assert.True(read.Overrides.Single(o => o.Mnemonic == "OTHER").IsEmpty);
        }

        // ---- the diff -----------------------------------------------------------------------

        [Fact]
        public void Append_mode_carries_the_fld_in_the_descriptions()
        {
            SwpSpineRow row = Build(DoorRow(), SwpMismatchMode.AppendToDescription);

            Assert.True(row.LocationMismatch);
            // (the legacy joins the two description parts with a space verbatim — an empty
            // variable part leaves a double space before the label; the cell compare trims)
            Assert.StartsWith("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed);
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);
            Assert.StartsWith("LA PORTE DE SÉCURITÉ EST OUVERTE", Edit(row, SwpField.Desc2U).Proposed);
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc2U).Proposed);
            Assert.Null(Edit(row, SwpField.DeviceJ));   // J stays blank — the device rides in the label
            Assert.Null(Edit(row, SwpField.FUnitH));
            Assert.Null(Edit(row, SwpField.LocationI));
        }

        [Fact]
        public void Overwrite_mode_carries_the_fld_in_the_rows_own_cells()
        {
            SwpSpineRow row = Build(DoorRow(), SwpMismatchMode.Overwrite);

            Assert.True(row.LocationMismatch);
            Assert.False(row.FUnitMismatch);
            Assert.Equal("+SG14", Edit(row, SwpField.LocationI).Proposed);   // I gets the location
            Assert.Equal("+SAFETYDOORS", Edit(row, SwpField.LocationI).Current);
            Assert.Null(Edit(row, SwpField.FUnitH));                         // H already agrees
            Assert.Equal("-B1", Edit(row, SwpField.DeviceJ).Proposed);       // J gets the device
            Assert.Equal("SAFETY DOOR OPEN", Edit(row, SwpField.Desc1T).Proposed.Trim());   // descriptions bare
            Assert.Null(Edit(row, SwpField.FormulaQ));
            Assert.Equal(SwpRowVerdict.LabelMismatch, row.Verdict);
        }

        [Fact]
        public void Shared_formula_cells_cannot_be_overwritten_and_fall_back_to_the_append()
        {
            var log = new List<SwpLogEntry>();
            SwpSpineRow row = Build(DoorRow(sharedFormulas: true), SwpMismatchMode.Overwrite, log);

            Assert.Null(Edit(row, SwpField.LocationI));
            Assert.EndsWith(" =S1+SG14-B1", Edit(row, SwpField.Desc1T).Proposed);
            Assert.Null(Edit(row, SwpField.DeviceJ));   // blank stays blank — the label carries the device
            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Warning &&
                l.Message.Contains("shared formulas") && l.Message.Contains("appended"));
        }

        [Fact]
        public void An_overwritten_row_reads_up_to_date_on_the_next_build()
        {
            // After the apply, I holds the plain value "+SG14" and J "-B1": the comparison now
            // stands on the row's own cells — no mismatch, nothing proposed.
            SwpRowSnapshot row = DoorRow();
            row.CachedI = "+SG14";
            row.FormulaI = "";
            row.DeviceJ = "-B1";
            row.Desc1T = "SAFETY DOOR OPEN";
            row.Desc2U = "LA PORTE DE SÉCURITÉ EST OUVERTE";
            row.TsRefL = "TS_009";
            row.AddressG = "235.32.5";
            row.ExportB = "X";

            SwpSpineRow spine = Build(row, SwpMismatchMode.Overwrite);

            Assert.False(spine.LocationMismatch);
            Assert.Equal(SwpRowVerdict.UpToDate, spine.Verdict);
            Assert.Empty(spine.Edits);
        }

        [Fact]
        public void In_append_mode_the_same_row_still_compares_against_the_block()
        {
            SwpRowSnapshot row = DoorRow();
            row.CachedI = "+SG14";   // a pasted value — append mode trusts the block, the cleaner restores
            row.FormulaI = "";

            SwpSpineRow spine = Build(row, SwpMismatchMode.AppendToDescription);

            Assert.True(spine.LocationMismatch);
            Assert.EndsWith("=S1+SG14-B1", Edit(spine, SwpField.Desc1T).Proposed);
        }

        // ---- the formula cleaner ------------------------------------------------------------

        [Fact]
        public void A_plain_value_equal_to_the_io_identity_is_deliberate_not_a_repair()
        {
            SwpRowSnapshot row = DoorRow();
            row.CachedI = "+SG14";
            row.FormulaI = "";
            var index = Index(row);
            index.Sheets["CABINET"].IsDiagnosticSheet = true;

            var deliberate = new Dictionary<(string Sheet, int Row), (string FUnit, string Location)>
            {
                [("CABINET", 2300)] = ("=S1", "+SG14"),
            };
            List<SwpProposedEdit> repairs = SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>(), deliberate);
            Assert.Empty(repairs);

            // The same value with no deliberate identity — or another one — is the pasted accident it looks like.
            Assert.Single(SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>()));
            deliberate[("CABINET", 2300)] = ("=S1", "+SG99");
            Assert.Single(SwpFormulaCheck.ProposeRepairs(index, new List<SwpLogEntry>(), deliberate));
        }

        // ---- the patcher --------------------------------------------------------------------

        private static Cell Inline(string reference, string text) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new XlsxText(text)),
        };

        private static Cell Number(string reference, double value) => new Cell
        {
            CellReference = reference,
            CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture)),
        };

        private static Cell StrFormula(string reference, string formula, string cached) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellFormula = new CellFormula(formula),
            CellValue = new CellValue(cached),
        };

        private static Cell SharedStrFormula(string reference, string cached) => new Cell
        {
            CellReference = reference,
            DataType = CellValues.String,
            CellFormula = new CellFormula("") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U },
            CellValue = new CellValue(cached),
        };

        /// <summary>A block "=S1 +SAFETYDOORS" with row 17 on the real file's plain per-cell H/I
        /// formulas and row 18 on shared ones.</summary>
        private static byte[] Workbook()
        {
            using var stream = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbook = document.AddWorkbookPart();
                workbook.Workbook = new Workbook();
                var cabinet = workbook.AddNewPart<WorksheetPart>();
                cabinet.Worksheet = new Worksheet(new SheetData(
                    new Row(Inline("B4", "DB_TEST_DIAG")) { RowIndex = 4 },
                    new Row(
                        Inline("B15", "Table entry:"),
                        Number("C15", 0),
                        StrFormula("F15", "H15&I15", "=S1+SAFETYDOORS"),
                        Inline("H15", "=S1"),
                        Inline("I15", "+SAFETYDOORS")) { RowIndex = 15 },
                    new Row(
                        StrFormula("A17", "$B$4&\".Alarm1\"", "DB_TEST_DIAG.Alarm1"),
                        Inline("B17", "x"),
                        Number("C17", 0),
                        Number("D17", 0),
                        StrFormula("H17", "H15", "=S1"),
                        StrFormula("I17", "I15", "+SAFETYDOORS"),
                        Inline("T17", "AL14")) { RowIndex = 17 },
                    new Row(
                        StrFormula("A18", "$B$4&\".Alarm1\"", "DB_TEST_DIAG.Alarm1"),
                        Inline("B18", "x"),
                        Number("C18", 0),
                        Number("D18", 1),
                        SharedStrFormula("H18", "=S1"),
                        SharedStrFormula("I18", "+SAFETYDOORS")) { RowIndex = 18 }));
                workbook.Workbook.AppendChild(new Sheets(
                    new Sheet { Id = workbook.GetIdOfPart(cabinet), SheetId = 1, Name = "CABINET" }));
            }
            return stream.ToArray();
        }

        private static SwpAction Action(SwpField field, string oldValue, string newValue, string coordinate) => new SwpAction
        {
            Sheet = "CABINET",
            CoordinateKey = coordinate,
            Field = field,
            Mnemonic = "CABINETAL01914",
            OldValue = oldValue,
            NewValue = newValue,
        };

        [Fact]
        public void The_patcher_writes_the_location_value_over_the_rows_plain_formula()
        {
            byte[] bytes = Workbook();
            SwpRowSnapshot before = SwpSheetReader.Read(bytes, new[] { "CABINET" }, new List<SwpLogEntry>())
                .Sheets["CABINET"].RowsByKey["DB_TEST_DIAG.Alarm1>0>0"];
            Assert.Equal("I15", before.FormulaI);
            Assert.False(before.FormulaIShared);

            SwpXlsxPatcher.PatchResult result = SwpXlsxPatcher.Apply(bytes,
                new[]
                {
                    Action(SwpField.LocationI, "+SAFETYDOORS", "+SG14", "DB_TEST_DIAG.Alarm1>0>0"),
                    Action(SwpField.DeviceJ, "", "-B1", "DB_TEST_DIAG.Alarm1>0>0"),
                },
                "2026-08-21.20:00:00", new List<SwpLogEntry>());

            Assert.All(result.Outcomes, o => Assert.Equal(SwpActionStatus.Applied, o.Result));
            SwpRowSnapshot after = SwpSheetReader.Read(result.PatchedBytes, new[] { "CABINET" }, new List<SwpLogEntry>())
                .Sheets["CABINET"].RowsByKey["DB_TEST_DIAG.Alarm1>0>0"];
            Assert.Equal("+SG14", after.CachedI);    // the FLD is in the row's own cell…
            Assert.Equal("", after.FormulaI);        // …as a plain value, the block formula gone
            Assert.Equal("=S1", after.CachedH);      // the untouched side keeps its formula
            Assert.Equal("H15", after.FormulaH);
            Assert.Equal("-B1", after.DeviceJ);
        }

        [Fact]
        public void The_patcher_refuses_a_shared_formula_cell_so_the_diff_must_have_appended()
        {
            SwpXlsxPatcher.PatchResult result = SwpXlsxPatcher.Apply(Workbook(),
                new[] { Action(SwpField.LocationI, "+SAFETYDOORS", "+SG14", "DB_TEST_DIAG.Alarm1>0>1") },
                "2026-08-21.20:00:00", new List<SwpLogEntry>());

            SwpXlsxPatcher.Outcome outcome = Assert.Single(result.Outcomes);
            Assert.Equal(SwpActionStatus.Conflict, outcome.Result);
            Assert.Contains("shared formula", outcome.Note);
            Assert.Null(result.PatchedBytes);
        }
    }
}
