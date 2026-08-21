using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// Renders the flat spine: one row per diagnostic mnemonic in IO-list order — IO side,
    /// computed coordinate, current SWP values, proposed values (blank when nothing changes),
    /// verdicts, and the five editable override columns. All writes are whole-block; pending
    /// highlights are painted directly (run-grouped), row states are conditional-format font
    /// colors keyed on the verdict column.
    /// </summary>
    internal static class SwpBridgeSpineSheet
    {
        internal const int MetaRow = 1;
        internal const int HeaderRow = 2;
        internal const int FirstDataRow = 3;

        // 1-based worksheet columns.
        internal const int ColCard = 1, ColRow = 2, ColMnemonic = 3, ColAddress = 4;
        internal const int ColSheet = 5, ColCoordinate = 6, ColSwpRow = 7;
        internal const int ColVerdict = 8, ColCsv = 9, ColBusfield = 10, ColAlarms = 11, ColBusExt = 12;
        internal const int ColExportNow = 13, ColExportNew = 14;
        internal const int ColAddrNow = 15, ColAddrNew = 16;
        internal const int ColFu = 17, ColLoc = 18;            // IOL / SWP identity, one column each
        internal const int ColDeviceNow = 19, ColDeviceNew = 20;
        internal const int ColTsNow = 21, ColTsNew = 22;
        internal const int ColDesc1Now = 23, ColDesc1New = 24;
        internal const int ColDesc2Now = 25, ColDesc2New = 26;
        internal const int ColFormulas = 27;
        internal const int ColSkipLabel = 28, ColDbIndex = 29, ColForceIds = 30, ColFldMode = 31;
        internal const int ColTrustConfig = 32;
        internal const int ColDesc1Override = 33, ColDesc2Override = 34;
        internal const int ColumnCount = 34;

        internal const string CheckedGlyph = "☑";

        private static readonly string[] Headers =
        {
            "Card", "Row", "Mnemonic", "Address", "SWP sheet", "Coordinate", "SWP row",
            "Verdict", "Plant_Inputs", "Busfield", "Alarm strings", "BUS ext. diag",
            "Export", "Export ▸", "I/O addr", "I/O addr ▸", "FU (IOL / SWP)", "Loc (IOL / SWP)",
            "Device", "Device ▸", "T.S. ref", "T.S. ref ▸", "Description 1", "Description 1 ▸",
            "Description 2", "Description 2 ▸", "Formulas",
            // ColFldMode's caption is computed at render time: it always names the OPPOSITE of the
            // workspace default, because ticking it is what the cell does.
            "[Mismatch OK]", "[DB index]", "[Force IDs]", "[Overwrite FLD]", "[Keep CONFIG text]",
            "[Desc 1 override]", "[Desc 2 override]",
        };

        /// <summary>
        /// The FLD column's caption: ticking a row flips it AWAY from the workspace default, so the
        /// header names the action the tick performs — [Overwrite FLD] under an append default,
        /// [Append FLD] under an overwrite one.
        /// </summary>
        internal static string FldModeHeader(SwpMismatchMode workspaceDefault) =>
            workspaceDefault == SwpMismatchMode.Overwrite ? "[Append FLD]" : "[Overwrite FLD]";

        /// <summary>
        /// The CONFIG-text column's caption, same opposite-of-the-default rule: under the normal
        /// setting the bridge writes its composed text, so ticking KEEPS CONFIG's; under "never
        /// overwrite CONFIG_PLANT descriptions" it is the IO text that a tick restores.
        /// </summary>
        internal static string TrustConfigHeader(bool workspaceDefault) =>
            workspaceDefault ? "[Write IO text]" : "[Keep CONFIG text]";

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);

        // Header colors say where a column's data comes from: the IO list (blue), one of the
        // checked production files (lavender), the SWP (orange), the bridge's own overrides (olive).
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        internal static readonly int CheckHeaderLavender = Rgb(221, 211, 245);
        internal static readonly int SwpHeaderOrange = Rgb(250, 200, 140);
        private static readonly int OverrideHeaderTint = Rgb(215, 228, 188);
        private static readonly int DoubleClickBlue = Rgb(0, 112, 192);
        private static readonly int MismatchTint = Rgb(253, 228, 200);    // FU/Loc cell whose two sides differ
        private static readonly int PendingYellow = Rgb(255, 235, 156);
        private static readonly int ConflictRed = Rgb(255, 199, 206);
        private static readonly int ExternalBlue = Rgb(221, 235, 247);
        private static readonly int MatchingGreen = Rgb(0x00, 0xCC, 0x5C);
        private static readonly int ErrorRed = Rgb(200, 30, 30);
        private static readonly int MismatchOrange = Rgb(200, 120, 0);

        /// <summary>The blue border every double-clickable cell wears, across all bridge sheets.</summary>
        internal static void MarkDoubleClickable(dynamic range)
        {
            try
            {
                dynamic borders = range.Borders;
                borders.LineStyle = 1;   // xlContinuous
                borders.Weight = 2;      // xlThin
                borders.Color = DoubleClickBlue;
            }
            catch { }
        }

        /// <summary>Creates or refills the spine sheet. Call inside an ExcelStateGuard.</summary>
        public static void Render(dynamic sheet, SwpBridgeSession session)
        {
            SwpWorkspaceModel model = session.Model;
            sheet.Cells.Clear();
            try { sheet.Cells.FormatConditions.Delete(); } catch { }
            try { sheet.Name = SwpBridgeWorkspace.SpineSheetName; } catch { }

            sheet.Cells(MetaRow, 1).Value =
                "SWP Bridge · " + session.Settings.SwpFileName + " · " + session.Settings.PlantType +
                " · blue-bordered cells react to a double-click: Verdict/Coordinate preview the " +
                "SWP block, Row opens the IO list at that row, SWP row opens the SWP there · " +
                "[bracketed] columns are editable and persist";
            sheet.Cells(MetaRow, 1).Font.Bold = true;

            var headers = new object[1, ColumnCount];
            for (int c = 0; c < ColumnCount; c++) headers[0, c] = Headers[c];
            headers[0, ColFldMode - 1] = FldModeHeader(session.Settings.MismatchMode);
            headers[0, ColTrustConfig - 1] = TrustConfigHeader(session.Settings.TrustConfigPlant);
            dynamic headerRange = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, ColumnCount));
            headerRange.NumberFormat = "@";
            WriteBlock(sheet, HeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;   // Card … Address: the IO list
            headerRange.Font.Bold = true;
            // Every other header wears the color of its source: the SWP (sheet, coordinate,
            // row, verdict, the current/proposed cells, formulas), the checked production files
            // (the four coherence verdicts), and the bridge's own override columns.
            sheet.Range(sheet.Cells(HeaderRow, ColSheet), sheet.Cells(HeaderRow, ColVerdict))
                .Interior.Color = SwpHeaderOrange;
            sheet.Range(sheet.Cells(HeaderRow, ColCsv), sheet.Cells(HeaderRow, ColBusExt))
                .Interior.Color = CheckHeaderLavender;
            sheet.Range(sheet.Cells(HeaderRow, ColExportNow), sheet.Cells(HeaderRow, ColFormulas))
                .Interior.Color = SwpHeaderOrange;
            sheet.Range(sheet.Cells(HeaderRow, ColSkipLabel), sheet.Cells(HeaderRow, ColDesc2Override))
                .Interior.Color = OverrideHeaderTint;

            session.ModelRowByWorksheetRow.Clear();
            session.WorksheetRowByModelRow.Clear();

            // The CF anchor below needs a Select; remember where the user was so an in-handler
            // re-render (override/dictionary edits) does not steal the selection.
            dynamic activeCell = null;
            try { activeCell = ExcelPowerTools.Interop.ExcelApp.App.ActiveCell; } catch { }

            int rows = model.Rows.Count;
            if (rows > 0)
            {
                var overrides = session.OverridesByMnemonic();
                var repairs = RepairsByTargetRow(session);
                var data = new object[rows, ColumnCount];
                for (int r = 0; r < rows; r++)
                {
                    FillRow(data, r, model.Rows[r], overrides, repairs,
                        session.Settings.MismatchMode, session.Settings.TrustConfigPlant);
                    session.ModelRowByWorksheetRow[FirstDataRow + r] = r;
                    session.WorksheetRowByModelRow[r] = FirstDataRow + r;
                }

                dynamic block = sheet.Range(
                    sheet.Cells(FirstDataRow, 1), sheet.Cells(FirstDataRow + rows - 1, ColumnCount));
                block.NumberFormat = "@";
                WriteBlock(sheet, FirstDataRow, 1, data);
                session.SpineData = data;

                ApplyVerdictFormats(sheet, rows);
                PaintActionCells(sheet, session);

                int last = FirstDataRow + rows - 1;
                foreach (int column in new[] { ColRow, ColCoordinate, ColSwpRow, ColVerdict,
                    ColCsv, ColBusfield, ColAlarms, ColBusExt, ColTsNow, ColTsNew })
                    MarkDoubleClickable(sheet.Range(
                        sheet.Cells(FirstDataRow, column), sheet.Cells(last, column)));
            }
            else
            {
                session.SpineData = null;
            }

            try
            {
                // Largest cell wins — no cap: the user kept re-autofitting the capped columns.
                int lastRow = FirstDataRow + Math.Max(0, rows - 1);
                sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(lastRow, ColumnCount)).Columns.AutoFit();
                for (int c = 1; c <= ColumnCount; c++)
                {
                    dynamic column = sheet.Columns(c);
                    double width = (double)column.ColumnWidth;
                    // AutoFit measures the caption but not the AutoFilter drop-down that sits over
                    // its right end — a column whose header is its widest content lost the last
                    // glyph ("T.S. ref ▸" showed without the ▸). Bold caption ≈ 1.15 units per
                    // character, the button ≈ 3.
                    double headerNeed = Headers[c - 1].Length * 1.15 + 3.0;
                    double wanted = Math.Max(6.0, Math.Max(width, headerNeed));
                    if (Math.Abs(wanted - width) > 0.1) column.ColumnWidth = wanted;
                }
                sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(lastRow, ColumnCount)).AutoFilter();
            }
            catch { }

            try
            {
                dynamic win = ExcelPowerTools.Interop.ExcelApp.App.ActiveWindow;
                if (!(bool)win.FreezePanes)
                {
                    win.SplitRow = HeaderRow;
                    win.SplitColumn = ColMnemonic;
                    win.FreezePanes = true;
                }
            }
            catch { }

            try { activeCell?.Select(); } catch { }
        }

        /// <summary>
        /// The formula REPAIRS (restored H/I references) are workbook-level edits, not spine-row
        /// edits — indexed here by the SWP row they target so the Formulas column can show them
        /// on the mnemonic that lives there. Header-level repairs have no spine row and stay on
        /// the queue only.
        /// </summary>
        private static Dictionary<(string Sheet, int Row), List<SwpProposedEdit>> RepairsByTargetRow(SwpBridgeSession session)
        {
            var repairs = new Dictionary<(string Sheet, int Row), List<SwpProposedEdit>>();
            if (session.Model?.Edits == null) return repairs;
            foreach (SwpProposedEdit edit in session.Model.Edits)
            {
                if (edit.Field != SwpField.FormulaH && edit.Field != SwpField.FormulaI) continue;
                var key = (edit.Sheet ?? "", edit.RowNumber);
                if (!repairs.TryGetValue(key, out List<SwpProposedEdit> list))
                    repairs[key] = list = new List<SwpProposedEdit>();
                list.Add(edit);
            }
            return repairs;
        }

        private static (string Sheet, int Row) TargetOf(SwpSpineRow row) =>
            row.Snapshot != null && row.Mapped != null && row.Mapped.Ok
                ? (row.Mapped.Coordinate.SheetName ?? "", row.Snapshot.RowNumber)
                : ("", -1);

        private static void FillRow(object[,] data, int r, SwpSpineRow row,
            Dictionary<string, SwpRowOverride> overrides,
            Dictionary<(string Sheet, int Row), List<SwpProposedEdit>> repairs,
            SwpMismatchMode workspaceDefault, bool trustConfigDefault)
        {
            DiagnosticEntry e = row.Entry;
            data[r, ColCard - 1] = e.CardSheet;
            data[r, ColRow - 1] = e.SourceRow > 0 ? e.SourceRow.ToString() : "";
            data[r, ColMnemonic - 1] = e.Mnemonic;
            data[r, ColAddress - 1] = e.IsBusModule ? "#" + (e.AsiNode.Length > 0 ? e.AsiNode : e.Node) : e.SwpAddress();
            data[r, ColVerdict - 1] = VerdictLabel(row.Verdict);
            data[r, ColCsv - 1] = CsvLabel(row.CsvVerdict);
            data[r, ColBusfield - 1] = BusfieldLabel(row.BusfieldVerdict);
            data[r, ColAlarms - 1] = AlarmLabel(row.AlarmVerdict);
            data[r, ColBusExt - 1] = BusExtLabel(row.BusExtVerdict);

            if (row.Mapped != null && row.Mapped.Ok)
            {
                data[r, ColSheet - 1] = row.Mapped.Coordinate.SheetName;
                data[r, ColCoordinate - 1] = row.Mapped.Coordinate.RowKey;
            }
            else
            {
                data[r, ColSheet - 1] = "";
                data[r, ColCoordinate - 1] = row.Mapped?.Error ?? "";
            }

            SwpRowSnapshot s = row.Snapshot;
            if (s != null)
            {
                data[r, ColSwpRow - 1] = s.RowNumber.ToString();
                data[r, ColExportNow - 1] = s.ExportB;
                data[r, ColAddrNow - 1] = s.AddressG;
                data[r, ColDeviceNow - 1] = s.DeviceJ;
                data[r, ColTsNow - 1] = s.TsRefL;
                data[r, ColDesc1Now - 1] = s.Desc1T;
                data[r, ColDesc2Now - 1] = s.Desc2U;
            }
            data[r, ColFu - 1] = Identity(e.FUnit, s?.BlockFUnit, row.FUnitMismatch);
            data[r, ColLoc - 1] = Identity(e.Location, s?.BlockLocation, row.LocationMismatch);

            // The Formulas column is the row's formula status: the queued Q/R description-formula
            // change (append / rewrite) and the queued repairs of its H/I formulas.
            string formulas = "";
            foreach (SwpProposedEdit edit in row.Edits)
            {
                switch (edit.Field)
                {
                    case SwpField.ExportB: data[r, ColExportNew - 1] = edit.Proposed; break;
                    case SwpField.AddressG: data[r, ColAddrNew - 1] = edit.Proposed; break;
                    case SwpField.DeviceJ: data[r, ColDeviceNew - 1] = edit.Proposed; break;
                    case SwpField.TsRefL: data[r, ColTsNew - 1] = edit.Proposed; break;
                    case SwpField.Desc1T: data[r, ColDesc1New - 1] = edit.Proposed; break;
                    case SwpField.Desc2U: data[r, ColDesc2New - 1] = edit.Proposed; break;
                    case SwpField.FormulaQ:
                    case SwpField.FormulaR:
                        formulas = edit.Proposed.StartsWith("CONCAT", StringComparison.Ordinal)
                            ? "rewrite" : "append";
                        break;
                }
            }
            if (repairs.TryGetValue(TargetOf(row), out List<SwpProposedEdit> rowRepairs))
            {
                bool h = false, i = false;
                foreach (SwpProposedEdit repair in rowRepairs)
                {
                    if (repair.Field == SwpField.FormulaH) h = true;
                    if (repair.Field == SwpField.FormulaI) i = true;
                }
                string repairText = h && i ? "repair H+I" : h ? "repair H" : "repair I";
                formulas = formulas.Length > 0 ? formulas + " · " + repairText : repairText;
            }
            data[r, ColFormulas - 1] = formulas;

            if (overrides.TryGetValue(e.Mnemonic, out SwpRowOverride o))
            {
                data[r, ColSkipLabel - 1] = o.SkipLabelCheck ? CheckedGlyph : "";
                data[r, ColDbIndex - 1] = o.SwpDbIndex.HasValue ? o.SwpDbIndex.Value.ToString() : "";
                data[r, ColForceIds - 1] = o.ForceElectricalIds ? CheckedGlyph : "";
                // Ticked only while the row genuinely departs from the default; an override that
                // says what the default already says is redundant, not a flip.
                data[r, ColFldMode - 1] =
                    o.FldMode.HasValue && o.FldMode.Value != workspaceDefault ? CheckedGlyph : "";
                data[r, ColTrustConfig - 1] =
                    o.TrustConfigDescriptions.HasValue && o.TrustConfigDescriptions.Value != trustConfigDefault
                        ? CheckedGlyph : "";
                data[r, ColDesc1Override - 1] = o.Desc1Override;
                data[r, ColDesc2Override - 1] = o.Desc2Override;
            }
        }

        /// <summary>
        /// One identity column per field: the value alone when the IO list and the SWP block
        /// agree (or only one side has it), "io / swp" when they differ — the header reads
        /// "FU (IOL / SWP)", so the order is the header's.
        /// </summary>
        private static string Identity(string ioValue, string swpValue, bool mismatch)
        {
            string io = (ioValue ?? "").Trim();
            string swp = (swpValue ?? "").Trim();
            if (!mismatch) return io.Length > 0 ? io : swp;
            return (io.Length > 0 ? io : "—") + " / " + (swp.Length > 0 ? swp : "—");
        }

        private static void ApplyVerdictFormats(dynamic sheet, int rowCount)
        {
            const int XlExpression = 2;
            dynamic block = sheet.Range(
                sheet.Cells(FirstDataRow, 1), sheet.Cells(FirstDataRow + rowCount - 1, ColumnCount));
            try { block.FormatConditions.Delete(); } catch { }

            // Select the block's first cell first so the RELATIVE row in the formulas anchors there.
            try { sheet.Cells(FirstDataRow, 1).Select(); } catch { }

            string verdictRef = "$H" + FirstDataRow;
            AddFontRule(block, XlExpression, "=" + verdictRef + "=\"up to date\"", MatchingGreen);
            AddFontRule(block, XlExpression,
                "=OR(" + verdictRef + "=\"row missing\"," + verdictRef + "=\"sheet missing\"," +
                verdictRef + "=\"invalid\"," + verdictRef + "=\"duplicate\")", ErrorRed);
            AddFontRule(block, XlExpression, "=" + verdictRef + "=\"label mismatch\"", MismatchOrange);
        }

        private static void AddFontRule(dynamic range, int type, string formula, int fontColor)
        {
            dynamic rule = range.FormatConditions.Add(type, Type.Missing, formula);
            rule.Font.Color = fontColor;
            rule.StopIfTrue = false;
        }

        /// <summary>
        /// Direct interior paint on the "▸" cells that carry an open action, grouped into
        /// contiguous runs per column so a thousand highlights cost a handful of COM calls.
        /// </summary>
        internal static void PaintActionCells(dynamic sheet, SwpBridgeSession session)
        {
            var actionByKey = new Dictionary<string, SwpAction>();
            foreach (SwpAction action in session.Queue.Actions)
                if (!actionByKey.ContainsKey(action.Key) || action.IsOpen)
                    actionByKey[action.Key] = action;

            // column → (row → color)
            var paints = new Dictionary<int, SortedList<int, int>>();
            foreach (var pair in session.WorksheetRowByModelRow)
            {
                SwpSpineRow row = session.Model.Rows[pair.Key];
                foreach (SwpProposedEdit edit in row.Edits)
                {
                    if (!actionByKey.TryGetValue(edit.Key, out SwpAction action)) continue;
                    int color;
                    switch (action.Status)
                    {
                        case SwpActionStatus.Pending: color = PendingYellow; break;
                        case SwpActionStatus.Conflict:
                        case SwpActionStatus.ConflictStale: color = ConflictRed; break;
                        case SwpActionStatus.AppliedExternally: color = ExternalBlue; break;
                        default: continue;
                    }
                    int column = NewColumnOf(edit.Field);
                    if (!paints.TryGetValue(column, out SortedList<int, int> byRow))
                        paints[column] = byRow = new SortedList<int, int>();
                    byRow[pair.Value] = color;
                }
            }

            // The H/I formula repairs are workbook-level edits: paint their status on the
            // Formulas cell of the spine row living on the repaired SWP row.
            var worksheetRowByTarget = new Dictionary<(string Sheet, int Row), int>();
            foreach (var pair in session.WorksheetRowByModelRow)
            {
                var target = TargetOf(session.Model.Rows[pair.Key]);
                if (target.Row > 0 && !worksheetRowByTarget.ContainsKey(target)) worksheetRowByTarget[target] = pair.Value;
            }
            foreach (SwpProposedEdit edit in session.Model.Edits ?? new List<SwpProposedEdit>())
            {
                if (edit.Field != SwpField.FormulaH && edit.Field != SwpField.FormulaI) continue;
                if (!worksheetRowByTarget.TryGetValue((edit.Sheet ?? "", edit.RowNumber), out int worksheetRow)) continue;
                if (!actionByKey.TryGetValue(edit.Key, out SwpAction action)) continue;
                int color;
                switch (action.Status)
                {
                    case SwpActionStatus.Pending: color = PendingYellow; break;
                    case SwpActionStatus.Conflict:
                    case SwpActionStatus.ConflictStale: color = ConflictRed; break;
                    case SwpActionStatus.AppliedExternally: color = ExternalBlue; break;
                    default: continue;
                }
                Paint(paints, ColFormulas, worksheetRow, color);
            }

            // Mandatory data: an IO-list alarm left without ANY T.S. Ref (IO list, glossary
            // card and SWP all empty) wears red on its T.S. columns. An identity column whose
            // two sides differ (and was not marked deliberate) wears the mismatch tint.
            foreach (var pair in session.WorksheetRowByModelRow)
            {
                SwpSpineRow row = session.Model.Rows[pair.Key];
                if (row.TsRefMissing)
                    foreach (int column in new[] { ColTsNow, ColTsNew })
                        Paint(paints, column, pair.Value, ConflictRed);
                if (!row.MismatchSuppressed)
                {
                    if (row.FUnitMismatch) Paint(paints, ColFu, pair.Value, MismatchTint);
                    if (row.LocationMismatch) Paint(paints, ColLoc, pair.Value, MismatchTint);
                }
            }

            foreach (var columnPair in paints)
            {
                int column = columnPair.Key;
                SortedList<int, int> byRow = columnPair.Value;
                int runStart = -1, runEnd = -1, runColor = 0;
                foreach (var rowPair in byRow)
                {
                    if (runStart >= 0 && rowPair.Key == runEnd + 1 && rowPair.Value == runColor)
                    {
                        runEnd = rowPair.Key;
                        continue;
                    }
                    FlushRun(sheet, column, runStart, runEnd, runColor);
                    runStart = runEnd = rowPair.Key;
                    runColor = rowPair.Value;
                }
                FlushRun(sheet, column, runStart, runEnd, runColor);
            }
        }

        private static void Paint(Dictionary<int, SortedList<int, int>> paints, int column, int row, int color)
        {
            if (!paints.TryGetValue(column, out SortedList<int, int> byRow))
                paints[column] = byRow = new SortedList<int, int>();
            byRow[row] = color;
        }

        private static void FlushRun(dynamic sheet, int column, int fromRow, int toRow, int color)
        {
            if (fromRow < 0) return;
            try
            {
                sheet.Range(sheet.Cells(fromRow, column), sheet.Cells(toRow, column)).Interior.Color = color;
            }
            catch { }
        }

        internal static int NewColumnOf(SwpField field)
        {
            switch (field)
            {
                case SwpField.ExportB: return ColExportNew;
                case SwpField.AddressG: return ColAddrNew;
                case SwpField.DeviceJ: return ColDeviceNew;
                case SwpField.TsRefL: return ColTsNew;
                case SwpField.Desc1T: return ColDesc1New;
                case SwpField.Desc2U: return ColDesc2New;
                case SwpField.FUnitH: return ColFu;
                case SwpField.LocationI: return ColLoc;
                default: return ColFormulas;
            }
        }

        internal static string VerdictLabel(SwpRowVerdict verdict)
        {
            switch (verdict)
            {
                case SwpRowVerdict.UpToDate: return "up to date";
                case SwpRowVerdict.Pending: return "pending";
                case SwpRowVerdict.LabelMismatch: return "label mismatch";
                case SwpRowVerdict.RowMissing: return "row missing";
                case SwpRowVerdict.SheetMissing: return "sheet missing";
                case SwpRowVerdict.DuplicateTarget: return "duplicate";
                default: return "invalid";
            }
        }

        private static string CsvLabel(SwpCsvVerdict verdict)
        {
            switch (verdict)
            {
                case SwpCsvVerdict.Ok: return "ok";
                case SwpCsvVerdict.NameMissing: return "name missing";
                case SwpCsvVerdict.AddressMismatch: return "address differs";
                default: return "";
            }
        }

        private static string BusfieldLabel(SwpBusfieldVerdict verdict)
        {
            switch (verdict)
            {
                case SwpBusfieldVerdict.Ok: return "ok";
                case SwpBusfieldVerdict.NodeMissing: return "node missing";
                case SwpBusfieldVerdict.DataMismatch: return "data differs";
                default: return "";
            }
        }

        private static string BusExtLabel(SwpBusExtVerdict verdict)
        {
            switch (verdict)
            {
                case SwpBusExtVerdict.Ok: return "ok";
                case SwpBusExtVerdict.NotFound: return "not bound";
                case SwpBusExtVerdict.NodeMismatch: return "node differs";
                case SwpBusExtVerdict.SubDeviceMismatch: return "sub-device differs";
                default: return "";
            }
        }

        private static string AlarmLabel(SwpAlarmVerdict verdict)
        {
            switch (verdict)
            {
                case SwpAlarmVerdict.Ok: return "ok";
                case SwpAlarmVerdict.NotFound: return "not exported";
                case SwpAlarmVerdict.DescriptionMismatch: return "text differs";
                default: return "";
            }
        }

        internal static void WriteBlock(dynamic sheet, int firstRow, int firstCol, object[,] data)
        {
            int rows = data.GetLength(0), cols = data.GetLength(1);
            if (rows == 0 || cols == 0) return;

            // COM consumes ONE leading apostrophe as the text-prefix marker even on array
            // writes; a value that genuinely STARTS with an apostrophe (legacy '-device residue
            // in SWP cells) must be doubled, or the sheet displays a lie — in the field, an
            // edit normalizing '-A1001 → -A1001 rendered as a baffling no-op. The copy leaves
            // the caller's array (the SpineData cache) untouched.
            object[,] display = null;
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    if (!(data[r, c] is string s) || s.Length == 0 || s[0] != '\'') continue;
                    if (display == null) display = (object[,])data.Clone();
                    display[r, c] = "'" + s;
                }

            dynamic range = sheet.Range(
                sheet.Cells(firstRow, firstCol), sheet.Cells(firstRow + rows - 1, firstCol + cols - 1));
            range.Value = display ?? data;
        }
    }
}
