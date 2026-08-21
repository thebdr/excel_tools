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
        internal const int ColVerdict = 8, ColCsv = 9, ColBusfield = 10, ColAlarms = 11;
        internal const int ColExportNow = 12, ColExportNew = 13;
        internal const int ColAddrNow = 14, ColAddrNew = 15;
        internal const int ColDeviceNow = 16, ColDeviceNew = 17;
        internal const int ColTsNow = 18, ColTsNew = 19;
        internal const int ColDesc1Now = 20, ColDesc1New = 21;
        internal const int ColDesc2Now = 22, ColDesc2New = 23;
        internal const int ColFormulas = 24;
        internal const int ColSkipLabel = 25, ColDbIndex = 26, ColForceIds = 27;
        internal const int ColDesc1Override = 28, ColDesc2Override = 29;
        internal const int ColFuSwp = 30, ColFuIo = 31, ColLocSwp = 32, ColLocIo = 33;
        internal const int ColumnCount = 33;

        internal const string CheckedGlyph = "☑";

        private static readonly string[] Headers =
        {
            "Card", "Row", "Mnemonic", "Address", "SWP sheet", "Coordinate", "SWP row",
            "Verdict", "Plant_Inputs", "Busfield", "Alarm strings",
            "Export", "Export ▸", "I/O addr", "I/O addr ▸", "Device", "Device ▸",
            "T.S. ref", "T.S. ref ▸", "Description 1", "Description 1 ▸",
            "Description 2", "Description 2 ▸", "Formulas",
            "[Mismatch OK]", "[DB index]", "[Force IDs]", "[Desc 1 override]", "[Desc 2 override]",
            "FU (SWP)", "FU (IO)", "Loc (SWP)", "Loc (IO)",
        };

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        private static readonly int DoubleClickBlue = Rgb(0, 112, 192);
        internal static readonly int SwpHeaderOrange = Rgb(250, 200, 140);   // the SWP side stands apart
        private static readonly int OverrideHeaderTint = Rgb(215, 228, 188);
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
            dynamic headerRange = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, ColumnCount));
            headerRange.NumberFormat = "@";
            WriteBlock(sheet, HeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;
            // The SWP-related block wears light orange for a quick spot (the IO side stays blue).
            sheet.Range(sheet.Cells(HeaderRow, ColSheet), sheet.Cells(HeaderRow, ColFormulas))
                .Interior.Color = SwpHeaderOrange;
            sheet.Cells(HeaderRow, ColFuSwp).Interior.Color = SwpHeaderOrange;
            sheet.Cells(HeaderRow, ColLocSwp).Interior.Color = SwpHeaderOrange;
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
                var data = new object[rows, ColumnCount];
                for (int r = 0; r < rows; r++)
                {
                    FillRow(data, r, model.Rows[r], overrides);
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
                foreach (int column in new[] { ColRow, ColCoordinate, ColSwpRow, ColVerdict, ColTsNow, ColTsNew })
                    MarkDoubleClickable(sheet.Range(
                        sheet.Cells(FirstDataRow, column), sheet.Cells(last, column)));
            }
            else
            {
                session.SpineData = null;
            }

            try
            {
                int lastRow = FirstDataRow + Math.Max(0, rows - 1);
                sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(lastRow, ColumnCount)).Columns.AutoFit();
                for (int c = 1; c <= ColumnCount; c++)
                {
                    dynamic column = sheet.Columns(c);
                    double width = (double)column.ColumnWidth;
                    if (width > 46) column.ColumnWidth = 46;
                    else if (width < 6) column.ColumnWidth = 6;
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

        private static void FillRow(object[,] data, int r, SwpSpineRow row,
            Dictionary<string, SwpRowOverride> overrides)
        {
            DiagnosticEntry e = row.Entry;
            data[r, ColCard - 1] = e.CardSheet;
            data[r, ColRow - 1] = e.SourceRow > 0 ? e.SourceRow.ToString() : "";
            data[r, ColMnemonic - 1] = e.Mnemonic;
            data[r, ColAddress - 1] = e.IsBusModule ? "#" + e.Node : e.SwpAddress();
            data[r, ColVerdict - 1] = VerdictLabel(row.Verdict);
            data[r, ColCsv - 1] = CsvLabel(row.CsvVerdict);
            data[r, ColBusfield - 1] = BusfieldLabel(row.BusfieldVerdict);
            data[r, ColAlarms - 1] = AlarmLabel(row.AlarmVerdict);

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
                data[r, ColFuSwp - 1] = s.BlockFUnit;
                data[r, ColLocSwp - 1] = s.BlockLocation;
            }
            data[r, ColFuIo - 1] = e.FUnit;
            data[r, ColLocIo - 1] = e.Location;

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
            data[r, ColFormulas - 1] = formulas;

            if (overrides.TryGetValue(e.Mnemonic, out SwpRowOverride o))
            {
                data[r, ColSkipLabel - 1] = o.SkipLabelCheck ? CheckedGlyph : "";
                data[r, ColDbIndex - 1] = o.SwpDbIndex.HasValue ? o.SwpDbIndex.Value.ToString() : "";
                data[r, ColForceIds - 1] = o.ForceElectricalIds ? CheckedGlyph : "";
                data[r, ColDesc1Override - 1] = o.Desc1Override;
                data[r, ColDesc2Override - 1] = o.Desc2Override;
            }
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

            // Mandatory data: an IO-list alarm left without ANY T.S. Ref (IO list, glossary
            // card and SWP all empty) wears red on its T.S. columns.
            foreach (var pair in session.WorksheetRowByModelRow)
            {
                SwpSpineRow row = session.Model.Rows[pair.Key];
                if (!row.TsRefMissing) continue;
                foreach (int column in new[] { ColTsNow, ColTsNew })
                {
                    if (!paints.TryGetValue(column, out SortedList<int, int> byRow))
                        paints[column] = byRow = new SortedList<int, int>();
                    byRow[pair.Value] = ConflictRed;
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
