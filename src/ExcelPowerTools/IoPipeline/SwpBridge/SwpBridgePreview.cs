using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The SWP block preview: a temporary _Preview sheet showing the double-clicked row IN
    /// CONTEXT. Default view = its 32-bit section as the SWP holds it; the "show full block"
    /// toggle (double-click) widens to the device's WHOLE block — every section — with every
    /// pending update of that device shown, each labeled with the mnemonic that feeds it.
    /// Changing cells read "current → new" and wear green; the clicked row stays yellow.
    /// Read-only: any edit re-renders.
    /// </summary>
    internal static class SwpBridgePreview
    {
        internal sealed class Context
        {
            public int ModelRow;
            public bool ShowFullBlock;
        }

        internal const int ToggleRow = 2;
        internal const int HeaderRow = 3;
        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int TargetYellow = Rgb(255, 235, 156);
        private static readonly int NewValueGreen = Rgb(198, 239, 206);

        private static readonly string[] Headers =
        {
            // The written columns in the SWP's own order — Export (B), I/O addr (G), then the FLD
            // as neighbours: FUnit (H) · Location (I) · Device (J) — so the preview reads like the file.
            "SWP row", "Tag type (A)", "Entry (C)", "Bit (D)", "Mnemonic",
            "Export (B)", "I/O addr (G)", "FUnit (H)", "Location (I)", "Device (J)",
            "T.S. ref (L)", "Description 1 (T)", "Description 2 (U)", "Final text (Q)", "Final text (R)", "Stamp (AA)",
        };

        public static void Open(SwpBridgeSession session, int modelRow)
        {
            SwpSpineRow row = session.Model.Rows[modelRow];
            if (row.Snapshot == null || row.Mapped == null || !row.Mapped.Ok)
            {
                try
                {
                    ExcelApp.App.StatusBar = "SWP Bridge: " +
                        (row.Mapped != null && row.Mapped.Ok
                            ? "no SWP row to preview — the coordinate was not found"
                            : "nothing to preview — the mnemonic did not map");
                }
                catch { }
                return;
            }

            bool keepFullBlock = session.PreviewContext?.ShowFullBlock ?? false;
            session.PreviewContext = new Context { ModelRow = modelRow, ShowFullBlock = keepFullBlock };
            Rerender(session);

            dynamic workbook = SwpBridgeWorkspace.FindLiveWorkbook(ExcelApp.App);
            try { workbook?.Worksheets(SwpBridgeWorkspace.PreviewSheetName).Activate(); } catch { }
        }

        /// <summary>
        /// Creates/refreshes the _Preview sheet for a row WITHOUT stealing focus — used at build
        /// time so the sheet exists (and its tab is colored) the moment the workspace opens.
        /// </summary>
        public static void Prepare(SwpBridgeSession session, int modelRow)
        {
            if (session?.Model == null || modelRow < 0 || modelRow >= session.Model.Rows.Count) return;
            SwpSpineRow row = session.Model.Rows[modelRow];
            if (row.Snapshot == null || row.Mapped == null || !row.Mapped.Ok) return;

            session.PreviewContext = new Context
            {
                ModelRow = modelRow,
                ShowFullBlock = session.PreviewContext?.ShowFullBlock ?? false,
            };
            Rerender(session);   // renders into _Preview; unlike Open it never activates the sheet
        }

        public static void ToggleFullBlock(SwpBridgeSession session)
        {
            if (session.PreviewContext == null) return;
            session.PreviewContext.ShowFullBlock = !session.PreviewContext.ShowFullBlock;
            Rerender(session);
        }

        public static void Rerender(SwpBridgeSession session)
        {
            Context context = session.PreviewContext;
            if (context == null || session.Model == null ||
                context.ModelRow >= session.Model.Rows.Count) return;

            dynamic workbook = SwpBridgeWorkspace.FindLiveWorkbook(ExcelApp.App);
            if (workbook == null) return;

            SwpBridgeEditCapture.RunSuppressed(() =>
            {
                dynamic sheet = SwpBridgeAuxSheets.EnsureSheet(workbook, SwpBridgeWorkspace.PreviewSheetName);
                PlaceAfterWorkspace(workbook, sheet);
                RenderInto(sheet, session, session.Model.Rows[context.ModelRow], context.ShowFullBlock);
            });
        }

        /// <summary>
        /// _Preview belongs next to the sheet it unfolds: the tab sits right after Workspace, not
        /// at the far end of the strip where EnsureSheet appends it. Moved only when it is in the
        /// wrong place — a Move on every re-render would flicker the tab strip on each toggle.
        /// </summary>
        private static void PlaceAfterWorkspace(dynamic workbook, dynamic sheet)
        {
            try
            {
                dynamic spine = SwpBridgeQueueSheet.FindSheet(workbook, SwpBridgeWorkspace.SpineSheetName);
                if (spine == null) return;
                if ((int)sheet.Index == (int)spine.Index + 1) return;
                sheet.Move(Type.Missing, spine);
            }
            catch { }
        }

        private static void RenderInto(dynamic sheet, SwpBridgeSession session, SwpSpineRow row, bool fullBlock)
        {
            sheet.Cells.Clear();

            SwpCoordinate coordinate = row.Mapped.Coordinate;
            SwpSheetIndex index = session.Index.Sheets[coordinate.SheetName];

            // The rows on show: the whole device block, or just the clicked section.
            var shown = new List<SwpRowSnapshot>();
            foreach (SwpRowSnapshot bit in index.BitRows)
            {
                if (bit.TableEntry != row.Snapshot.TableEntry) continue;
                if (!fullBlock && bit.TagType != row.Snapshot.TagType) continue;
                shown.Add(bit);
            }

            // Every spine row feeding THIS device: SWP row number → (edits, mnemonic).
            var editsByRow = new Dictionary<int, (List<SwpProposedEdit> Edits, string Mnemonic)>();
            foreach (SwpSpineRow other in session.Model.Rows)
            {
                if (other.Snapshot == null || other.Mapped == null || !other.Mapped.Ok) continue;
                if (other.Mapped.Coordinate.SheetName != coordinate.SheetName) continue;
                if (other.Snapshot.TableEntry != row.Snapshot.TableEntry) continue;
                if (other.Edits.Count == 0 && other != row) continue;
                editsByRow[other.Snapshot.RowNumber] = (other.Edits, other.Entry.Mnemonic);
            }

            sheet.Cells(1, 1).Value =
                "Preview · " + coordinate.FullKey + " · " + row.Entry.Mnemonic +
                " · block \"" + row.Snapshot.BlockLabel + "\" (FU " + row.Snapshot.BlockFUnit +
                ", Loc " + row.Snapshot.BlockLocation + ") · green = value Apply writes · " +
                "scratch view, rebuilt on refresh";
            sheet.Cells(1, 1).Font.Bold = true;

            sheet.Cells(ToggleRow, 1).Value = (fullBlock ? "☑" : "☐") +
                " Show full block — double-click here to toggle (every pending update of this device)";
            sheet.Cells(ToggleRow, 1).Font.Bold = true;
            SwpBridgeSpineSheet.MarkDoubleClickable(sheet.Cells(ToggleRow, 1));

            var headerData = new object[1, Headers.Length];
            for (int c = 0; c < Headers.Length; c++) headerData[0, c] = Headers[c];
            dynamic headerRange = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, Headers.Length));
            headerRange.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, HeaderRow, 1, headerData);
            headerRange.Interior.Color = SwpBridgeSpineSheet.SwpHeaderOrange;   // this is all SWP
            headerRange.Font.Bold = true;

            if (shown.Count == 0)
            {
                AutoFitColumns(sheet, HeaderRow);
                return;
            }

            var data = new object[shown.Count, Headers.Length];
            var greenCells = new List<(int Offset, int Column)>();
            int targetOffset = -1;

            for (int r = 0; r < shown.Count; r++)
            {
                SwpRowSnapshot bit = shown[r];
                bool isTarget = bit.RowNumber == row.Snapshot.RowNumber;
                if (isTarget) targetOffset = r;

                var rowEdits = new Dictionary<SwpField, SwpProposedEdit>();
                string mnemonic = "";
                if (editsByRow.TryGetValue(bit.RowNumber, out var feeding))
                {
                    mnemonic = feeding.Mnemonic;
                    foreach (SwpProposedEdit edit in feeding.Edits) rowEdits[edit.Field] = edit;
                }

                data[r, 0] = bit.RowNumber.ToString();
                data[r, 1] = bit.TagType;
                data[r, 2] = bit.TableEntry.ToString();
                data[r, 3] = bit.Bit.ToString();
                data[r, 4] = mnemonic;
                data[r, 5] = Show(bit.ExportB, rowEdits, SwpField.ExportB, r, greenCells);
                data[r, 6] = Show(bit.AddressG, rowEdits, SwpField.AddressG, r, greenCells);
                data[r, 7] = Show(bit.CachedH, rowEdits, SwpField.FUnitH, r, greenCells);
                data[r, 8] = Show(bit.CachedI, rowEdits, SwpField.LocationI, r, greenCells);
                data[r, 9] = Show(bit.DeviceJ, rowEdits, SwpField.DeviceJ, r, greenCells);
                data[r, 10] = Show(bit.TsRefL, rowEdits, SwpField.TsRefL, r, greenCells);
                data[r, 11] = Show(bit.Desc1T, rowEdits, SwpField.Desc1T, r, greenCells);
                data[r, 12] = Show(bit.Desc2U, rowEdits, SwpField.Desc2U, r, greenCells);
                data[r, 13] = FinalText(session, index, bit, bit.CachedQ, bit.Desc1T,
                    bit.FormulaQ, rowEdits, SwpField.Desc1T, SwpField.FormulaQ);
                data[r, 14] = FinalText(session, index, bit, bit.CachedR, bit.Desc2U,
                    bit.FormulaR, rowEdits, SwpField.Desc2U, SwpField.FormulaR);
                data[r, 15] = bit.StampAA;
            }

            dynamic block = sheet.Range(
                sheet.Cells(HeaderRow + 1, 1), sheet.Cells(HeaderRow + shown.Count, Headers.Length));
            block.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, HeaderRow + 1, 1, data);

            // The "SWP row" column opens the SWP workbook at that row on double-click.
            SwpBridgeSpineSheet.MarkDoubleClickable(sheet.Range(
                sheet.Cells(HeaderRow + 1, 1), sheet.Cells(HeaderRow + shown.Count, 1)));

            if (targetOffset >= 0)
            {
                dynamic targetRow = sheet.Range(
                    sheet.Cells(HeaderRow + 1 + targetOffset, 1),
                    sheet.Cells(HeaderRow + 1 + targetOffset, Headers.Length));
                targetRow.Interior.Color = TargetYellow;
                targetRow.Font.Bold = true;
            }
            foreach ((int offset, int column) in greenCells)
                sheet.Cells(HeaderRow + 1 + offset, column).Interior.Color = NewValueGreen;

            AutoFitColumns(sheet, HeaderRow + shown.Count);
        }

        /// <summary>Every column fits its longest text, header row included — the description
        /// and final-text columns are NOT capped: the preview exists to read them whole (the
        /// title and toggle lines above the header spill freely and do not count).</summary>
        private static void AutoFitColumns(dynamic sheet, int lastRow)
        {
            try
            {
                sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(lastRow, Headers.Length)).Columns.AutoFit();
                for (int c = 1; c <= Headers.Length; c++)
                {
                    dynamic column = sheet.Columns(c);
                    if ((double)column.ColumnWidth < 6) column.ColumnWidth = 6;
                }
            }
            catch { }
        }

        /// <summary>
        /// The Q/R text as it will read AFTER Excel recalculates — the cached final with its
        /// trailing description swapped for the proposed one (plus the pending formula append or
        /// template rebuild when queued). Unchanged rows just show the cached text.
        /// </summary>
        private static string FinalText(SwpBridgeSession session, SwpSheetIndex index,
            SwpRowSnapshot bit, string cached, string currentDesc, string currentFormula,
            Dictionary<SwpField, SwpProposedEdit> edits, SwpField descField, SwpField formulaField)
        {
            bool hasDesc = edits.TryGetValue(descField, out SwpProposedEdit descEdit);
            bool hasFormula = edits.TryGetValue(formulaField, out SwpProposedEdit formulaEdit);
            if (!hasDesc && !hasFormula) return cached;

            // The same test the diff's extension guard uses: does the CURRENT formula display
            // the description cell? (Template-only formulas show it only via a pending append.)
            string descCell = (descField == SwpField.Desc1T ? "T" : "U") + bit.RowNumber;
            bool referencesDesc = SwpDescriptionFormula.ReferencesCell(currentFormula, descCell);

            bool rebuild = hasFormula && SwpDescriptionFormula.IsTruncatingForm(formulaEdit.Proposed);
            string simulated = SwpDiff.SimulateRecalculatedFinal(
                cached, currentDesc,
                hasDesc ? descEdit.Proposed : currentDesc,
                formulaReferencesDesc: referencesDesc,
                formulaAppendPending: hasFormula && !rebuild,
                rebuildPending: rebuild,
                index.Description, bit.BlockLabel);
            return simulated == cached ? cached : simulated + "   ⟵ after recalc";
        }

        /// <summary>A changing cell reads "current → new" and remembers itself for the green paint.</summary>
        private static string Show(string current, Dictionary<SwpField, SwpProposedEdit> edits,
            SwpField field, int rowOffset, List<(int, int)> greenCells)
        {
            if (!edits.TryGetValue(field, out SwpProposedEdit edit)) return current;
            greenCells.Add((rowOffset, PreviewColumnOf(field)));
            return (current ?? "").Trim().Length == 0
                ? "→ " + edit.Proposed
                : current + "  →  " + edit.Proposed;
        }

        private static int PreviewColumnOf(SwpField field)
        {
            switch (field)
            {
                case SwpField.ExportB: return 6;
                case SwpField.AddressG: return 7;
                case SwpField.FUnitH: return 8;
                case SwpField.LocationI: return 9;
                case SwpField.DeviceJ: return 10;
                case SwpField.TsRefL: return 11;
                case SwpField.Desc1T: return 12;
                case SwpField.Desc2U: return 13;
                default: return 0;
            }
        }
    }
}
