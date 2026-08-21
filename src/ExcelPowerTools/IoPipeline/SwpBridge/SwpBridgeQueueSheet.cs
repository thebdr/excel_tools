using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The _Queue sheet: one row per action in queue order, the Include glyph column first
    /// (☑/☐ toggled by double-click, typed values normalized too — the IO Bridge pattern).
    /// </summary>
    internal static class SwpBridgeQueueSheet
    {
        internal const int HeaderRow = 3;
        internal const int IncludeColumn = 1;

        internal const string IncludedGlyph = "☑";
        internal const string ExcludedGlyph = "☐";

        private static readonly string[] Headers =
        {
            "Include", "Status", "Sheet", "Coordinate", "Mnemonic", "Field",
            "Old value", "New value", "Queued (UTC)", "Applied (UTC)", "Note",
        };

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        private static readonly int PendingYellow = Rgb(255, 235, 156);
        private static readonly int ConflictRed = Rgb(255, 199, 206);
        private static readonly int AppliedGreen = Rgb(198, 239, 206);
        private static readonly int ExternalBlue = Rgb(221, 235, 247);

        /// <summary>Creates or replaces the _Queue sheet content. Call inside an ExcelStateGuard.</summary>
        public static void Render(dynamic workbook, SwpBridgeSession session)
        {
            dynamic sheet = FindSheet(workbook, SwpBridgeWorkspace.QueueSheetName);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets(workbook.Worksheets.Count));
                sheet.Name = SwpBridgeWorkspace.QueueSheetName;
            }
            RenderInto(sheet, session);
        }

        /// <summary>Re-renders after a queue mutation.</summary>
        public static void Refresh(SwpBridgeSession session)
        {
            dynamic app = ExcelApp.App;
            dynamic workbook = SwpBridgeWorkspace.FindLiveWorkbook(app);
            if (workbook == null) return;

            using (new ExcelStateGuard(suppressEvents: true))
            {
                Render(workbook, session);
            }
        }

        internal static dynamic FindSheet(dynamic workbook, string name)
        {
            try
            {
                foreach (dynamic ws in workbook.Worksheets)
                    if ((string)ws.Name == name) return ws;
            }
            catch { }
            return null;
        }

        private static void RenderInto(dynamic sheet, SwpBridgeSession session)
        {
            sheet.Cells.Clear();
            try { sheet.Cells.FormatConditions.Delete(); } catch { }

            sheet.Cells(1, 1).Value =
                "SWP Bridge queue · diff-driven: the pending writes recompute on every refresh · " +
                "double-click Include to toggle · unchecked actions are skipped by Apply";
            sheet.Cells(1, 1).Font.Bold = true;

            var headers = new object[1, Headers.Length];
            for (int c = 0; c < Headers.Length; c++) headers[0, c] = Headers[c];
            dynamic headerRange = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, Headers.Length));
            headerRange.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, HeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;

            List<SwpAction> actions = session.Queue.Actions;
            if (actions.Count > 0)
            {
                var data = new object[actions.Count, Headers.Length];
                for (int r = 0; r < actions.Count; r++)
                {
                    SwpAction a = actions[r];
                    data[r, 0] = a.Status == SwpActionStatus.Pending
                        ? (a.Include ? IncludedGlyph : ExcludedGlyph)
                        : "";
                    data[r, 1] = a.Status.ToString();
                    data[r, 2] = a.Sheet;
                    data[r, 3] = a.CoordinateKey;
                    data[r, 4] = a.Mnemonic;
                    data[r, 5] = FieldLabel(a.Field);
                    data[r, 6] = a.OldValue;
                    data[r, 7] = a.NewValue;
                    data[r, 8] = a.QueuedUtc;
                    data[r, 9] = a.AppliedUtc;
                    data[r, 10] = a.StatusNote;
                }

                dynamic block = sheet.Range(
                    sheet.Cells(HeaderRow + 1, 1), sheet.Cells(HeaderRow + actions.Count, Headers.Length));
                block.NumberFormat = "@";
                SwpBridgeSpineSheet.WriteBlock(sheet, HeaderRow + 1, 1, data);

                ApplyStatusFormats(sheet, actions.Count);
                SwpBridgeSpineSheet.MarkDoubleClickable(sheet.Range(
                    sheet.Cells(HeaderRow + 1, IncludeColumn), sheet.Cells(HeaderRow + actions.Count, IncludeColumn)));
            }

            // AutoFit only the used range — whole-sheet AutoFit walks all 16,384 columns.
            try
            {
                int lastRow = HeaderRow + Math.Max(1, actions.Count);
                sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(lastRow, Headers.Length)).Columns.AutoFit();
                sheet.Columns(IncludeColumn).ColumnWidth = 8.5;
            }
            catch { }
        }

        private static void ApplyStatusFormats(dynamic sheet, int actionCount)
        {
            const int XlExpression = 2;
            dynamic block = sheet.Range(
                sheet.Cells(HeaderRow + 1, 1), sheet.Cells(HeaderRow + actionCount, Headers.Length));
            try { block.FormatConditions.Delete(); } catch { }
            try { sheet.Cells(HeaderRow + 1, 1).Select(); } catch { }

            string statusRef = "$B" + (HeaderRow + 1);
            AddRule(block, XlExpression, "=" + statusRef + "=\"Pending\"", PendingYellow);
            AddRule(block, XlExpression, "=OR(" + statusRef + "=\"Conflict\"," + statusRef + "=\"ConflictStale\")", ConflictRed);
            AddRule(block, XlExpression, "=" + statusRef + "=\"Applied\"", AppliedGreen);
            AddRule(block, XlExpression, "=" + statusRef + "=\"AppliedExternally\"", ExternalBlue);
        }

        private static void AddRule(dynamic range, int type, string formula, int color)
        {
            dynamic rule = range.FormatConditions.Add(type, Type.Missing, formula);
            rule.Interior.Color = color;
            rule.StopIfTrue = false;
        }

        /// <summary>The action shown on a 1-based _Queue worksheet row, or null.</summary>
        public static SwpAction ActionAtRow(SwpBridgeSession session, int worksheetRow)
        {
            int index = worksheetRow - HeaderRow - 1;
            return index >= 0 && index < session.Queue.Actions.Count ? session.Queue.Actions[index] : null;
        }

        internal static string FieldLabel(SwpField field)
        {
            switch (field)
            {
                case SwpField.ExportB: return "Export (B)";
                case SwpField.AddressG: return "I/O addr (G)";
                case SwpField.DeviceJ: return "Device (J)";
                case SwpField.TsRefL: return "T.S. ref (L)";
                case SwpField.Desc1T: return "Description 1 (T)";
                case SwpField.Desc2U: return "Description 2 (U)";
                case SwpField.FormulaQ: return "Formula (Q)";
                case SwpField.FormulaR: return "Formula (R)";
                case SwpField.StampAA: return "Stamp (AA)";
                case SwpField.FormulaH: return "Repair FUnit formula (H)";
                case SwpField.FormulaI: return "Repair Location formula (I)";
                case SwpField.HeaderLocationI: return "Repair header Location (I)";
                case SwpField.HeaderFormulaI: return "Repair strobe formula (I)";
                default: return field.ToString();
            }
        }
    }
}
