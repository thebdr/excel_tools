using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// Renders the _Queue sheet: one row per action in queue order, the Include checkbox column first.
    /// The checkbox is a text glyph (☑/☐) toggled by double-click — form controls would have to be
    /// mass-deleted and re-created on every rebuild and drift under sorting, so the glyph keeps the
    /// regenerate-fresh model intact. Typed values are normalized by the edit-capture handler too.
    /// </summary>
    internal static class IoBridgeQueueSheet
    {
        internal const int HeaderRow = 3;
        internal const int IncludeColumn = 1;

        internal const string IncludedGlyph = "☑";
        internal const string ExcludedGlyph = "☐";

        private static readonly string[] Headers =
        {
            "Include", "Status", "Card", "Address", "Mnemonic", "Target", "Action",
            "Field", "Old value", "New value", "Queued (UTC)", "Applied (UTC)", "Note",
        };

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        private static readonly int PendingYellow = Rgb(255, 235, 156);
        private static readonly int ConflictRed = Rgb(255, 199, 206);
        private static readonly int AppliedGreen = Rgb(198, 239, 206);
        private static readonly int ExternalBlue = Rgb(221, 235, 247);

        /// <summary>Creates or replaces the _Queue sheet content. Call inside an ExcelStateGuard.</summary>
        public static void Render(dynamic workbook, IoBridgeSession session)
        {
            dynamic sheet = FindQueueSheet(workbook);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets(workbook.Worksheets.Count));
                sheet.Name = IoBridgeWorkspace.QueueSheetName;
            }

            RenderInto(sheet, session);
        }

        /// <summary>Re-renders an existing _Queue sheet after a queue mutation.</summary>
        public static void Refresh(IoBridgeSession session)
        {
            dynamic app = ExcelApp.App;
            dynamic workbook = IoBridgeWorkspace.FindLiveWorkbook(app);
            if (workbook == null) return;

            using (new ExcelStateGuard(suppressEvents: true))
            {
                Render(workbook, session);
            }
        }

        private static dynamic FindQueueSheet(dynamic workbook)
        {
            try
            {
                foreach (dynamic ws in workbook.Worksheets)
                    if ((string)ws.Name == IoBridgeWorkspace.QueueSheetName) return ws;
            }
            catch { }
            return null;
        }

        private static void RenderInto(dynamic sheet, IoBridgeSession session)
        {
            sheet.Cells.Clear();
            try { sheet.Cells.FormatConditions.Delete(); } catch { }

            sheet.Cells(1, 1).Value =
                "IO Bridge queue · double-click the Include cell to toggle · unchecked actions are " +
                "skipped by Apply and stay queued";
            sheet.Cells(1, 1).Font.Bold = true;

            var headers = new object[1, Headers.Length];
            for (int c = 0; c < Headers.Length; c++) headers[0, c] = Headers[c];
            dynamic headerRange = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, Headers.Length));
            headerRange.NumberFormat = "@";
            IoBridgeBuilder.WriteBlock(sheet, HeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;

            List<BridgeAction> actions = session.Queue.Actions;
            if (actions.Count > 0)
            {
                var data = new object[actions.Count, Headers.Length];
                for (int r = 0; r < actions.Count; r++)
                {
                    BridgeAction a = actions[r];
                    data[r, 0] = a.Status == BridgeActionStatus.Pending
                        ? (a.Include ? IncludedGlyph : ExcludedGlyph)
                        : "";
                    data[r, 1] = a.Status.ToString();
                    data[r, 2] = a.Signal.CardKey;
                    data[r, 3] = a.Signal.Address;
                    data[r, 4] = a.Signal.Mnemonic;
                    data[r, 5] = TargetLabel(a.Target);
                    data[r, 6] = KindLabel(a.Kind);
                    data[r, 7] = a.FieldHeader;
                    data[r, 8] = a.OldValue;
                    data[r, 9] = a.NewValue;
                    data[r, 10] = a.QueuedUtc;
                    data[r, 11] = a.AppliedUtc;
                    data[r, 12] = a.StatusNote;
                }

                dynamic block = sheet.Range(
                    sheet.Cells(HeaderRow + 1, 1), sheet.Cells(HeaderRow + actions.Count, Headers.Length));
                block.NumberFormat = "@";
                IoBridgeBuilder.WriteBlock(sheet, HeaderRow + 1, 1, data);

                ApplyStatusFormats(sheet, actions.Count);
            }

            // AutoFit ONLY the used columns — sheet.Columns.AutoFit() walks all 16,384 and takes
            // ~10 seconds, which this refresh pays on EVERY queued edit.
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
        public static BridgeAction ActionAtRow(IoBridgeSession session, int worksheetRow)
        {
            int index = worksheetRow - HeaderRow - 1;
            return index >= 0 && index < session.Queue.Actions.Count ? session.Queue.Actions[index] : null;
        }

        internal static string TargetLabel(BridgeTarget target)
        {
            switch (target)
            {
                case BridgeTarget.InputsCsv: return "Inputs CSV";
                case BridgeTarget.OutputsCsv: return "Outputs CSV";
                default: return "IO list";
            }
        }

        internal static string KindLabel(BridgeActionKind kind)
        {
            switch (kind)
            {
                case BridgeActionKind.SetField: return "Set field";
                case BridgeActionKind.CreateCsvRow: return "Integrate into CSV";
                case BridgeActionKind.DisableCsvRow: return "Disable";
                case BridgeActionKind.EnableCsvRow: return "Enable";
                case BridgeActionKind.DeleteCsvRow: return "Delete row";
                default: return kind.ToString();
            }
        }
    }
}
