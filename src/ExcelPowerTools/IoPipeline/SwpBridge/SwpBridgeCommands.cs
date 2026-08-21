using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// IO Pipeline — SWP Bridge commands. The workspace joins the IO list's diagnostic mnemonics
    /// to the SWP workbook as a reviewable diff: only differing cells become pending actions, the
    /// dashboard shows live tallies, and a headless batch apply patches the SWP bytes (in place
    /// with .bak, or into a new dated version folder). Decisions persist in swp-bridge.json at
    /// the system's SWP root. To remove the whole feature: delete this folder,
    /// Core/IoPipeline/SwpBridge, tests/IoPipeline/SwpBridge, the pane entries, the ribbon group
    /// and the three AddIn.cs event lines.
    /// </summary>
    public static class SwpBridgeCommands
    {
        /// <summary>
        /// A field-debuggable error text: the message chain with exception types, plus the full
        /// stacks dumped to a file next to %TEMP% so "it said X" becomes a pinpointed report.
        /// </summary>
        internal static string Describe(Exception ex)
        {
            var chain = new System.Text.StringBuilder();
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                if (chain.Length > 0) chain.Append("\n← ");
                chain.Append(e.GetType().Name).Append(": ").Append(e.Message);
            }

            string dumpPath = "";
            try
            {
                dumpPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SwpBridge_Error.txt");
                System.IO.File.WriteAllText(dumpPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n" + ex + "\r\n");
            }
            catch { dumpPath = ""; }

            return chain + (dumpPath.Length > 0 ? "\n\nDetails: " + dumpPath : "");
        }

        [ExcelCommand(Name = "PT_SwpBridgeOpen", MenuName = "Power Tools", MenuText = "SWP Bridge: open workspace")]
        public static void OpenWorkspace()
        {
            dynamic app = ExcelApp.App;

            SwpSidecarDocument document;
            string rootFolder;
            using (var form = new SwpBridgeSetupForm())
            {
                if (form.ShowDialog() != DialogResult.OK || form.Result == null) return;
                document = form.Result;
                rootFolder = form.SwpRootFolder;
            }

            dynamic existing = SwpBridgeWorkspace.FindLiveWorkbook(app);
            if (existing != null)
            {
                if (!Dialogs.Confirm("An SWP Bridge workspace is already open. Close it and build the new one?\n" +
                                     "(Decisions are safe — they live in the sidecar.)"))
                    return;
                try { existing.Close(false); } catch { }
                SwpBridgeSession.Teardown();
            }

            var session = new SwpBridgeSession
            {
                SwpRootFolder = rootFolder,
                Sidecar = document,
            };
            if (document.LoadNotes.Count > 0)
                Dialogs.Info("Sidecar notes:\n\n" + string.Join("\n", document.LoadNotes));

            dynamic workbook = null;
            string workspaceId = null;
            Exception failure = null;

            // The sequence the user sees: curtain up IMMEDIATELY on click, the dashboard on top
            // of it narrating every phase, one clean reveal at the end. Dialogs only after the
            // curtain fell (a plain MessageBox would hide beneath the topmost overlay).
            using (new ExcelBuildOverlay("SWP Bridge — building the workspace…"))
            {
                session.Dashboard = SwpBridgeDashboard.Show();
                try
                {
                    SwpBridgeBuilder.BuildModel(app, session);

                    using (new ExcelPaintFreeze())
                    using (new ExcelStateGuard(suppressEvents: true, manualCalculation: true))
                    {
                        workbook = SwpBridgeWorkspace.CreateWorkbook(app, out workspaceId);
                        session.WorkspaceId = workspaceId;
                        SwpBridgeBuilder.RenderWorkspace(workbook, session);
                    }

                    SwpBridgeSession.Current = session;
                    session.InstallSheetExclusion();
                    session.SaveSidecarNow();
                    session.Dashboard?.PostModel(session);
                }
                catch (Exception ex)
                {
                    failure = ex;
                    SwpBridgeWorkspace.ForgetWorkspace(workspaceId);
                    try { workbook?.Close(false); } catch { }
                    try { session.Dashboard?.CloseFromExcel(); } catch { }
                    session.Dashboard = null;
                }
            }

            if (failure != null)
                Dialogs.Info(failure is OperationCanceledException
                    ? "Build cancelled."
                    : "The workspace could not be built:\n" + Describe(failure));
        }

        [ExcelCommand(Name = "PT_SwpBridgeRefresh", MenuName = "Power Tools", MenuText = "SWP Bridge: refresh from documents")]
        public static void RefreshWorkspace()
        {
            dynamic app = ExcelApp.App;
            SwpBridgeSession session = SwpBridgeSession.Current;
            dynamic workbook = SwpBridgeWorkspace.FindLiveWorkbook(app);
            if (session == null || workbook == null)
            {
                Dialogs.Info("No live SWP Bridge workspace. Use \"SWP Bridge: open workspace\" first.");
                return;
            }

            Exception failure = null;
            using (new ExcelBuildOverlay("SWP Bridge — refreshing…"))
            {
                try
                {
                    SwpBridgeBuilder.BuildModel(app, session);
                    using (new ExcelPaintFreeze())
                    using (new ExcelStateGuard(suppressEvents: true, manualCalculation: true))
                    {
                        session.PreviewContext = null;
                        SwpBridgeBuilder.RenderWorkspace(workbook, session);
                    }
                    session.SaveSidecarNow();
                }
                catch (Exception ex)
                {
                    failure = ex;
                    session.Dashboard?.PostPhase("Ready");
                }
            }

            if (failure != null)
                Dialogs.Info(failure is OperationCanceledException
                    ? "Refresh cancelled."
                    : "The refresh failed:\n" + Describe(failure) +
                      "\n\nThe sidecar is untouched; fix the documents and refresh again.");
        }

        [ExcelCommand(Name = "PT_SwpBridgeApply", MenuName = "Power Tools", MenuText = "SWP Bridge: apply queued batch")]
        public static void ApplyBatch()
        {
            dynamic app = ExcelApp.App;
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null || session.Model == null)
            {
                Dialogs.Info("No live SWP Bridge workspace. Use \"SWP Bridge: open workspace\" first.");
                return;
            }

            try
            {
                SwpBridgeApplyEngine.Apply(app, session);

                // Rebuild so the workspace shows the post-apply reality (also after conflicts).
                dynamic workbook = SwpBridgeWorkspace.FindLiveWorkbook(app);
                if (workbook != null)
                {
                    SwpBridgeBuilder.BuildModel(app, session);
                    using (new ExcelBuildOverlay("SWP Bridge — rendering…"))
                    using (new ExcelPaintFreeze())
                    using (new ExcelStateGuard(suppressEvents: true, manualCalculation: true))
                    {
                        session.PreviewContext = null;
                        SwpBridgeBuilder.RenderWorkspace(workbook, session);
                    }
                    session.SaveSidecarNow();
                }
            }
            catch (Exception ex)
            {
                session.Dashboard?.PostPhase("Ready");
                Dialogs.Info("Apply failed:\n" + Describe(ex));
            }
        }

        [ExcelCommand(Name = "PT_SwpBridgePreview", MenuName = "Power Tools", MenuText = "SWP Bridge: preview SWP block")]
        public static void PreviewSelection()
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null || session.Model == null)
            {
                Dialogs.Info("No live SWP Bridge workspace.");
                return;
            }
            try
            {
                dynamic app = ExcelApp.App;
                dynamic sheet = app.ActiveSheet;
                if ((string)sheet.Name != SwpBridgeWorkspace.SpineSheetName)
                {
                    Dialogs.Info("Select a row on the Spine sheet first.");
                    return;
                }
                int row = (int)app.ActiveCell.Row;
                if (!session.ModelRowByWorksheetRow.TryGetValue(row, out int modelRow))
                {
                    Dialogs.Info("The selection is not on a data row.");
                    return;
                }
                SwpBridgePreview.Open(session, modelRow);
            }
            catch (Exception ex)
            {
                Dialogs.Info("Preview failed:\n" + Describe(ex));
            }
        }

        [ExcelCommand(Name = "PT_SwpBridgeDashboard", MenuName = "Power Tools", MenuText = "SWP Bridge: show dashboard")]
        public static void ShowDashboard()
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null)
            {
                Dialogs.Info("No live SWP Bridge workspace.");
                return;
            }
            if (session.Dashboard == null) session.Dashboard = SwpBridgeDashboard.Show();
            session.Dashboard.PostPhase("Ready");
            session.Dashboard.PostModel(session);
            session.Dashboard.BringToFront();
        }

        [ExcelCommand(Name = "PT_SwpBridgeClearHistory", MenuName = "Power Tools", MenuText = "SWP Bridge: clear applied history")]
        public static void ClearHistory()
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null)
            {
                Dialogs.Info("No live SWP Bridge workspace.");
                return;
            }
            int removed = session.Queue.PurgeHistory();
            session.SaveSidecarNow();
            SwpBridgeQueueSheet.Refresh(session);
            Dialogs.Info(removed + " applied action(s) removed from the queue history.");
        }

        // ---- dashboard dispatches (Excel thread, macro context) -----------------------------

        internal static void FilterSpineToSheet(string sheetName)
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            dynamic workbook = session == null ? null : SwpBridgeWorkspace.FindLiveWorkbook(ExcelApp.App);
            if (workbook == null) return;
            try
            {
                dynamic spine = workbook.Worksheets(SwpBridgeWorkspace.SpineSheetName);
                spine.Activate();
                int lastRow = SwpBridgeSpineSheet.FirstDataRow + Math.Max(0, session.Model.Rows.Count - 1);
                dynamic range = spine.Range(
                    spine.Cells(SwpBridgeSpineSheet.HeaderRow, 1),
                    spine.Cells(lastRow, SwpBridgeSpineSheet.ColumnCount));
                if (sheetName == "(unmapped)")
                    range.AutoFilter(SwpBridgeSpineSheet.ColSheet, "=");
                else
                    range.AutoFilter(SwpBridgeSpineSheet.ColSheet, sheetName);
            }
            catch { }
        }

        internal static void JumpToMnemonic(string mnemonic)
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            dynamic workbook = session == null ? null : SwpBridgeWorkspace.FindLiveWorkbook(ExcelApp.App);
            if (workbook == null) return;
            try
            {
                for (int i = 0; i < session.Model.Rows.Count; i++)
                {
                    if (session.Model.Rows[i].Entry.Mnemonic != mnemonic) continue;
                    if (!session.WorksheetRowByModelRow.TryGetValue(i, out int row)) return;
                    dynamic spine = workbook.Worksheets(SwpBridgeWorkspace.SpineSheetName);
                    spine.Activate();
                    ExcelApp.App.Goto(spine.Cells(row, SwpBridgeSpineSheet.ColMnemonic), true);
                    return;
                }
            }
            catch { }
        }
    }
}
