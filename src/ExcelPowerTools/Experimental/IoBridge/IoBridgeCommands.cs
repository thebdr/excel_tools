using System;
using ExcelDna.Integration;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Experimental.IoBridge
{
    /// <summary>
    /// EXPERIMENTAL — IO Bridge commands. The workspace joins one IO-list workbook's card sheets with
    /// one Inputs + one Outputs plant CSV; edits queue instead of applying, and a batch apply writes
    /// the real files. Sheets are rebuilt fresh from the documents every time; the queue lives in a
    /// JSON sidecar next to the IO list. To remove the whole feature: delete this folder,
    /// Core/Experimental/IoBridge, tests/Experimental/IoBridge, and the pane entries.
    /// </summary>
    public static class IoBridgeCommands
    {
        [ExcelCommand(Name = "PT_XIoBridgeOpen", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: open workspace")]
        public static void OpenWorkspace()
        {
            dynamic app = ExcelApp.App;

            IoBridgeSetupResult setup = IoBridgeSetupForm.Ask();
            if (setup == null) return;

            // One live bridge at a time: a new setup replaces the old session (its queue is on disk).
            dynamic existing = IoBridgeWorkspace.FindLiveWorkbook(app);
            if (existing != null)
            {
                if (!Dialogs.Confirm("An IO Bridge workspace is already open. Close it and build the new one?\n" +
                                     "(Queued actions are safe — they live in the sidecar.)"))
                    return;
                try { existing.Close(false); } catch { }
                IoBridgeSession.Teardown();
            }

            var session = new IoBridgeSession
            {
                IoListPath = setup.IoListPath,
                InputsCsvPath = setup.InputsCsvPath,
                OutputsCsvPath = setup.OutputsCsvPath,
                IoColumnLimit = setup.IoColumnLimit,
                VirtualIncludeRegex = setup.VirtualIncludeRegex,
                ExcludeMnemonicRegex = setup.ExcludeMnemonicRegex,
                SidecarPath = IoBridgeSidecarStore.PathFor(setup.IoListPath),
            };
            session.CardSheets.AddRange(setup.CardSheets);

            // A persisted queue survives sessions — load before the build so re-anchoring sees it.
            try
            {
                var sidecar = IoBridgeSidecarStore.Load(setup.IoListPath);
                if (sidecar != null)
                {
                    session.Queue = sidecar.Queue;
                    if (sidecar.LoadNotes.Count > 0)
                        Dialogs.Info("Queue sidecar notes:\n\n" + string.Join("\n", sidecar.LoadNotes));
                }
            }
            catch (Exception ex)
            {
                Dialogs.Info(ex.Message + "\n\nThe workspace opens with an EMPTY queue; the damaged " +
                             "sidecar is only overwritten when you queue a new action.");
            }

            dynamic workbook = null;
            string workspaceId = null;
            try
            {
                // Curtain up, paint frozen, events guarded — from workbook creation through the whole
                // build the UI stays still, then the finished workspace appears in one repaint.
                using (new ExcelBuildOverlay("IO Bridge — building the workspace…"))
                using (new ExcelPaintFreeze())
                using (new ExcelStateGuard(suppressEvents: true))
                {
                    workbook = IoBridgeWorkspace.CreateWorkbook(app, out workspaceId);
                    session.WorkspaceId = workspaceId;

                    IoBridgeBuilder.BuildWorkspace(app, workbook, session);
                }

                IoBridgeSession.Current = session;
                ManualEditTracker.SheetExclusion = IoBridgeWorkspace.IsBridgeSheet;
                IoBridgeContextMenu.Install();

                int pending = session.Queue.IncludedPending().Count;
                Dialogs.Info("IO Bridge workspace built.\n\n" +
                             $"Card sheets: {string.Join(", ", session.CardSheets)}\n" +
                             $"Queued actions loaded: {session.Queue.Actions.Count}" +
                             (pending > 0 ? $" ({pending} pending and included)" : "") +
                             "\n\nType into IO or CSV data cells to queue changes; the _Queue sheet " +
                             "lists them. Nothing touches the real files until Apply.");
            }
            catch (Exception ex)
            {
                IoBridgeWorkspace.ForgetWorkspace(workspaceId);
                try { workbook?.Close(false); } catch { }
                IoBridgeSession.Teardown();
                Dialogs.Info("The workspace could not be built:\n" + ex.Message);
            }
        }

        [ExcelCommand(Name = "PT_XIoBridgeRowActions", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: row actions")]
        public static void RowActions() => IoBridgeRowActions.ShowForSelection();

        [ExcelCommand(Name = "PT_XIoBridgeExpand", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: expanded view")]
        public static void ExpandView() => IoBridgeExpandView.ShowForSelection();

        [ExcelCommand(Name = "PT_XIoBridgeApply", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: apply queued batch")]
        public static void ApplyBatch() => IoBridgeApplyEngine.Apply();

        [ExcelCommand(Name = "PT_XIoBridgeApplyColumns", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: apply column config")]
        public static void ApplyColumnConfig()
        {
            dynamic app = ExcelApp.App;
            IoBridgeSession session = IoBridgeSession.Current;
            dynamic workbook = IoBridgeWorkspace.FindLiveWorkbook(app);
            if (session == null || workbook == null)
            {
                Dialogs.Info("No live IO Bridge workspace. Use \"IO Bridge: open workspace\" first.");
                return;
            }

            try
            {
                dynamic columnsSheet = null;
                foreach (dynamic ws in workbook.Worksheets)
                    if ((string)ws.Name == IoBridgeWorkspace.ColumnsSheetName) { columnsSheet = ws; break; }
                if (columnsSheet == null)
                {
                    Dialogs.Info("The _Columns sheet is missing — refresh the workspace first.");
                    return;
                }

                var rows = new System.Collections.Generic.List<string[]>();
                int lastRow = (int)RangeHelpers.LastRow(columnsSheet);
                int width = Core.Experimental.IoBridge.IoBridgeColumnLayout.ColumnsSheetHeaders.Length;
                for (int r = 1; r <= lastRow; r++)
                {
                    var line = new string[width];
                    for (int c = 1; c <= width; c++)
                        line[c - 1] = (string)Commands.CsvSheet.CellText(
                            columnsSheet.Cells(r, c).Value2, columnsSheet, r, c);
                    rows.Add(line);
                }

                var config = Core.Experimental.IoBridge.IoBridgeColumnLayout.ParseColumnsSheet(rows);
                IoBridgeRegistry.SaveColumnConfig(config);

                using (new ExcelBuildOverlay("IO Bridge — regrouping the workspace…"))
                using (new ExcelPaintFreeze())
                {
                    IoBridgeBuilder.BuildWorkspace(app, workbook, session);
                }
                Dialogs.Info("Column configuration saved and the workspace regrouped.");
            }
            catch (Exception ex)
            {
                Dialogs.Info("Applying the column configuration failed:\n" + ex.Message);
            }
        }

        [ExcelCommand(Name = "PT_XIoBridgeClearHistory", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: clear applied history")]
        public static void ClearHistory()
        {
            IoBridgeSession session = IoBridgeSession.Current;
            if (session == null)
            {
                Dialogs.Info("No live IO Bridge workspace.");
                return;
            }

            int removed = session.Queue.PurgeHistory();
            session.SaveSidecarNow();
            IoBridgeQueueSheet.Refresh(session);
            Dialogs.Info(removed + " applied action(s) removed from the queue history.");
        }

        [ExcelCommand(Name = "PT_XIoBridgeRefresh", MenuName = "Power Tools", MenuText = "[exp] IO Bridge: refresh from documents")]
        public static void RefreshWorkspace()
        {
            dynamic app = ExcelApp.App;
            IoBridgeSession session = IoBridgeSession.Current;
            dynamic workbook = IoBridgeWorkspace.FindLiveWorkbook(app);
            if (session == null || workbook == null)
            {
                Dialogs.Info("No live IO Bridge workspace. Use \"IO Bridge: open workspace\" first.");
                return;
            }

            try
            {
                using (new ExcelBuildOverlay("IO Bridge — refreshing from the documents…"))
                using (new ExcelPaintFreeze())
                {
                    IoBridgeBuilder.BuildWorkspace(app, workbook, session);
                }
                Dialogs.Info("Workspace rebuilt from the current documents.\n\n" + QueueSummary(session));
            }
            catch (Exception ex)
            {
                Dialogs.Info("The refresh failed:\n" + ex.Message +
                             "\n\nThe queue sidecar is untouched; fix the documents and refresh again.");
            }
        }

        internal static string QueueSummary(IoBridgeSession session)
        {
            int pending = 0, conflicts = 0, external = 0, applied = 0;
            foreach (var action in session.Queue.Actions)
            {
                switch (action.Status)
                {
                    case Core.Experimental.IoBridge.BridgeActionStatus.Pending: pending++; break;
                    case Core.Experimental.IoBridge.BridgeActionStatus.Conflict:
                    case Core.Experimental.IoBridge.BridgeActionStatus.ConflictStale: conflicts++; break;
                    case Core.Experimental.IoBridge.BridgeActionStatus.AppliedExternally: external++; break;
                    case Core.Experimental.IoBridge.BridgeActionStatus.Applied: applied++; break;
                }
            }
            return $"Queue: {pending} pending, {conflicts} in conflict, {external} applied externally, {applied} applied.";
        }
    }
}
