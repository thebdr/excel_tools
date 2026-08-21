using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using ExcelPowerTools.Core.Files;
using ExcelPowerTools.Interop;
using CoreIoListLogic = ExcelPowerTools.Core.IoPipeline.IoListLogic;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// The batch apply: included pending actions, grouped per target document. CSVs are patched
    /// headlessly (three-way per action, round-trip verified, TOCTOU-checked, atomic backup write);
    /// the IO-list .xlsx is patched via COM with rows located by WorksheetFunction (never whole-sheet
    /// Find). Statuses flip only after a target's write actually happened, and the workspace is
    /// rebuilt afterwards so the sheets show the new reality.
    /// </summary>
    internal static class IoBridgeApplyEngine
    {
        private sealed class TargetTally
        {
            public int Applied, Conflicts, External, Skipped;
            public readonly List<string> Notes = new List<string>();
            public bool FileWritten;
        }

        public static void Apply()
        {
            dynamic app = ExcelApp.App;
            IoBridgeSession session = IoBridgeSession.Current;
            dynamic workbook = IoBridgeWorkspace.FindLiveWorkbook(app);
            if (session == null || workbook == null)
            {
                Dialogs.Info("No live IO Bridge workspace. Use \"IO Bridge: open workspace\" first.");
                return;
            }

            List<BridgeAction> pending = session.Queue.IncludedPending();
            if (pending.Count == 0)
            {
                Dialogs.Info("Nothing to apply — no pending, included actions in the queue.");
                return;
            }

            var byTarget = new Dictionary<BridgeTarget, List<BridgeAction>>();
            foreach (BridgeAction action in pending)
            {
                if (!byTarget.TryGetValue(action.Target, out List<BridgeAction> list))
                    byTarget[action.Target] = list = new List<BridgeAction>();
                list.Add(action);
            }

            var counts = new StringBuilder();
            foreach (KeyValuePair<BridgeTarget, List<BridgeAction>> group in byTarget)
                counts.AppendLine("  " + IoBridgeQueueSheet.TargetLabel(group.Key) + ": " + group.Value.Count + " action(s)");
            if (!Dialogs.Confirm("Apply the queued batch to the real files?\n\n" + counts +
                                 "\nConflicting actions are skipped and flagged; the rest proceed."))
                return;

            var tallies = new Dictionary<BridgeTarget, TargetTally>();
            if (byTarget.TryGetValue(BridgeTarget.InputsCsv, out List<BridgeAction> inputActions))
                tallies[BridgeTarget.InputsCsv] = ApplyCsvTarget(session.InputsCsvPath, inputActions);
            if (byTarget.TryGetValue(BridgeTarget.OutputsCsv, out List<BridgeAction> outputActions))
                tallies[BridgeTarget.OutputsCsv] = ApplyCsvTarget(session.OutputsCsvPath, outputActions);
            if (byTarget.TryGetValue(BridgeTarget.IoList, out List<BridgeAction> ioActions))
                tallies[BridgeTarget.IoList] = ApplyIoListTarget(app, session, ioActions);

            session.SaveSidecarNow();

            // The documents just changed on disk — rebuild so the sheets and flags show the new truth.
            string rebuildNote = "";
            try
            {
                using (new ExcelBuildOverlay("IO Bridge — applying and rebuilding…"))
                using (new ExcelPaintFreeze())
                {
                    IoBridgeBuilder.BuildWorkspace(app, workbook, session);
                }
            }
            catch (Exception ex)
            {
                rebuildNote = "\n\nThe post-apply rebuild failed (" + ex.Message +
                              ") — run \"IO Bridge: refresh\" once the documents are readable.";
            }

            var summary = new StringBuilder("Apply finished.\n");
            foreach (KeyValuePair<BridgeTarget, TargetTally> pair in tallies)
            {
                TargetTally t = pair.Value;
                summary.AppendLine();
                summary.AppendLine(IoBridgeQueueSheet.TargetLabel(pair.Key) + ":");
                summary.AppendLine("  applied " + t.Applied + " · conflicts " + t.Conflicts +
                                   " · already done " + t.External + " · skipped " + t.Skipped +
                                   (t.FileWritten ? " · file written (.bak kept)" : " · file NOT written"));
                foreach (string note in t.Notes) summary.AppendLine("  " + note);
            }
            summary.Append(rebuildNote);
            Dialogs.Info(summary.ToString());
        }

        // ---- CSV targets -----------------------------------------------------------------------------

        private static TargetTally ApplyCsvTarget(string path, List<BridgeAction> actions)
        {
            var tally = new TargetTally();

            // Contention: the delimited workspace owns this file's save pipeline while a session is open.
            if (CsvSession.FindByPath(path) != null)
            {
                tally.Skipped = actions.Count;
                tally.Notes.Add("open in the delimited-file workspace — close it there first");
                return tally;
            }

            SourceReadState state = AtomicFileStore.ReadSource(path, out byte[] bytes, out string readError);
            if (state != SourceReadState.Present)
            {
                tally.Skipped = actions.Count;
                tally.Notes.Add("file could not be read" + (string.IsNullOrEmpty(readError) ? "" : ": " + readError));
                MarkAll(actions, BridgeActionStatus.ConflictStale, "file missing or unreadable at apply time");
                return tally;
            }

            CsvDocument doc;
            CsvColumnsMap map;
            try
            {
                doc = CsvFile.Load(bytes);
                map = CsvColumnsMap.Locate(doc);
            }
            catch (Exception ex)
            {
                tally.Skipped = actions.Count;
                tally.Notes.Add("file could not be parsed: " + ex.Message);
                return tally;
            }
            if (map == null)
            {
                tally.Skipped = actions.Count;
                tally.Notes.Add("not a plant CSV any more (key columns missing)");
                return tally;
            }

            CsvPatchResult result = CsvPatchApplier.Apply(doc, actions, map);

            if (result.AnyMutation)
            {
                byte[] outBytes = doc.Save();
                string verifyError = CsvDiff.VerifyRoundTrip(outBytes, doc.Format, doc.Records);
                if (verifyError != null)
                {
                    tally.Skipped = actions.Count;
                    tally.Notes.Add("write aborted — round-trip verification failed: " + verifyError);
                    return tally;   // statuses untouched: everything stays Pending
                }
                if (!AtomicFileStore.DiskStillMatches(path, expectedExists: true, expectedBytes: bytes))
                {
                    tally.Skipped = actions.Count;
                    tally.Notes.Add("write aborted — the file changed while applying; refresh and retry");
                    return tally;
                }

                try
                {
                    AtomicFileStore.WriteWithBackup(path, outBytes);
                }
                catch (Exception ex)
                {
                    tally.Skipped = actions.Count;
                    tally.Notes.Add("write failed: " + ex.Message);
                    return tally;
                }
                tally.FileWritten = true;
            }

            // The write (if any) succeeded — NOW the outcomes become statuses.
            string stamp = DateTime.UtcNow.ToString("O");
            foreach (CsvActionOutcome outcome in result.Outcomes)
            {
                outcome.Action.Status = outcome.Outcome;
                outcome.Action.StatusNote = outcome.Note;
                switch (outcome.Outcome)
                {
                    case BridgeActionStatus.Applied:
                        outcome.Action.AppliedUtc = stamp;
                        tally.Applied++;
                        break;
                    case BridgeActionStatus.AppliedExternally:
                        tally.External++;
                        break;
                    default:
                        tally.Conflicts++;
                        if (tally.Notes.Count < 4)
                            tally.Notes.Add("conflict: " + outcome.Action.Signal.Mnemonic + " — " + outcome.Note);
                        break;
                }
            }
            return tally;
        }

        private static void MarkAll(List<BridgeAction> actions, BridgeActionStatus status, string note)
        {
            foreach (BridgeAction action in actions)
            {
                action.Status = status;
                action.StatusNote = note;
            }
        }

        // ---- the IO-list target ----------------------------------------------------------------------

        private static TargetTally ApplyIoListTarget(dynamic app, IoBridgeSession session, List<BridgeAction> actions)
        {
            var tally = new TargetTally();

            dynamic wb = IoBridgeBuilder.FindOpenWorkbook(app, session.IoListPath);
            bool wasAlreadyOpen = wb != null;
            bool opened = false;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }

            try
            {
                if (wb == null)
                {
                    try
                    {
                        wb = app.Workbooks.Open(session.IoListPath, 0, false);   // read-write
                        opened = true;
                    }
                    catch (Exception ex)
                    {
                        tally.Skipped = actions.Count;
                        tally.Notes.Add("workbook locked or unreadable: " + ex.Message);
                        return tally;
                    }
                }
                if ((bool)wb.ReadOnly)
                {
                    tally.Skipped = actions.Count;
                    tally.Notes.Add("the workbook is open read-only — nothing was written");
                    return tally;
                }

                var patched = new List<BridgeAction>();
                using (new ExcelStateGuard(suppressEvents: true, manualCalculation: true))
                {
                    foreach (BridgeAction action in actions)
                    {
                        string note = PatchIoListAction(wb, session, action);
                        if (note == null)
                        {
                            patched.Add(action);
                        }
                        else if (note.Length == 0)
                        {
                            action.Status = BridgeActionStatus.AppliedExternally;
                            action.StatusNote = "the workbook already holds the new value";
                            tally.External++;
                        }
                        else
                        {
                            action.Status = BridgeActionStatus.Conflict;
                            action.StatusNote = note;
                            tally.Conflicts++;
                            if (tally.Notes.Count < 4)
                                tally.Notes.Add("conflict: " + action.Signal.Mnemonic + " — " + note);
                        }
                    }
                }

                if (patched.Count == 0) return tally;

                bool saved = false;
                if (opened)
                {
                    wb.Save();
                    saved = true;
                }
                else
                {
                    // Never silently save a workbook the user had open — they may hold other edits.
                    try { app.DisplayAlerts = prevAlerts; } catch { }
                    saved = Dialogs.Confirm(
                        "The IO list is open in this Excel. " + patched.Count + " cell(s) were patched " +
                        "in the open workbook.\n\nSave \"" + Path.GetFileName(session.IoListPath) + "\" now?");
                    try { app.DisplayAlerts = false; } catch { }
                    if (saved) wb.Save();
                }

                string stamp = DateTime.UtcNow.ToString("O");
                foreach (BridgeAction action in patched)
                {
                    if (saved)
                    {
                        action.Status = BridgeActionStatus.Applied;
                        action.StatusNote = "";
                        action.AppliedUtc = stamp;
                        tally.Applied++;
                    }
                    else
                    {
                        action.StatusNote = "patched in the open workbook — not yet saved";
                        tally.Skipped++;
                    }
                }
                tally.FileWritten = saved;
                if (!saved && !wasAlreadyOpen) tally.Notes.Add("changes were not saved");
                if (!saved && wasAlreadyOpen)
                    tally.Notes.Add("save declined — the actions stay pending until the workbook is saved");
                return tally;
            }
            finally
            {
                // Close(false) always: a successful path already called Save(); a failed one must
                // DISCARD partial patches so the file matches the untouched action statuses.
                if (opened) try { wb?.Close(false); } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
            }
        }

        /// <summary>
        /// Patches one action into the open IO-list workbook. Returns null on success, "" when the
        /// cell already holds the new value, or a conflict note. Row location: the build-time source
        /// row re-verified, then MATCH on the raw Bit text (visibility-proof), then the mnemonic —
        /// with COUNTIF uniqueness so an ambiguous list is an error, not a guess.
        /// </summary>
        private static string PatchIoListAction(dynamic wb, IoBridgeSession session, BridgeAction action)
        {
            BridgeRow row = session.Model.FindByIdentity(action.Signal);
            if (row == null || row.Kind == BridgeRowKind.CsvOnly) return "signal not in the built model";

            BridgeSheetModel sheetModel = FindSheetModel(session, row);
            if (sheetModel == null) return "card sheet not in this workspace";

            dynamic sheet = FindWorksheet(wb, sheetModel.SheetName);
            if (sheet == null) return "sheet \"" + sheetModel.SheetName + "\" not found in the workbook";

            int fieldColumn = FindIoColumn(sheetModel.IoHeaders, action.FieldHeader);
            if (fieldColumn < 0) return "column \"" + action.FieldHeader + "\" not found on the sheet";

            int rowNumber = LocateRow(sheet, sheetModel, row);
            if (rowNumber == -1) return "row not found in the workbook (the list changed?)";
            if (rowNumber == -2) return "the address matches more than one row";

            dynamic cell = sheet.Cells(rowNumber, fieldColumn + 1);
            string current = (string)CsvSheet.CellText(cell.Value2, sheet, rowNumber, fieldColumn + 1);

            if (current == action.NewValue) return "";
            if (current != action.OldValue) return "cell now holds \"" + current + "\"";

            // Minimal formatting touch: force text only when the value would otherwise coerce.
            if (WouldCoerce(action.NewValue)) try { cell.NumberFormat = "@"; } catch { }
            cell.Value = action.NewValue;
            return null;
        }

        private static bool WouldCoerce(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            if (value.StartsWith("=", StringComparison.Ordinal) ||
                value.StartsWith("+", StringComparison.Ordinal) ||
                value.StartsWith("-", StringComparison.Ordinal)) return true;
            return double.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out _) ||
                DateTime.TryParse(value, out _);
        }

        private static BridgeSheetModel FindSheetModel(IoBridgeSession session, BridgeRow row)
        {
            foreach (BridgeSheetModel sheet in session.Model.Sheets)
                if (sheet.SpineRows.Contains(row)) return sheet;
            return null;
        }

        private static dynamic FindWorksheet(dynamic wb, string name)
        {
            try
            {
                foreach (dynamic ws in wb.Worksheets)
                    if (string.Equals((string)ws.Name, name, StringComparison.OrdinalIgnoreCase)) return ws;
            }
            catch { }
            return null;
        }

        private static int FindIoColumn(string[] headers, string normalizedHeader)
        {
            for (int c = 0; c < headers.Length; c++)
                if (CoreIoListLogic.NormalizeHeader(headers[c]) == normalizedHeader) return c;
            return -1;
        }

        /// <summary>1-based worksheet row of the action's signal; -1 not found, -2 ambiguous.</summary>
        private static int LocateRow(dynamic sheet, BridgeSheetModel model, BridgeRow row)
        {
            dynamic wf = ExcelApp.App.WorksheetFunction;
            int bitColumn = model.BitColumn + 1;

            // The build-time row, verified — rows may have been inserted since.
            if (row.IoSourceRow > 0 && RowStillMatches(sheet, model, row, row.IoSourceRow))
                return row.IoSourceRow;

            if (row.Kind == BridgeRowKind.Signal && row.RawBitText.Trim().Length > 0)
            {
                string escaped = EscapeMatchValue(row.RawBitText);
                int count;
                try { count = Convert.ToInt32(wf.CountIf(sheet.Columns(bitColumn), escaped)); }
                catch { count = -1; }
                if (count > 1) return -2;
                if (count == 1)
                {
                    try { return Convert.ToInt32(wf.Match(escaped, sheet.Columns(bitColumn), 0)); }
                    catch { }
                }
            }

            // Module rows (and re-addressed signals): the mnemonic column.
            int mnemonicColumn = FindIoColumn(model.IoHeaders, "mnemonic") + 1;
            if (mnemonicColumn > 0 && row.Signal.Mnemonic.Length > 0)
            {
                string escaped = EscapeMatchValue(row.Signal.Mnemonic);
                int count;
                try { count = Convert.ToInt32(wf.CountIf(sheet.Columns(mnemonicColumn), escaped)); }
                catch { count = -1; }
                if (count > 1) return -2;
                if (count == 1)
                {
                    try { return Convert.ToInt32(wf.Match(escaped, sheet.Columns(mnemonicColumn), 0)); }
                    catch { }
                }
            }
            return -1;
        }

        private static bool RowStillMatches(dynamic sheet, BridgeSheetModel model, BridgeRow row, int rowNumber)
        {
            try
            {
                int bitColumn = model.BitColumn + 1;
                string bitText = (string)CsvSheet.CellText(
                    sheet.Cells(rowNumber, bitColumn).Value2, sheet, rowNumber, bitColumn);

                if (row.Kind == BridgeRowKind.Signal)
                    return bitText.Trim() == row.RawBitText.Trim();

                int idColumn = FindIoColumn(model.IoHeaders, "id") + 1;
                if (idColumn <= 0) return false;
                string idText = (string)CsvSheet.CellText(
                    sheet.Cells(rowNumber, idColumn).Value2, sheet, rowNumber, idColumn);
                return CoreIoListLogic.ModuleAddress(idText) == row.Signal.Address;
            }
            catch { return false; }
        }

        private static string EscapeMatchValue(string value) =>
            value.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");
    }
}
