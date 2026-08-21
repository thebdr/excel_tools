using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelPowerTools.Core.Files;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// Applies the included pending actions to the SWP — headless, three-way checked, verified
    /// before written, atomic on disk. Two modes: in-place on the current file (with .bak), or
    /// version-folder (the next NN_dd-MMM-yyyy sibling is created, SWP + CONFIG_PLANT copied,
    /// the COPY patched — the original is never written). Statuses flip only after the disk
    /// write returns; a findings .txt lands next to the SWP; the workspace then rebuilds.
    /// </summary>
    internal static class SwpBridgeApplyEngine
    {
        public static void Apply(dynamic app, SwpBridgeSession session)
        {
            List<SwpAction> included = session.Queue.IncludedPending();
            if (included.Count == 0)
            {
                Dialogs.Info("Nothing to apply — no included pending action.", "SWP Bridge");
                return;
            }

            string sourcePath = session.SwpPath;
            if (SwpBridgeCardReader.FindOpenWorkbook(app, sourcePath) != null)
            {
                Dialogs.Info(
                    "The SWP is open in this Excel:\n" + sourcePath +
                    "\n\nClose it first — the bridge patches the file on disk.",
                    "SWP Bridge");
                return;
            }

            var perSheet = new SortedDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (SwpAction action in included)
                perSheet[action.Sheet] = (perSheet.TryGetValue(action.Sheet, out int n) ? n : 0) + 1;
            var summary = new StringBuilder();
            foreach (var pair in perSheet) summary.Append("\n  · ").Append(pair.Key).Append(": ").Append(pair.Value);

            bool versionFolder = session.Settings.ApplyToVersionFolder;
            string question = "Apply " + included.Count + " action(s)?" + summary + "\n\n" +
                (versionFolder
                    ? "Mode: a NEW dated version folder is created and patched; the current file stays untouched."
                    : "Mode: in-place on\n" + sourcePath + "\n(a .bak of the previous state is kept).");
            if (!Dialogs.OkCancel(question, "SWP Bridge — apply")) return;

            // TOCTOU: the file must still be what the diff was computed against.
            byte[] currentBytes = SwpBridgeBuilder.ReadFileBytes(sourcePath);
            if (!SameBytes(currentBytes, session.SwpBytes))
            {
                Dialogs.Info(
                    "The SWP changed on disk since the workspace was built.\n" +
                    "Refresh the workspace and review the new diff before applying.",
                    "SWP Bridge");
                return;
            }

            session.Dashboard?.PostPhase("Applying " + included.Count + " action(s)…");
            string stamp = DateTime.Now.ToString("yyyy-MM-dd.HH:mm:ss", CultureInfo.InvariantCulture);
            var log = new List<SwpLogEntry>();
            SwpXlsxPatcher.PatchResult result = SwpXlsxPatcher.Apply(currentBytes, included, stamp, log);
            session.BuildLog.AddRange(log);

            string targetPath = sourcePath;
            int applied = 0, conflicts = 0, external = 0;
            foreach (SwpXlsxPatcher.Outcome outcome in result.Outcomes)
            {
                if (outcome.Result == SwpActionStatus.Applied) applied++;
                else if (outcome.Result == SwpActionStatus.Conflict) conflicts++;
                else if (outcome.Result == SwpActionStatus.AppliedExternally) external++;
            }

            if (result.AppliedCount > 0)
            {
                if (!VerifyPatched(result, session))
                {
                    session.Dashboard?.PostPhase("Ready");
                    Dialogs.Info(
                        "The patched workbook did not verify — NOTHING was written.\n" +
                        "The queue is unchanged; refresh and retry.",
                        "SWP Bridge");
                    return;
                }

                if (versionFolder)
                {
                    targetPath = CreateNextVersionFolder(session);
                    AtomicFileStore.WriteWithBackup(targetPath, result.PatchedBytes);
                }
                else
                {
                    if (!AtomicFileStore.DiskStillMatches(sourcePath, expectedExists: true, expectedBytes: currentBytes))
                    {
                        session.Dashboard?.PostPhase("Ready");
                        Dialogs.Info(
                            "The SWP changed on disk while applying — NOTHING was written.\n" +
                            "Refresh the workspace and retry.",
                            "SWP Bridge");
                        return;
                    }
                    AtomicFileStore.WriteWithBackup(sourcePath, result.PatchedBytes);
                }
            }

            // The write is real — only now do the statuses flip.
            string appliedUtc = DateTime.UtcNow.ToString("O");
            foreach (SwpXlsxPatcher.Outcome outcome in result.Outcomes)
            {
                outcome.Action.Status = outcome.Result;
                outcome.Action.StatusNote = outcome.Note;
                if (outcome.Result == SwpActionStatus.Applied ||
                    outcome.Result == SwpActionStatus.AppliedExternally)
                    outcome.Action.AppliedUtc = appliedUtc;
            }
            session.SaveSidecarNow();
            ExportLogText(session, targetPath);

            string message = "Applied " + applied + " · already there " + external + " · conflicts " + conflicts +
                (versionFolder && result.AppliedCount > 0
                    ? "\n\nNew version:\n" + targetPath
                    : "") +
                (conflicts > 0 ? "\n\nConflicted actions stay queued — the refreshed diff shows the new reality." : "");
            Dialogs.Info(message, "SWP Bridge — apply");
        }

        /// <summary>Re-reads the patched bytes and confirms every applied action landed.</summary>
        private static bool VerifyPatched(SwpXlsxPatcher.PatchResult result, SwpBridgeSession session)
        {
            try
            {
                var sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (SwpXlsxPatcher.Outcome outcome in result.Outcomes)
                    if (outcome.Result == SwpActionStatus.Applied) sheets.Add(outcome.Action.Sheet);

                SwpWorkbookIndex reread = SwpSheetReader.Read(result.PatchedBytes, sheets, new List<SwpLogEntry>());
                foreach (SwpXlsxPatcher.Outcome outcome in result.Outcomes)
                {
                    if (outcome.Result != SwpActionStatus.Applied) continue;
                    string current = SwpBridgeBuilder.ResolveCurrent(reread, outcome.Action);
                    if (current == null ||
                        !SwpFieldCompare.Same(outcome.Action.Field, current, outcome.Action.NewValue))
                        return false;
                }
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// The legacy folder convention: NN_dd-MMM-yyyy siblings under the SWP root, NN one past
        /// the highest existing. SWP and CONFIG_PLANT are copied in (the external [1] link is
        /// folder-relative, so the copy keeps resolving); the session then works on the new copy.
        /// </summary>
        private static string CreateNextVersionFolder(SwpBridgeSession session)
        {
            string root = session.SwpRootFolder;
            int highest = 0;
            foreach (string dir in Directory.GetDirectories(root))
            {
                Match m = Regex.Match(Path.GetFileName(dir), @"^(\d+)_");
                if (m.Success && int.TryParse(m.Groups[1].Value, out int n) && n > highest) highest = n;
            }

            string folderName = (highest + 1).ToString("00", CultureInfo.InvariantCulture) + "_" +
                DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.CurrentCulture);
            string folder = Path.Combine(root, folderName);
            Directory.CreateDirectory(folder);

            string oldFolder = Path.Combine(root, session.Settings.CurrentVersionFolder ?? "");
            string configName = session.Settings.ConfigFileName;
            if (!string.IsNullOrEmpty(configName) && File.Exists(Path.Combine(oldFolder, configName)))
                File.Copy(Path.Combine(oldFolder, configName), Path.Combine(folder, configName), overwrite: true);

            session.Settings.CurrentVersionFolder = folderName;
            return Path.Combine(folder, session.Settings.SwpFileName);
        }

        private static void ExportLogText(SwpBridgeSession session, string swpPath)
        {
            try
            {
                string path = Path.Combine(Path.GetDirectoryName(swpPath) ?? session.SwpRootFolder,
                    "SwpBridge_Log.txt");
                File.WriteAllLines(path, SwpBridgeAuxSheets.LogLines(session), Encoding.UTF8);
            }
            catch { }
        }

        private static bool SameBytes(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            return true;
        }
    }
}
