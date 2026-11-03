using System;
using System.Collections.Generic;
using System.IO;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The one live SWP Bridge session of this Excel process: the sidecar document (settings,
    /// overrides, dictionaries, templates, queue), the built model, the SWP index and build-time
    /// bytes (the TOCTOU base), the harvested IO rows (so override/dictionary edits recompute the
    /// diff WITHOUT re-reading any file), and the spine row map. Everything user-decided persists
    /// through <see cref="SaveSidecarNow"/> on every mutation.
    /// </summary>
    internal sealed class SwpBridgeSession
    {
        public static SwpBridgeSession Current;

        public string WorkspaceId = "";
        public string SwpRootFolder = "";

        /// <summary>The session FILE every decision is written back to. Normally
        /// <c>SwpRootFolder\swp-bridge.json</c>; a named session ("Save session as…") puts it
        /// wherever the user chose, which is what lets one IO list drive two SWP systems.</summary>
        public string SidecarPath = "";

        /// <summary>The file decisions are written to right now: the named session's, or the root's
        /// default sidecar when none was named. Save as moves it; nothing else does.</summary>
        public string EffectiveSidecarPath =>
            SidecarPath.Length > 0 ? SidecarPath : SwpBridgeSidecarStore.PathFor(SwpRootFolder);

        public SwpSidecarDocument Sidecar = new SwpSidecarDocument();
        public SwpBridgeSettings Settings => Sidecar.Settings;
        public SwpQueue Queue => Sidecar.Queue;

        /// <summary>The harvested source rows (IO list + generators), the in-memory diff input.</summary>
        public List<IoRowData> IoRows = new List<IoRowData>();
        public List<DiagnosticEntry> Entries = new List<DiagnosticEntry>();

        public SwpWorkbookIndex Index;

        /// <summary>The reference SWP's index for the formula audit (null = no audit).</summary>
        public SwpWorkbookIndex BaseIndex;

        /// <summary>The alarm-strings files' raw lines, cached at build (path → lines).</summary>
        public List<(string Path, string[] Lines)> AlarmStringsFiles = new List<(string, string[])>();

        /// <summary>The index-rules table active in this workspace (loaded on every recompute).</summary>
        public SwpIndexRules IndexRules;
        public SwpWorkspaceModel Model;
        public List<SwpLogEntry> BuildLog = new List<SwpLogEntry>();

        /// <summary>Build-time bytes of the SWP — the fingerprint base for apply's TOCTOU check.</summary>
        public byte[] SwpBytes;

        /// <summary>Spine worksheet row (1-based) → model row index; and the inverse.</summary>
        public readonly Dictionary<int, int> ModelRowByWorksheetRow = new Dictionary<int, int>();
        public readonly Dictionary<int, int> WorksheetRowByModelRow = new Dictionary<int, int>();

        /// <summary>The spine's rendered data block, kept to restore stray edits cell-by-cell.</summary>
        public object[,] SpineData;

        /// <summary>What the _Preview sheet currently unfolds; null when none is open.</summary>
        public SwpBridgePreview.Context PreviewContext;

        // ---- glossary (optional) ------------------------------------------------------------

        /// <summary>Baselines parsed from the glossary files at build time (no re-read needed).</summary>
        public List<SwpGlossaryPair> GlossaryStandard = new List<SwpGlossaryPair>();
        public List<SwpGlossaryPair> GlossaryShadow = new List<SwpGlossaryPair>();

        /// <summary>Baselines + tweaks, rebuilt on every recompute; null = no glossary in play.</summary>
        public List<SwpGlossaryPair> GlossaryMerged;

        /// <summary>The session-relevant pairs shown on _Glossary, with their occurrence counts.</summary>
        public List<SwpGlossaryPair> GlossaryRelevant = new List<SwpGlossaryPair>();
        public Dictionary<string, int> GlossaryOccurrences = new Dictionary<string, int>();

        /// <summary>The EN keys _Glossary displayed last render — the scope its parse may rewrite.</summary>
        public HashSet<string> GlossaryDisplayedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public SwpBridgeDashboard Dashboard;

        /// <summary>Set by the dashboard's Cancel button; the builder polls it between phases.</summary>
        public volatile bool CancelRequested;

        private Func<object, bool> _previousSheetExclusion;

        /// <summary>The SWP file the session is working against (root + version folder + name).</summary>
        public string SwpPath =>
            Path.Combine(SwpRootFolder, Settings.CurrentVersionFolder ?? "", Settings.SwpFileName ?? "");

        public string ConfigPath =>
            string.IsNullOrEmpty(Settings.ConfigFileName)
                ? ""
                : Path.Combine(SwpRootFolder, Settings.CurrentVersionFolder ?? "", Settings.ConfigFileName);

        public Dictionary<string, SwpRowOverride> OverridesByMnemonic()
        {
            var map = new Dictionary<string, SwpRowOverride>(StringComparer.Ordinal);
            foreach (SwpRowOverride o in Sidecar.Overrides)
                if (o.Mnemonic.Length > 0 && !map.ContainsKey(o.Mnemonic)) map[o.Mnemonic] = o;
            return map;
        }

        public SwpRowOverride OverrideFor(string mnemonic, bool create)
        {
            foreach (SwpRowOverride o in Sidecar.Overrides)
                if (o.Mnemonic == mnemonic) return o;
            if (!create) return null;
            var fresh = new SwpRowOverride { Mnemonic = mnemonic };
            Sidecar.Overrides.Add(fresh);
            return fresh;
        }

        public void SaveSidecarNow()
        {
            try
            {
                Sidecar.Overrides.RemoveAll(o => o.IsEmpty);
                SwpBridgeSidecarStore.SaveFile(EffectiveSidecarPath, Sidecar);
            }
            catch (Exception ex)
            {
                // Persistence must never crash a gesture; surface on the status bar instead.
                try { ExcelApp.App.StatusBar = "SWP Bridge: sidecar save failed — " + ex.Message; } catch { }
            }
        }

        /// <summary>
        /// Installs the manual-edit-tracker exclusion by WRAPPING whatever is already there (the
        /// IO Bridge may hold the slot), so both bridges stay excluded while either lives.
        /// </summary>
        public void InstallSheetExclusion()
        {
            _previousSheetExclusion = ManualEditTracker.SheetExclusion;
            Func<object, bool> previous = _previousSheetExclusion;
            ManualEditTracker.SheetExclusion = sheet =>
                SwpBridgeWorkspace.IsBridgeSheet(sheet) || (previous?.Invoke(sheet) ?? false);
        }

        public static void Teardown()
        {
            SwpBridgeSession session = Current;
            Current = null;
            if (session == null) return;
            SwpBridgeWorkspace.ForgetWorkspace(session.WorkspaceId);
            ManualEditTracker.SheetExclusion = session._previousSheetExclusion;
            try { session.Dashboard?.CloseFromExcel(); } catch { }
            session.Dashboard = null;
        }

        /// <summary>
        /// Retires the session's runtime state (dashboard, the chained sheet-exclusion) but KEEPS
        /// the workbook and its workspace id live — a reconfigure via Setup rebuilds into the same
        /// workbook, so it is never closed. Closing the bridge's only workbook is what made Excel
        /// quit on a settings change.
        /// </summary>
        public void RetireForRebuild()
        {
            ManualEditTracker.SheetExclusion = _previousSheetExclusion;
            try { Dashboard?.CloseFromExcel(); } catch { }
            Dashboard = null;
            if (ReferenceEquals(Current, this)) Current = null;
        }
    }
}
