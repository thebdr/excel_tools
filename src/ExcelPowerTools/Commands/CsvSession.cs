using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Runtime state for one source-file sheet. Primary identity is the workspace GUID plus a hidden sheet
    /// anchor. A workbook-name hint is retained only as a non-COM fallback for damaged IDs and copy safety.
    /// No COM object is retained by the registry.
    /// </summary>
    internal sealed class CsvSession
    {
        public string WorkspaceId { get; private set; }
        public string SheetCodeName { get; private set; }
        public string RuntimeWorkbookName { get; private set; }
        public string SourcePath { get; set; }
        public CsvFormat Format { get; set; }
        public int HeaderSheetRow { get; set; } = -1;
        public int DataWidth { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public string RecordMarkerToken { get; set; }
        public int RecordMarkerColumn { get; set; }
        public DateTime OpenedAt { get; set; }
        public DateTime SourceModifiedUtc { get; private set; }
        public long SourceLength { get; private set; }
        public CsvSourceFingerprint SourceFingerprint { get; private set; }
        public bool IsDirty { get; set; }
        public string Notes { get; set; } = "";

        private List<CsvRecord> _baselineRecords = new List<CsvRecord>();
        private static readonly List<CsvSession> Sessions = new List<CsvSession>();

        internal static bool HasAnySessions => Sessions.Count != 0;

        public dynamic ResolveSheet()
        {
            dynamic sheet = CsvWorkspace.FindSheet(WorkspaceId, SheetCodeName);
            int markerColumn = -1;
            if (sheet == null || !CsvSheet.TryValidateRecordMarker(
                sheet, RecordMarkerToken, out markerColumn)) return null;
            RecordMarkerColumn = markerColumn;
            RememberRuntimeWorkbookName(sheet.Parent);
            return sheet;
        }

        public void UpdateBaseline(byte[] bytes, IReadOnlyList<CsvRecord> records)
        {
            bytes = bytes ?? Array.Empty<byte>();
            SourceFingerprint = CsvSourceFingerprint.FromBytes(bytes);
            SourceLength = bytes.LongLength;
            try { SourceModifiedUtc = File.GetLastWriteTimeUtc(SourcePath); }
            catch { SourceModifiedUtc = DateTime.MinValue; }

            _baselineRecords = CloneRecords(records);
            IsDirty = false;
        }

        public bool HasLocalChanges(bool force = false)
        {
            return HasLocalChanges(ResolveSheet(), force);
        }

        public bool HasLocalChanges(dynamic sheet, bool force = false)
        {
            if (!force && !IsDirty) return false;
            if (sheet == null) return force;
            try { return HasLocalChanges(CsvSheet.ReadRecords(sheet, DataWidth)); }
            catch { return true; }
        }

        public bool HasLocalChanges(IReadOnlyList<CsvRecord> currentRecords)
        {
            bool changed = CsvDiff.Summarize(_baselineRecords, currentRecords).HasChanges;
            IsDirty = changed;
            return changed;
        }

        /// <summary>Exact content check; timestamps and lengths are display metadata only.</summary>
        public bool HasExternalChange()
        {
            if (string.IsNullOrEmpty(SourcePath) || !File.Exists(SourcePath)) return true;
            try
            {
                var info = new FileInfo(SourcePath);
                byte[] disk = File.ReadAllBytes(SourcePath);
                if (SourceFingerprint != null && SourceFingerprint.Matches(disk))
                {
                    SourceLength = info.Length;
                    SourceModifiedUtc = info.LastWriteTimeUtc;
                    return false;
                }
                return true;
            }
            catch { return true; }
        }

        public string StatusText() => StatusText(ResolveSheet());

        public string StatusText(dynamic sheet)
        {
            bool local = HasLocalChanges(sheet);
            if (string.IsNullOrEmpty(SourcePath) || !File.Exists(SourcePath))
                return local ? "Conflict: edited + missing" : "Missing on disk";

            bool external = HasExternalChange();
            if (external && local) return "Conflict: local + external edits";
            if (external) return "Changed externally";
            return local ? "Edited" : "Unchanged";
        }

        public static void Set(dynamic sheet, CsvSession session)
        {
            if (sheet == null || session == null) return;
            dynamic workbook = sheet.Parent;
            session.WorkspaceId = CsvWorkspace.RequireRuntimeWorkspaceId(workbook);
            session.SheetCodeName = CsvWorkspace.EnsureSheetIdentity(
                sheet, session.RecordMarkerToken, session.RecordMarkerColumn);
            session.RememberRuntimeWorkbookName(workbook);
            Remove(sheet);
            Sessions.Add(session);
        }

        public static CsvSession Get(dynamic sheet)
        {
            if (sheet == null) return null;
            string workspaceId;
            string codeName;
            string markerToken;
            try
            {
                workspaceId = CsvWorkspace.TryGetWorkspaceId(sheet.Parent);
                if (!CsvSheet.TryReadSingleInternalRecordMarker(
                    sheet, out markerToken, out int ignoredColumn)) return null;
                codeName = CsvSheet.TryGetRecordMarkerId(markerToken);
            }
            catch { return null; }
            if (string.IsNullOrEmpty(workspaceId) || string.IsNullOrEmpty(codeName)) return null;

            CsvSession session = Sessions.FirstOrDefault(s =>
                string.Equals(s.WorkspaceId, workspaceId, StringComparison.Ordinal) &&
                string.Equals(s.SheetCodeName, codeName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(s.RecordMarkerToken, markerToken, StringComparison.Ordinal));
            return session != null && CsvWorkspace.SameWorksheet(session.ResolveSheet(), sheet) ? session : null;
        }

        public static CsvSession FindByPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return null;
            foreach (CsvSession session in Sessions
                .Where(s => CsvWorkspaceLogic.SameSourcePath(s.SourcePath, path)).ToList())
            {
                if (session.ResolveSheet() != null) return session;
            }
            return null;
        }

        public static IReadOnlyList<CsvSession> ForWorkbook(dynamic workbook)
        {
            string id = CsvWorkspace.TryGetWorkspaceId(workbook);
            if (string.IsNullOrEmpty(id)) return new List<CsvSession>();
            return Sessions.Where(s => string.Equals(s.WorkspaceId, id, StringComparison.Ordinal)).ToList();
        }

        /// <summary>
        /// Close-time lookup that survives a missing/corrupt workbook ID. The source anchors and hidden
        /// markers independently tie the in-memory sessions back to the workbook being closed.
        /// </summary>
        internal static IReadOnlyList<CsvSession> ForClosingWorkbook(dynamic workbook)
        {
            var result = new List<CsvSession>();
            string workspaceId = CsvWorkspace.TryGetWorkspaceId(workbook);
            if (!string.IsNullOrEmpty(workspaceId))
                result.AddRange(Sessions.Where(s =>
                    string.Equals(s.WorkspaceId, workspaceId, StringComparison.Ordinal)));

            HashSet<string> sheetIds = CsvWorkspace.WorkbookSheetIdentityIds(workbook);
            foreach (CsvSession session in Sessions)
            {
                if (result.Contains(session)) continue;

                bool sheetIdentityMatches = sheetIds.Contains(session.SheetCodeName);
                dynamic knownOwnerWorkbook = CsvWorkspace.FindWorkbook(session.WorkspaceId);
                if (knownOwnerWorkbook == null)
                    knownOwnerWorkbook = CsvWorkspace.FindWorkbookByName(session.RuntimeWorkbookName);
                bool knownOwnerWorkbookIsOpen = knownOwnerWorkbook != null;
                bool knownOwnerWorkbookIsClosing = knownOwnerWorkbookIsOpen &&
                    CsvWorkspace.SameWorkbook(knownOwnerWorkbook, workbook);
                if (CsvWorkspaceLogic.UseSheetIdentityForClosingWorkbook(
                    sheetIdentityMatches, knownOwnerWorkbookIsOpen, knownOwnerWorkbookIsClosing))
                    result.Add(session);
            }
            return result;
        }

        public static dynamic WorkspaceWorkbook()
        {
            foreach (string id in Sessions.Select(s => s.WorkspaceId).Distinct(StringComparer.Ordinal).ToList())
            {
                dynamic wb = CsvWorkspace.FindWorkbook(id);
                if (wb != null) return wb;
            }
            return null;
        }

        public static bool IsWorkspaceWorkbook(dynamic workbook) => ForWorkbook(workbook).Count > 0;

        public static void Remove(dynamic sheet)
        {
            if (sheet == null) return;
            CsvSession session;
            try { session = Get(sheet); }
            catch { session = null; }
            if (session != null)
            {
                Sessions.Remove(session);
                return;
            }

            // Marker damage must not leave a ghost session when Excel is about to delete the real source
            // sheet. Anchor matching is independent of the marker and ignores same-workbook copied sheets,
            // whose duplicate defined names are worksheet-scoped.
            dynamic workbook;
            HashSet<string> anchoredIds;
            try
            {
                workbook = sheet.Parent;
                anchoredIds = CsvWorkspace.SheetIdentityIdsTargeting(sheet);
            }
            catch { return; }
            if (anchoredIds.Count == 0) return;

            string currentWorkspaceId = CsvWorkspace.TryGetWorkspaceId(workbook);
            Sessions.RemoveAll(candidate =>
            {
                if (!anchoredIds.Contains(candidate.SheetCodeName)) return false;
                dynamic registeredWorkbook = CsvWorkspace.FindWorkbook(candidate.WorkspaceId);
                if (registeredWorkbook == null)
                    registeredWorkbook = CsvWorkspace.FindWorkbookByName(candidate.RuntimeWorkbookName);
                if (registeredWorkbook != null)
                    return CsvWorkspace.SameWorkbook(registeredWorkbook, workbook);
                return string.IsNullOrEmpty(currentWorkspaceId) ||
                    string.Equals(currentWorkspaceId, candidate.WorkspaceId, StringComparison.Ordinal);
            });
        }

        private void RememberRuntimeWorkbookName(dynamic workbook)
        {
            try { RuntimeWorkbookName = workbook == null ? null : (string)workbook.Name; }
            catch { }
        }

        public static void RemoveWorkbook(dynamic workbook)
        {
            string id = CsvWorkspace.TryGetWorkspaceId(workbook);
            if (!string.IsNullOrEmpty(id))
                Sessions.RemoveAll(s => string.Equals(s.WorkspaceId, id, StringComparison.Ordinal));
        }

        public static void RemoveWorkspace(string workspaceId)
        {
            if (!string.IsNullOrEmpty(workspaceId))
                Sessions.RemoveAll(s => string.Equals(s.WorkspaceId, workspaceId, StringComparison.Ordinal));
        }

        private static List<CsvRecord> CloneRecords(IReadOnlyList<CsvRecord> records)
        {
            var clone = new List<CsvRecord>();
            if (records == null) return clone;
            foreach (CsvRecord record in records)
            {
                string[] fields = record?.Fields == null ? Array.Empty<string>() : (string[])record.Fields.Clone();
                clone.Add(CsvRecord.Data(fields));
            }
            return clone;
        }

    }
}
