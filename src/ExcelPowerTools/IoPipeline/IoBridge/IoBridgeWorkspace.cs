using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// Workbook identity for the IO Bridge workspace — a deliberate, simpler parallel to
    /// <see cref="Commands.CsvWorkspace"/>. The bridge workbook is ephemeral by design (its sheets are
    /// rebuilt from the source documents every session, the queue lives in the sidecar), so identity
    /// only has to answer one question: is this workbook THE live bridge of this Excel process?
    /// A saved copy reopened later is re-keyed and therefore never live; commands direct the user to
    /// rebuild instead of trusting stale sheets.
    /// </summary>
    internal static class IoBridgeWorkspace
    {
        private const string WorkspaceDefinedName = "__PT_IOB_WORKSPACE_ID";

        internal const string QueueSheetName = "_Queue";
        internal const string ColumnsSheetName = "_Columns";

        private static readonly HashSet<string> RuntimeWorkspaceIds = new HashSet<string>(StringComparer.Ordinal);

        public static string EnsureWorkspaceId(dynamic workbook)
        {
            string existing = TryGetWorkspaceId(workbook);
            if (!string.IsNullOrEmpty(existing) && RuntimeWorkspaceIds.Contains(existing)) return existing;

            string id = Guid.NewGuid().ToString("N");
            try
            {
                try { workbook.Names.Item(WorkspaceDefinedName).Delete(); } catch { }
                dynamic name = workbook.Names.Add(WorkspaceDefinedName, "=\"" + id + "\"");
                try { name.Visible = false; } catch { }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not initialize the IO Bridge workspace identity.", ex);
            }
            RuntimeWorkspaceIds.Add(id);
            return id;
        }

        public static string TryGetWorkspaceId(dynamic workbook)
        {
            if (workbook == null) return null;
            try
            {
                dynamic name = workbook.Names.Item(WorkspaceDefinedName);
                string value = ParseDefinedNameValue((string)name.RefersTo);
                return CsvWorkspaceLogic.IsWorkspaceId(value) ? value : null;
            }
            catch { return null; }
        }

        /// <summary>
        /// A bridge workbook saved to disk carries its hidden ID; on a reopened copy that ID is only
        /// historical and must not alias the live workspace of this process.
        /// </summary>
        public static void RekeyOpenedWorkbook(dynamic workbook)
        {
            if (string.IsNullOrEmpty(TryGetWorkspaceId(workbook))) return;
            bool wasSaved = true;
            try { wasSaved = (bool)workbook.Saved; } catch { }
            try
            {
                try { workbook.Names.Item(WorkspaceDefinedName).Delete(); } catch { }
                dynamic name = workbook.Names.Add(WorkspaceDefinedName, "=\"" + Guid.NewGuid().ToString("N") + "\"");
                try { name.Visible = false; } catch { }
            }
            catch { }
            if (wasSaved) try { workbook.Saved = true; } catch { }
        }

        public static bool IsLiveBridgeWorkbook(dynamic workbook)
        {
            string id = TryGetWorkspaceId(workbook);
            return !string.IsNullOrEmpty(id) && RuntimeWorkspaceIds.Contains(id);
        }

        public static void ForgetWorkspace(string id)
        {
            if (!string.IsNullOrEmpty(id)) RuntimeWorkspaceIds.Remove(id);
        }

        /// <summary>
        /// The live bridge workbook of this process, or null. Resolved by scanning open workbooks for
        /// the runtime id — never by name, which the user can change with Save As.
        /// </summary>
        public static dynamic FindLiveWorkbook(dynamic app)
        {
            try
            {
                foreach (dynamic wb in app.Workbooks)
                    if (IsLiveBridgeWorkbook(wb)) return wb;
            }
            catch { }
            return null;
        }

        public static dynamic CreateWorkbook(dynamic app, out string workspaceId)
        {
            dynamic workbook = app.Workbooks.Add(Xl.xlWBATWorksheet);
            // The window title instead of "Book1" — the caption is free to set on an unsaved
            // workbook (the workbook NAME is not; it would need a save).
            try { workbook.Windows(1).Caption = "PT_IOBridge"; } catch { }
            workspaceId = null;
            try
            {
                workspaceId = EnsureWorkspaceId(workbook);
                return workbook;
            }
            catch
            {
                ForgetWorkspace(workspaceId);
                try { workbook.Close(false); } catch { }
                throw;
            }
        }

        /// <summary>Fast gate for the event handlers: is this sheet part of the live bridge?</summary>
        public static bool IsBridgeSheet(object sheetObject)
        {
            IoBridgeSession session = IoBridgeSession.Current;
            if (session == null) return false;
            try
            {
                dynamic sheet = sheetObject;
                if (!IsLiveBridgeWorkbook(sheet.Parent)) return false;
                string name = (string)sheet.Name;
                return name == QueueSheetName || name == ColumnsSheetName || session.SheetMaps.ContainsKey(name);
            }
            catch { return false; }
        }

        private static string ParseDefinedNameValue(string refersTo)
        {
            // RefersTo of a literal-string name reads ="value".
            if (string.IsNullOrEmpty(refersTo) || refersTo.Length < 3) return null;
            if (refersTo[0] != '=' || refersTo[1] != '"' || refersTo[refersTo.Length - 1] != '"') return null;
            return refersTo.Substring(2, refersTo.Length - 3);
        }
    }
}
