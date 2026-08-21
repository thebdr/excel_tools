using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// Workbook identity for the SWP Bridge workspace — the same ephemeral-workbook model as the
    /// IO Bridge: sheets are rebuilt from the documents every session, decisions live in the
    /// sidecar, so identity only answers "is this workbook THE live SWP bridge of this process?".
    /// A saved copy reopened later is re-keyed and never live.
    /// </summary>
    internal static class SwpBridgeWorkspace
    {
        private const string WorkspaceDefinedName = "__PT_SWP_WORKSPACE_ID";

        internal const string HomeSheetName = "Home";
        internal const string SpineSheetName = "Workspace";
        internal const string QueueSheetName = "_Queue";
        internal const string LogSheetName = "_Log";
        internal const string DictionarySheetName = "_Dictionary";
        internal const string GlossarySheetName = "_Glossary";
        internal const string OffsetsSheetName = "_Offsets";
        internal const string TsRefSheetName = "_TSRef";
        internal const string StrobeSheetName = "_StrobeTemplates";
        internal const string PreviewSheetName = "_Preview";

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
                throw new InvalidOperationException("Could not initialize the SWP Bridge workspace identity.", ex);
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

        /// <summary>The live SWP bridge workbook of this process, found by runtime id, never name.</summary>
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
            try { workbook.Windows(1).Caption = "PT_SWPBridge"; } catch { }
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

        /// <summary>Fast gate for the event handlers: is this sheet part of the live SWP bridge?</summary>
        public static bool IsBridgeSheet(object sheetObject)
        {
            SwpBridgeSession session = SwpBridgeSession.Current;
            if (session == null) return false;
            try
            {
                dynamic sheet = sheetObject;
                if (!IsLiveBridgeWorkbook(sheet.Parent)) return false;
                string name = (string)sheet.Name;
                return name == HomeSheetName || name == SpineSheetName || name == QueueSheetName ||
                       name == LogSheetName || name == DictionarySheetName ||
                       name == GlossarySheetName || name == OffsetsSheetName ||
                       name == TsRefSheetName || name == StrobeSheetName ||
                       name == PreviewSheetName;
            }
            catch { return false; }
        }

        private static string ParseDefinedNameValue(string refersTo)
        {
            if (string.IsNullOrEmpty(refersTo) || refersTo.Length < 3) return null;
            if (refersTo[0] != '=' || refersTo[1] != '"' || refersTo[refersTo.Length - 1] != '"') return null;
            return refersTo.Substring(2, refersTo.Length - 3);
        }
    }
}
