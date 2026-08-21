using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>Per-workspace-sheet render map: worksheet coordinates ↔ model.</summary>
    internal sealed class BridgeSheetMap
    {
        public string SheetName;
        public BridgeSheetModel Model;

        /// <summary>Worksheet row (1-based) → model row. Separator/meta rows are absent.</summary>
        public readonly Dictionary<int, BridgeRow> RowByWorksheetRow = new Dictionary<int, BridgeRow>();

        /// <summary>Per layout column: index into this sheet's own IoHeaders, or -1 (IO side only).</summary>
        public int[] IoColOfLayout = new int[0];

        public int FirstDataRow, LastDataRow;
    }

    /// <summary>
    /// The one live IO Bridge session of this Excel process: document paths, the built model (the
    /// old-value source for edit capture and the three-way check), the queue, the column layout, and
    /// the per-sheet render maps. Rebuilt wholesale by the builder; the queue survives in the sidecar.
    /// </summary>
    internal sealed class IoBridgeSession
    {
        public static IoBridgeSession Current;

        public string WorkspaceId = "";
        public string IoListPath = "", InputsCsvPath = "", OutputsCsvPath = "";
        public List<string> CardSheets = new List<string>();

        public BridgeWorkspaceModel Model;
        public IoBridgeQueue Queue = new IoBridgeQueue();
        public IoBridgeColumnConfig Config = new IoBridgeColumnConfig();
        public ColumnLayout Layout = new ColumnLayout();
        public List<ColumnDef> KnownColumns = new List<ColumnDef>();

        /// <summary>Build-time bytes of both CSVs — the fingerprint base for apply's TOCTOU checks.</summary>
        public byte[] InputsBytes, OutputsBytes;

        /// <summary>How many source IO-list columns each card sheet contributes; 0 = all.</summary>
        public int IoColumnLimit;

        /// <summary>Regex: virtual CSV rows (5/6/9/10) whose mnemonic matches join by name; "" = none.</summary>
        public string VirtualIncludeRegex = "";

        /// <summary>Regex: IO-list rows whose mnemonic matches are shown but ignored by the comparison.</summary>
        public string ExcludeMnemonicRegex = "";

        /// <summary>
        /// The source revision-tracking fills per sheet, kept so un-queueing an action can restore the
        /// cell's ORIGINAL interior instead of blanking it.
        /// </summary>
        public Dictionary<string, Dictionary<(int Row, int Col), int>> SourceFills =
            new Dictionary<string, Dictionary<(int Row, int Col), int>>(StringComparer.OrdinalIgnoreCase);

        public readonly Dictionary<string, BridgeSheetMap> SheetMaps =
            new Dictionary<string, BridgeSheetMap>(StringComparer.Ordinal);

        public string SidecarPath = "";

        /// <summary>What the temporary _Expand sheet currently unfolds; null when none is open.</summary>
        public IoBridgeExpandContext ExpandContext;

        // ---- layout ↔ worksheet column mapping -------------------------------------------------------
        //
        // Column POSITIONS mirror the source document exactly — no reordering, NO spacer columns
        // (user decision): the workspace reads side by side with the original. The mapping is the
        // identity; it stays as an indirection because every renderer already routes through it.
        // Consequence, by Excel's own outline rules: two named groups with no ungrouped column
        // between them MERGE into one +/- region. Leave a column ungrouped to keep groups apart.

        /// <summary>1-based worksheet column per layout column.</summary>
        public int[] SheetColumnOfLayout = new int[0];

        /// <summary>Per worksheet column (index = col − 1): the layout column, or −1 when none.</summary>
        public int[] LayoutOfSheetColumn = new int[0];

        public int TotalSheetColumns;

        public void ComputeSheetColumns()
        {
            var toSheet = new int[Layout.Columns.Count];
            for (int i = 0; i < toSheet.Length; i++) toSheet[i] = i + 1;
            SheetColumnOfLayout = toSheet;
            TotalSheetColumns = toSheet.Length;

            var toLayout = new int[TotalSheetColumns];
            for (int i = 0; i < toLayout.Length; i++) toLayout[i] = i;
            LayoutOfSheetColumn = toLayout;
        }

        /// <summary>Worksheet column (1-based) of a layout column on every bridge card sheet.</summary>
        public int WorksheetColumnOf(int layoutIndex) =>
            layoutIndex >= 0 && layoutIndex < SheetColumnOfLayout.Length
                ? SheetColumnOfLayout[layoutIndex]
                : layoutIndex + 1;

        /// <summary>Layout column behind a worksheet column; −1 for spacers and out-of-range.</summary>
        public int LayoutIndexOfSheetColumn(int worksheetColumn) =>
            worksheetColumn >= 1 && worksheetColumn <= LayoutOfSheetColumn.Length
                ? LayoutOfSheetColumn[worksheetColumn - 1]
                : -1;

        /// <summary>
        /// The value the SOURCE DOCUMENT held at build time for one cell — the old value queued
        /// actions three-way against. Null when the cell is not an editable data cell.
        /// </summary>
        public string ModelValueAt(BridgeSheetMap map, BridgeRow row, int layoutIndex)
        {
            if (row == null || layoutIndex < 0 || layoutIndex >= Layout.Columns.Count) return null;
            LayoutColumn column = Layout.Columns[layoutIndex];

            switch (column.Def.Side)
            {
                case "IO":
                {
                    if (row.Kind == BridgeRowKind.CsvOnly) return null;
                    int ioCol = map.IoColOfLayout[layoutIndex];
                    if (ioCol < 0) return null;
                    return ioCol < row.IoValues.Length && row.IoValues[ioCol] != null ? row.IoValues[ioCol] : "";
                }
                case "CSV":
                {
                    // A multi-record binding shows JOINED values; a typed edit could not name which
                    // record it means, so those cells are view-only.
                    if (row.Csv == null || row.Csv.IsMulti) return null;
                    int canonical = CanonicalIndexOf(column.Def.NormalizedHeader);
                    if (canonical < 0) return null;
                    return canonical < row.CsvValues.Length ? row.CsvValues[canonical] : "";
                }
                default:
                    return null;   // status columns are generated, not editable
            }
        }

        public static int CanonicalIndexOf(string layoutKey)
        {
            for (int i = 0; i < CsvCanonical.Labels.Length; i++)
                if (CsvCanonical.LayoutKey(i) == layoutKey) return i;
            return -1;
        }

        /// <summary>
        /// The FILE's real normalized header behind a CSV-block layout column, for the file the row is
        /// bound to — this is what a queued action's FieldHeader must carry.
        /// </summary>
        public string CsvFieldHeaderFor(BridgeRow row, int canonicalIndex)
        {
            if (row?.Csv == null || canonicalIndex < 0) return "";
            string[] headers = row.Csv.File == BridgeTarget.InputsCsv
                ? Model.InputsFieldHeaders
                : Model.OutputsFieldHeaders;
            return canonicalIndex < headers.Length ? headers[canonicalIndex] : "";
        }

        public void SaveSidecarNow()
        {
            try
            {
                IoBridgeSidecarStore.Save(this);
            }
            catch (Exception ex)
            {
                // Queue persistence must never crash an edit gesture; surface once per failure burst.
                try { ExcelApp.App.StatusBar = "IO Bridge: sidecar save failed — " + ex.Message; } catch { }
            }
        }

        public static void Teardown()
        {
            IoBridgeSession session = Current;
            Current = null;
            if (session != null) IoBridgeWorkspace.ForgetWorkspace(session.WorkspaceId);
            ManualEditTracker.SheetExclusion = null;
            IoBridgeContextMenu.Remove();
        }
    }
}
