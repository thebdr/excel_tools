using System;
using System.Collections.Generic;
using System.IO;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using ExcelPowerTools.Core.Files;
using ExcelPowerTools.Interop;
using CoreIoListLogic = ExcelPowerTools.Core.IoPipeline.IoListLogic;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// Builds (and rebuilds) the bridge workspace: reads the selected card sheets and both CSVs, runs
    /// the Core join, re-anchors the persisted queue, and renders every sheet. Rendering is always a
    /// full regenerate — the sidecar, not the workbook, is the durable state.
    /// </summary>
    internal static class IoBridgeBuilder
    {
        private const int MetaRow = 1, GroupRow = 2, HeaderRow = 3, FirstDataRow = 4;

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        internal static readonly int CsvHeaderOrange = Rgb(250, 200, 140);   // the CSV block stands apart
        private static readonly int GroupGrey = Rgb(217, 217, 217);
        private static readonly int PendingYellow = Rgb(255, 235, 156);
        private static readonly int ConflictRed = Rgb(255, 199, 206);
        private static readonly int ExternalBlue = Rgb(221, 235, 247);
        private static readonly int SeparatorGrey = Rgb(191, 191, 191);
        private static readonly int QuietGrey = Rgb(166, 166, 166);      // unused spares fade away
        private static readonly int MatchingGreen = Rgb(0x00, 0xCC, 0x5C);   // clean rows; anomalies keep full contrast

        /// <summary>Builds the whole workspace into <paramref name="workbook"/> for an initialized session.</summary>
        public static void BuildWorkspace(dynamic app, dynamic workbook, IoBridgeSession session)
        {
            // ONE guard over the entire build — reading (which may open the source workbook) included —
            // so the screen never repaints mid-build. This is the VBA ScreenUpdating=False discipline.
            using (new ExcelStateGuard(suppressEvents: true, manualCalculation: true))
            {
                // 1. Source documents → model.
                List<CardSheetGrid> cards = ReadCardSheets(app, session.IoListPath, session.CardSheets);
                if (cards.Count == 0)
                    throw new InvalidOperationException("None of the selected sheets could be read as card sheets.");

                session.InputsBytes = ReadCsv(session.InputsCsvPath);
                session.OutputsBytes = ReadCsv(session.OutputsCsvPath);
                CsvDocument inputs = CsvFile.Load(session.InputsBytes);
                CsvDocument outputs = CsvFile.Load(session.OutputsBytes);

                session.Model = IoBridgeJoin.Build(cards, inputs, outputs, new IoBridgeJoinOptions
                {
                    IncludeVirtualsMatching = session.VirtualIncludeRegex,
                    ExcludeIoMnemonicsMatching = session.ExcludeMnemonicRegex,
                });
                if (session.Model.Sheets.Count == 0)
                    throw new InvalidOperationException("No selected sheet carries a Bit + Mnemonic header.");
                if (session.Model.InputsMap == null)
                    throw new InvalidOperationException(
                        "The Inputs CSV could not be read as a plant CSV: no header names anything " +
                        "like Signal ID/Node/Offset/Bit/Card, and the legacy column positions " +
                        "(2nd, 12th, 13th, 14th, 16th) hold no signal-shaped data either.");
                if (session.Model.OutputsMap == null)
                    throw new InvalidOperationException(
                        "The Outputs CSV could not be read as a plant CSV: no header names anything " +
                        "like Output Name/Node/Offset/Bit/Board number, and the legacy column positions " +
                        "(2nd, 5th, 6th, 7th, 9th) hold no signal-shaped data either.");

                // Interior colors from the saved file bytes — the electrical dept's revision-tracking
                // fills, read headlessly and KEPT on the session so un-queueing can restore them.
                session.SourceFills = ReadSourceFills(session);

                // 2. Column layout from the discovered columns + per-user registry config.
                session.KnownColumns = DiscoverColumns(session.Model, session.IoColumnLimit);
                session.Config = IoBridgeRegistry.LoadColumnConfig();
                session.Layout = IoBridgeColumnLayout.ComputeLayout(session.KnownColumns, session.Config);
                session.ComputeSheetColumns();   // spacer-aware layout → worksheet column mapping

                // 3. Re-anchor the persisted queue against the fresh documents.
                ReanchorQueue(session, inputs, outputs);

                // 4. Render.
                RemoveGeneratedSheets(workbook);
                session.SheetMaps.Clear();
                session.ExpandContext = null;   // the _Expand sheet is gone with the rebuild

                dynamic anchorSheet = workbook.Worksheets(1);   // the single sheet a fresh workbook has
                bool anchorConsumed = false;

                foreach (BridgeSheetModel sheetModel in session.Model.Sheets)
                {
                    dynamic target;
                    if (!anchorConsumed)
                    {
                        target = anchorSheet;
                        anchorConsumed = true;
                    }
                    else
                    {
                        target = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets(workbook.Worksheets.Count));
                    }
                    session.SourceFills.TryGetValue(sheetModel.SheetName, out Dictionary<(int Row, int Col), int> sheetFills);
                    RenderCardSheet(target, sheetModel, session, sheetFills);
                }

                RenderColumnsSheet(workbook, session);
                IoBridgeQueueSheet.Render(workbook, session);

                try { workbook.Worksheets(session.Model.Sheets[0].SheetName).Activate(); } catch { }
            }

            session.SaveSidecarNow();
        }

        /// <summary>
        /// The source fills, read from the file's SAVED bytes (share-tolerant, so an instance open in
        /// Excel is fine — unsaved recolors simply aren't visible until saved). Never fails a build.
        /// </summary>
        private static Dictionary<string, Dictionary<(int Row, int Col), int>> ReadSourceFills(IoBridgeSession session)
        {
            try
            {
                byte[] bytes;
                using (var input = new FileStream(session.IoListPath, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var buffer = new MemoryStream())
                {
                    input.CopyTo(buffer);
                    bytes = buffer.ToArray();
                }
                return IoListFillReader.ReadFills(bytes, session.CardSheets);
            }
            catch
            {
                return new Dictionary<string, Dictionary<(int Row, int Col), int>>();
            }
        }

        // ---- source reading --------------------------------------------------------------------------

        private static byte[] ReadCsv(string path)
        {
            SourceReadState state = AtomicFileStore.ReadSource(path, out byte[] bytes, out string error);
            if (state != SourceReadState.Present)
                throw new InvalidOperationException("Could not read the CSV:\n" + path +
                    (string.IsNullOrEmpty(error) ? "" : "\n" + error));
            return bytes;
        }

        /// <summary>
        /// Reads the selected sheets of the IO-list workbook. An instance already open in THIS Excel is
        /// read live; otherwise the workbook is opened hidden read-only and closed again. Column widths
        /// are captured alongside the values.
        /// </summary>
        private static List<CardSheetGrid> ReadCardSheets(dynamic app, string ioListPath, List<string> sheetNames)
        {
            var wanted = new HashSet<string>(sheetNames, StringComparer.OrdinalIgnoreCase);
            var result = new List<CardSheetGrid>();

            dynamic wb = FindOpenWorkbook(app, ioListPath);
            bool opened = false;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
            try
            {
                if (wb == null)
                {
                    wb = app.Workbooks.Open(ioListPath, 0, true);   // UpdateLinks:=0, ReadOnly:=true
                    opened = true;
                }

                foreach (dynamic ws in wb.Worksheets)
                {
                    string name;
                    try { name = (string)ws.Name; } catch { continue; }
                    if (!wanted.Contains(name)) continue;

                    CardSheetGrid grid = ReadSheetGrid(ws);
                    grid.SheetName = name;
                    CaptureColumnWidths(ws, grid);
                    result.Add(grid);
                }
                return result;
            }
            finally
            {
                if (opened) try { wb?.Close(false); } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
            }
        }

        private static void CaptureColumnWidths(dynamic sheet, CardSheetGrid grid)
        {
            int columns = 0;
            foreach (string[] row in grid.AllRows) columns = Math.Max(columns, row.Length);
            for (int c = 1; c <= columns; c++)
            {
                double width = 0;
                try { width = (double)sheet.Columns(c).ColumnWidth; } catch { }
                grid.ColumnWidths.Add(width);
            }
        }

        internal static dynamic FindOpenWorkbook(dynamic app, string path)
        {
            try
            {
                foreach (dynamic wb in app.Workbooks)
                {
                    string full = null;
                    try { full = (string)wb.FullName; } catch { }
                    if (CsvWorkspaceLogic.SameSourcePath(full, path)) return wb;
                }
            }
            catch { }
            return null;
        }

        private static CardSheetGrid ReadSheetGrid(dynamic sheet)
        {
            var grid = new CardSheetGrid { FirstRowNumber = 1 };

            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);
            if (lastRow < 1) return grid;

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            object raw = range.Value2;

            int rows = (int)lastRow, cols = (int)lastCol;
            if (raw is object[,] block)
            {
                int rb = block.GetLowerBound(0), cb = block.GetLowerBound(1);
                for (int r = 0; r < rows; r++)
                {
                    var line = new string[cols];
                    for (int c = 0; c < cols; c++)
                        line[c] = CsvSheet.CellText(block[rb + r, cb + c], sheet, r + 1, c + 1);
                    grid.AllRows.Add(line);
                }
            }
            else
            {
                grid.AllRows.Add(new string[] { (string)CsvSheet.CellText(raw, sheet, 1, 1) });
            }
            return grid;
        }

        // ---- columns & re-anchoring ------------------------------------------------------------------

        /// <summary>
        /// IO columns (union across sheets, first-seen order/sample), CSV canonical block, status.
        /// <paramref name="ioColumnLimit"/> &gt; 0 takes only that many SOURCE columns per sheet — the
        /// IO lists trail off into rarely-needed columns and the workspace can stop before them.
        /// </summary>
        private static List<ColumnDef> DiscoverColumns(BridgeWorkspaceModel model, int ioColumnLimit)
        {
            var columns = new List<ColumnDef>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            foreach (BridgeSheetModel sheet in model.Sheets)
            {
                int take = ioColumnLimit > 0 ? Math.Min(ioColumnLimit, sheet.IoHeaders.Length) : sheet.IoHeaders.Length;
                for (int c = 0; c < take; c++)
                {
                    string key = CoreIoListLogic.NormalizeHeader(sheet.IoHeaders[c]);
                    if (key.Length == 0 || !seen.Add(key)) continue;
                    columns.Add(new ColumnDef { NormalizedHeader = key, SampleHeader = sheet.IoHeaders[c].Trim(), Side = "IO" });
                }
            }

            for (int c = 0; c < CsvCanonical.Labels.Length; c++)
                columns.Add(new ColumnDef
                {
                    NormalizedHeader = CsvCanonical.LayoutKey(c),
                    SampleHeader = "CSV " + CsvCanonical.Labels[c],
                    Side = "CSV",
                });

            columns.Add(new ColumnDef { NormalizedHeader = "status.match", SampleHeader = "//Match", Side = "Status" });
            columns.Add(new ColumnDef { NormalizedHeader = "status.row", SampleHeader = "//Row", Side = "Status" });
            return columns;
        }

        private static void ReanchorQueue(IoBridgeSession session, CsvDocument inputs, CsvDocument outputs)
        {
            var inputsResolver = new CsvRecordResolver(inputs, session.Model.InputsMap);
            var outputsResolver = new CsvRecordResolver(outputs, session.Model.OutputsMap);

            session.Queue.Reanchor(action =>
            {
                if (action.Target == BridgeTarget.IoList)
                {
                    BridgeRow row = session.Model.FindByIdentity(action.Signal);
                    if (row == null || row.Kind == BridgeRowKind.CsvOnly) return ReanchorProbe.NotFound;

                    int ioCol = IoColumnFor(session, row, action.FieldHeader);
                    if (ioCol < 0) return ReanchorProbe.NotFound;
                    return ReanchorProbe.Of(ioCol < row.IoValues.Length ? row.IoValues[ioCol] ?? "" : "");
                }

                CsvDocument doc = action.Target == BridgeTarget.InputsCsv ? inputs : outputs;
                CsvColumnsMap map = action.Target == BridgeTarget.InputsCsv ? session.Model.InputsMap : session.Model.OutputsMap;
                CsvRecordResolver resolver = action.Target == BridgeTarget.InputsCsv ? inputsResolver : outputsResolver;

                int record = resolver.Resolve(action.Signal, out _);
                switch (action.Kind)
                {
                    case BridgeActionKind.CreateCsvRow:
                        return ReanchorProbe.Of(record >= 0 ? "+" : "");
                    case BridgeActionKind.DisableCsvRow:
                    case BridgeActionKind.EnableCsvRow:
                    {
                        if (record < 0) return ReanchorProbe.NotFound;
                        string[] f = doc.Records[record].Fields;
                        bool disabled = f.Length > 0 && f[0] != null && f[0].StartsWith("#", StringComparison.Ordinal);
                        return ReanchorProbe.Of(disabled ? "1" : "0");
                    }
                    case BridgeActionKind.DeleteCsvRow:
                        return ReanchorProbe.Of(record >= 0 ? "" : "+");   // Old="", New="+": "+" means gone
                    default:
                    {
                        if (record < 0) return ReanchorProbe.NotFound;
                        int field = map.FieldFor(action.FieldHeader);
                        if (field < 0) return ReanchorProbe.NotFound;
                        string[] f = doc.Records[record].Fields;
                        return ReanchorProbe.Of(field < f.Length && f[field] != null ? f[field] : "");
                    }
                }
            });
        }

        /// <summary>The card-sheet column index behind an IO-side FieldHeader, for the row's own sheet.</summary>
        private static int IoColumnFor(IoBridgeSession session, BridgeRow row, string fieldHeader)
        {
            foreach (BridgeSheetModel sheet in session.Model.Sheets)
                if (sheet.SpineRows.Contains(row))
                {
                    for (int c = 0; c < sheet.IoHeaders.Length; c++)
                        if (CoreIoListLogic.NormalizeHeader(sheet.IoHeaders[c]) == fieldHeader) return c;
                    return -1;
                }
            return -1;
        }

        // ---- rendering -------------------------------------------------------------------------------

        private static void RemoveGeneratedSheets(dynamic workbook)
        {
            // A rebuild starts from one clean sheet: keep sheet 1, drop the rest (all generated).
            // Alerts off — Worksheet.Delete otherwise prompts once per non-empty sheet.
            dynamic app = ExcelApp.App;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
            try
            {
                while ((int)workbook.Worksheets.Count > 1)
                    workbook.Worksheets(workbook.Worksheets.Count).Delete();
                dynamic first = workbook.Worksheets(1);
                first.Cells.Clear();
                try { first.Cells.ClearOutline(); } catch { }
                try { first.Name = "_building"; } catch { }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not reset the workspace workbook.", ex);
            }
            finally
            {
                try { app.DisplayAlerts = prevAlerts; } catch { }
            }
        }

        private static void RenderCardSheet(
            dynamic sheet, BridgeSheetModel model, IoBridgeSession session,
            Dictionary<(int Row, int Col), int> sourceFills)
        {
            sheet.Name = UniqueName(sheet.Parent, model.SheetName);

            var map = new BridgeSheetMap { SheetName = (string)sheet.Name, Model = model };
            map.IoColOfLayout = BuildIoColumnMap(session.Layout, model);
            session.SheetMaps[map.SheetName] = map;

            int layoutCount = session.Layout.Columns.Count;
            int sheetColumns = session.TotalSheetColumns;
            int totalRows = model.SpineRows.Count + (model.CsvOnlyRows.Count > 0 ? 1 + model.CsvOnlyRows.Count : 0);

            // Header block (meta + group labels + headers). All blocks span the SHEET columns —
            // spacer positions simply stay empty.
            sheet.Cells(MetaRow, 1).Value =
                "IO Bridge · " + model.SheetName +
                " · IO list " + Path.GetFileName(session.IoListPath) +
                " · Inputs " + Path.GetFileName(session.InputsCsvPath) +
                " · Outputs " + Path.GetFileName(session.OutputsCsvPath) +
                " · built " + DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            sheet.Cells(MetaRow, 1).Font.Bold = true;

            var groupLabels = new object[1, sheetColumns];
            foreach (GroupSpan span in session.Layout.Spans)
                if (span.Name.Length > 0)
                    groupLabels[0, session.WorksheetColumnOf(span.FirstIndex) - 1] = span.Name;
            WriteBlock(sheet, GroupRow, 1, groupLabels);
            dynamic groupRange = sheet.Range(sheet.Cells(GroupRow, 1), sheet.Cells(GroupRow, sheetColumns));
            groupRange.Interior.Color = GroupGrey;
            groupRange.Font.Bold = true;

            var headers = new object[1, sheetColumns];
            for (int c = 0; c < layoutCount; c++)
                headers[0, session.WorksheetColumnOf(c) - 1] = session.Layout.Columns[c].Def.SampleHeader;
            dynamic headerRange = sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(HeaderRow, sheetColumns));
            headerRange.NumberFormat = "@";
            WriteBlock(sheet, HeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            PaintCsvHeaderBlock(sheet, session, HeaderRow);
            headerRange.WrapText = true;
            try { headerRange.Font.Name = "Arial Narrow"; } catch { }

            // Data block.
            if (totalRows > 0)
            {
                var data = new object[totalRows, sheetColumns];
                int r = 0;
                foreach (BridgeRow row in model.SpineRows)
                {
                    FillRow(data, r, row, map, session);
                    map.RowByWorksheetRow[FirstDataRow + r] = row;
                    r++;
                }
                if (model.CsvOnlyRows.Count > 0)
                {
                    data[r, 0] = "— CSV-only signals —";
                    int separatorRow = FirstDataRow + r;
                    r++;
                    foreach (BridgeRow row in model.CsvOnlyRows)
                    {
                        FillRow(data, r, row, map, session);
                        map.RowByWorksheetRow[FirstDataRow + r] = row;
                        r++;
                    }
                    dynamic sep = sheet.Range(sheet.Cells(separatorRow, 1), sheet.Cells(separatorRow, sheetColumns));
                    sep.Interior.Color = SeparatorGrey;
                }

                dynamic dataRange = sheet.Range(
                    sheet.Cells(FirstDataRow, 1), sheet.Cells(FirstDataRow + totalRows - 1, sheetColumns));
                dataRange.NumberFormat = "@";
                WriteBlock(sheet, FirstDataRow, 1, data);
            }
            map.FirstDataRow = FirstDataRow;
            map.LastDataRow = FirstDataRow + Math.Max(0, totalRows - 1);

            // The source's revision-tracking interior colors BEFORE our own layers, so the
            // mirror-keyed conditional formats end up on top of them.
            ApplySourceFills(sheet, model, map, session, sourceFills);

            RenderActionHighlights(sheet, map, session);
            ApplyStatusFormats(sheet, map, session, layoutCount, totalRows);
            OutlineRows(sheet, map, session);

            // Widths BEFORE the outline hides: setting ColumnWidth on a hidden column un-hides it,
            // which is exactly how hidden-by-default groups were silently reappearing.
            ApplyColumnWidthsAndHeaderHeight(sheet, map, session, totalRows);
            ApplyOutlineGroups(sheet, session);

            // Ready-to-use AutoFilter over the whole grid. Safe here: nothing in the bridge locates
            // rows with Range.Find, so filter-hidden rows cannot break capture, paint, or apply.
            try
            {
                if ((bool)sheet.AutoFilterMode) sheet.AutoFilterMode = false;   // drop any stale filter
                if (totalRows > 0)
                    sheet.Range(sheet.Cells(HeaderRow, 1), sheet.Cells(map.LastDataRow, sheetColumns)).AutoFilter();
            }
            catch { }

            try
            {
                dynamic win = ExcelApp.App.ActiveWindow;
                sheet.Activate();
                win = ExcelApp.App.ActiveWindow;
                win.FreezePanes = false;
                win.SplitRow = HeaderRow;
                win.SplitColumn = 0;
                win.FreezePanes = true;
            }
            catch { }
        }

        /// <summary>
        /// Column widths: IO columns take the SOURCE sheet's own widths (electrical's layout, captured
        /// at read time), the rest autofit from the data with the usual floor. The header row is
        /// auto-fitted and then CAPPED — narrow columns would otherwise wrap the long header labels
        /// into towers.
        /// </summary>
        private static void ApplyColumnWidthsAndHeaderHeight(
            dynamic sheet, BridgeSheetMap map, IoBridgeSession session, int totalRows)
        {
            const double minColumnWidth = 8.5;
            const double maxHeaderHeight = 48;   // ~3 lines — labels stay readable, the grid stays visible

            int layoutCount = session.Layout.Columns.Count;
            List<double> sourceWidths = map.Model.SourceColumnWidths;
            try
            {
                if (totalRows > 0)
                    sheet.Range(sheet.Cells(FirstDataRow, 1), sheet.Cells(map.LastDataRow, session.TotalSheetColumns)).Columns.AutoFit();

                for (int c = 0; c < layoutCount; c++)
                {
                    dynamic column = sheet.Columns(session.WorksheetColumnOf(c));

                    int ioCol = map.IoColOfLayout[c];
                    if (ioCol >= 0 && ioCol < sourceWidths.Count && sourceWidths[ioCol] > 0.5)
                    {
                        // A near-zero width means the source column is hidden — don't inherit invisibility.
                        column.ColumnWidth = sourceWidths[ioCol];
                        continue;
                    }

                    if ((double)column.ColumnWidth < minColumnWidth) column.ColumnWidth = minColumnWidth;
                }

                sheet.Rows(HeaderRow).AutoFit();
                if ((double)sheet.Rows(HeaderRow).RowHeight > maxHeaderHeight)
                    sheet.Rows(HeaderRow).RowHeight = maxHeaderHeight;
            }
            catch { }
        }

        /// <summary>
        /// Paints the source's revision-tracking INTERIOR colors onto the IO columns — the one piece
        /// of original formatting that matters here. Colors come pre-extracted from the file bytes;
        /// application groups contiguous same-color cells per column, so the COM cost is one
        /// Interior.Color set per colored run (typically a handful per sheet), not per cell.
        /// </summary>
        private static void ApplySourceFills(
            dynamic sheet, BridgeSheetModel model, BridgeSheetMap map, IoBridgeSession session,
            Dictionary<(int Row, int Col), int> sourceFills)
        {
            if (sourceFills == null || sourceFills.Count == 0 || model.SpineRows.Count == 0) return;

            try
            {
                for (int c = 0; c < session.Layout.Columns.Count; c++)
                {
                    int ioCol = map.IoColOfLayout[c];
                    if (ioCol < 0) continue;

                    int runStart = -1, runColor = 0, runLength = 0;
                    for (int i = 0; i <= model.SpineRows.Count; i++)
                    {
                        int? color = null;
                        if (i < model.SpineRows.Count &&
                            sourceFills.TryGetValue((model.SpineRows[i].IoSourceRow, ioCol + 1), out int rgb))
                            color = rgb;

                        if (runLength > 0 && color.HasValue && color.Value == runColor &&
                            FirstDataRow + i == runStart + runLength)
                        {
                            runLength++;
                            continue;
                        }

                        if (runLength > 0)
                        {
                            int sheetColumn = session.WorksheetColumnOf(c);
                            dynamic range = sheet.Range(
                                sheet.Cells(runStart, sheetColumn), sheet.Cells(runStart + runLength - 1, sheetColumn));
                            range.Interior.Color = ToOleColor(runColor);
                        }

                        if (color.HasValue)
                        {
                            runStart = FirstDataRow + i;
                            runColor = color.Value;
                            runLength = 1;
                        }
                        else
                        {
                            runLength = 0;
                        }
                    }
                }
            }
            catch
            {
                try { ExcelApp.App.StatusBar = "IO Bridge: source fills could not be applied for " + model.SheetName + "."; } catch { }
            }
        }

        /// <summary>0xRRGGBB → the BGR integer Excel's Interior.Color expects.</summary>
        private static int ToOleColor(int rgb) =>
            ((rgb >> 16) & 0xFF) | (rgb & 0x00FF00) | ((rgb & 0xFF) << 16);

        private static void FillRow(object[,] data, int r, BridgeRow row, BridgeSheetMap map, IoBridgeSession session)
        {
            for (int c = 0; c < session.Layout.Columns.Count; c++)
            {
                LayoutColumn column = session.Layout.Columns[c];
                switch (column.Def.Side)
                {
                    case "IO":
                    {
                        int ioCol = map.IoColOfLayout[c];
                        if (row.Kind != BridgeRowKind.CsvOnly && ioCol >= 0 && ioCol < row.IoValues.Length)
                            data[r, session.WorksheetColumnOf(c) - 1] = row.IoValues[ioCol];
                        break;
                    }
                    case "CSV":
                    {
                        int canonical = IoBridgeSession.CanonicalIndexOf(column.Def.NormalizedHeader);
                        if (canonical >= 0 && canonical < row.CsvValues.Length)
                            data[r, session.WorksheetColumnOf(c) - 1] = row.CsvValues[canonical];
                        break;
                    }
                    default:
                        data[r, session.WorksheetColumnOf(c) - 1] =
                            column.Def.NormalizedHeader == "status.match" ? row.MatchVerdict : row.RowLabel;
                        break;
                }
            }
        }

        private static int[] BuildIoColumnMap(ColumnLayout layout, BridgeSheetModel model)
        {
            var result = new int[layout.Columns.Count];
            for (int c = 0; c < result.Length; c++)
            {
                result[c] = -1;
                if (layout.Columns[c].Def.Side != "IO") continue;
                for (int io = 0; io < model.IoHeaders.Length; io++)
                    if (CoreIoListLogic.NormalizeHeader(model.IoHeaders[io]) == layout.Columns[c].Def.NormalizedHeader)
                    {
                        result[c] = io;
                        break;
                    }
            }
            return result;
        }

        /// <summary>
        /// Paints action-state highlights DIRECTLY on the affected cells (pending yellow, stale
        /// conflict red, externally-applied blue) and re-displays a pending action's values, so queued
        /// work survives a rebuild visibly. SetField anchors on its own column; row-level actions
        /// (integrate, disable/enable, delete) anchor on the CSV Name cell, and a pending INTEGRATE
        /// additionally shows the record it will create across the CSV block. Direct painting — not
        /// conditional formatting — so the same helper gives immediate live feedback.
        /// </summary>
        internal static void RenderActionHighlights(dynamic sheet, BridgeSheetMap map, IoBridgeSession session)
        {
            foreach (BridgeAction action in session.Queue.Actions)
            {
                if (!action.IsOpen && action.Status != BridgeActionStatus.AppliedExternally) continue;

                BridgeRow row = session.Model.FindByIdentity(action.Signal);
                if (row == null) continue;

                int worksheetRow = WorksheetRowOf(map, row);
                if (worksheetRow < 0) continue;

                if (action.Kind == BridgeActionKind.SetField)
                {
                    int layoutCol = LayoutColumnFor(session, map, row, action);
                    if (layoutCol < 0) continue;

                    if (action.Status == BridgeActionStatus.Pending)
                        sheet.Cells(worksheetRow, session.WorksheetColumnOf(layoutCol)).Value = action.NewValue;
                    PaintActionCell(sheet, session, map, worksheetRow, layoutCol, action.Status);
                    continue;
                }

                // Row-level CSV actions: the Name cell carries the evidence.
                int nameCol = LayoutIndexOf(session, CsvCanonical.LayoutKey(CsvCanonical.Name));
                if (nameCol < 0) continue;

                if (action.Kind == BridgeActionKind.CreateCsvRow &&
                    action.Status == BridgeActionStatus.Pending && action.RowFields != null)
                    DisplayPendingRecord(sheet, session, worksheetRow, action);

                PaintActionCell(sheet, session, map, worksheetRow, nameCol, action.Status);
            }
        }

        /// <summary>The record a pending integrate will create, shown across the row's CSV block.</summary>
        private static void DisplayPendingRecord(
            dynamic sheet, IoBridgeSession session, int worksheetRow, BridgeAction action)
        {
            CsvColumnsMap map = action.Target == BridgeTarget.InputsCsv
                ? session.Model.InputsMap
                : session.Model.OutputsMap;
            if (map == null) return;

            for (int canonical = 0; canonical < CsvCanonical.Labels.Length; canonical++)
            {
                int layoutCol = LayoutIndexOf(session, CsvCanonical.LayoutKey(canonical));
                if (layoutCol < 0) continue;

                int field = CsvCanonical.FieldIndex(map, canonical);
                if (field < 0 || field >= action.RowFields.Length || action.RowFields[field] == null) continue;

                sheet.Cells(worksheetRow, session.WorksheetColumnOf(layoutCol)).Value = action.RowFields[field];
            }
        }

        /// <summary>
        /// The interior color of one action cell: yellow/red/blue by status, or — for a cleared
        /// action — the cell's ORIGINAL source fill (electrical's revision color) or no fill at all.
        /// Called by the builder and, live, by the edit-capture handler.
        /// </summary>
        internal static void PaintActionCell(
            dynamic sheet, IoBridgeSession session, BridgeSheetMap map,
            int worksheetRow, int layoutCol, BridgeActionStatus? status)
        {
            try
            {
                dynamic interior = sheet.Cells(worksheetRow, session.WorksheetColumnOf(layoutCol)).Interior;
                switch (status)
                {
                    case BridgeActionStatus.Pending: interior.Color = PendingYellow; return;
                    case BridgeActionStatus.Conflict:
                    case BridgeActionStatus.ConflictStale: interior.Color = ConflictRed; return;
                    case BridgeActionStatus.AppliedExternally: interior.Color = ExternalBlue; return;
                }

                // Cleared: back to the source's own fill, or none.
                if (TrySourceFill(session, map, worksheetRow, layoutCol, out int rgb))
                    interior.Color = ToOleColor(rgb);
                else
                    interior.ColorIndex = Xl.xlColorIndexNone;
            }
            catch { }
        }

        private static bool TrySourceFill(
            IoBridgeSession session, BridgeSheetMap map, int worksheetRow, int layoutCol, out int rgb)
        {
            rgb = 0;
            if (!session.SourceFills.TryGetValue(map.Model.SheetName, out Dictionary<(int Row, int Col), int> fills))
                return false;
            if (!map.RowByWorksheetRow.TryGetValue(worksheetRow, out BridgeRow row) || row.IoSourceRow <= 0)
                return false;

            int ioCol = layoutCol >= 0 && layoutCol < map.IoColOfLayout.Length ? map.IoColOfLayout[layoutCol] : -1;
            if (ioCol < 0) return false;

            return fills.TryGetValue((row.IoSourceRow, ioCol + 1), out rgb);
        }

        internal static int WorksheetRowOf(BridgeSheetMap map, BridgeRow row)
        {
            foreach (KeyValuePair<int, BridgeRow> pair in map.RowByWorksheetRow)
                if (ReferenceEquals(pair.Value, row)) return pair.Key;
            return -1;
        }

        /// <summary>The layout column an action's field renders in, for flag placement.</summary>
        private static int LayoutColumnFor(IoBridgeSession session, BridgeSheetMap map, BridgeRow row, BridgeAction action)
        {
            for (int c = 0; c < session.Layout.Columns.Count; c++)
            {
                LayoutColumn column = session.Layout.Columns[c];
                if (action.Target == BridgeTarget.IoList)
                {
                    if (column.Def.Side == "IO" && column.Def.NormalizedHeader == action.FieldHeader) return c;
                }
                else if (column.Def.Side == "CSV")
                {
                    int canonical = IoBridgeSession.CanonicalIndexOf(column.Def.NormalizedHeader);
                    if (canonical >= 0 && session.CsvFieldHeaderFor(row, canonical) == action.FieldHeader) return c;
                }
            }
            return -1;
        }

        /// <summary>
        /// The conditional-format layer over the data block — FONT colors only (cell highlights are
        /// painted directly). Rule 1, higher priority: UNUSED spares (empty descriptions + device)
        /// fade to light grey. Rule 2: rows that match cleanly turn the review green, so the anomalies
        /// keep full-contrast text and stand out.
        ///
        /// Relative row references in OM-added rules resolve against the ACTIVE cell, so the block's
        /// first cell is selected before adding — the classic COM conditional-format gotcha.
        /// </summary>
        private static void ApplyStatusFormats(
            dynamic sheet, BridgeSheetMap map, IoBridgeSession session, int layoutCount, int totalRows)
        {
            if (totalRows <= 0 || layoutCount == 0) return;
            const int XlExpression = 2;

            dynamic block = sheet.Range(
                sheet.Cells(FirstDataRow, 1), sheet.Cells(FirstDataRow + totalRows - 1, session.TotalSheetColumns));
            try { block.FormatConditions.Delete(); } catch { }

            int matchColumn = LayoutIndexOf(session, "status.match");
            int rowColumn = LayoutIndexOf(session, "status.row");
            if (matchColumn < 0 || rowColumn < 0) return;

            try
            {
                sheet.Activate();
                sheet.Cells(FirstDataRow, 1).Select();   // normalize relative-reference resolution
            }
            catch { }

            // Absolute column, relative row: $M4 walks down the block with each cell.
            string matchRef = (string)sheet.Cells(FirstDataRow, session.WorksheetColumnOf(matchColumn)).Address(false, true);
            string rowRef = (string)sheet.Cells(FirstDataRow, session.WorksheetColumnOf(rowColumn)).Address(false, true);

            // Unused spares AND comparison-excluded rows fade the same way.
            dynamic unused = block.FormatConditions.Add(XlExpression, Type.Missing,
                "=OR(RIGHT(" + rowRef + ",6)=\"unused\",RIGHT(" + rowRef + ",8)=\"excluded\")");
            unused.Font.Color = QuietGrey;
            unused.StopIfTrue = false;

            dynamic matching = block.FormatConditions.Add(XlExpression, Type.Missing,
                "=AND(" + matchRef + "=\"\",LEFT(" + rowRef + ",3)<>\"CSV\"," + rowRef + "<>\"\"," +
                "RIGHT(" + rowRef + ",6)<>\"unused\",RIGHT(" + rowRef + ",8)<>\"excluded\")");
            matching.Font.Color = MatchingGreen;
            matching.StopIfTrue = false;
        }

        private static int LayoutIndexOf(IoBridgeSession session, string normalizedKey)
        {
            for (int c = 0; c < session.Layout.Columns.Count; c++)
                if (session.Layout.Columns[c].Def.NormalizedHeader == normalizedKey) return c;
            return -1;
        }

        // ---- outlines: double-click cells (blue) and virtual twins (red) ----------------------------

        private static readonly int DoubleClickBlue = Rgb(0, 112, 192);
        private static readonly int VirtualRed = Rgb(200, 30, 30);
        private const int XlContinuous = 1, XlThin = 2, XlMedium = -4138;

        /// <summary>The IO-side column whose double-click opens the IO list at the row — the user's
        /// "Functional unit" (column O of the FVZ layout), the mnemonic when a list lacks it.</summary>
        internal static int IoOpenLayoutColumn(IoBridgeSession session)
        {
            int c = LayoutIndexOf(session, "functionalunit");
            return c >= 0 ? c : LayoutIndexOf(session, "mnemonic");
        }

        /// <summary>The CSV-side column whose double-click opens the CSV at the record — the user's
        /// "CSV Card" (column AG of the FVZ layout).</summary>
        internal static int CsvOpenLayoutColumn(IoBridgeSession session)
        {
            int c = LayoutIndexOf(session, CsvCanonical.LayoutKey(CsvCanonical.Card));
            return c >= 0 ? c : LayoutIndexOf(session, CsvCanonical.LayoutKey(CsvCanonical.Name));
        }

        /// <summary>
        /// Blue borders mark the double-clickable cells (IO-open column on rows from the IO list,
        /// CSV-open column on rows with a record), painted per contiguous run — one COM call each.
        /// A VIRTUAL TWIN's CSV block gets a red outline: the record carries the mnemonic but
        /// acquires nothing from the field, and the row action offers to change its type.
        /// </summary>
        private static void OutlineRows(dynamic sheet, BridgeSheetMap map, IoBridgeSession session)
        {
            try
            {
                int ioOpen = IoOpenLayoutColumn(session);
                int csvOpen = CsvOpenLayoutColumn(session);
                int csvFirst = -1, csvLast = -1;
                for (int c = 0; c < session.Layout.Columns.Count; c++)
                {
                    if (session.Layout.Columns[c].Def.Side != "CSV") continue;
                    if (csvFirst < 0) csvFirst = c;
                    csvLast = c;
                }

                var ioRuns = new List<(int First, int Last)>();
                var csvRuns = new List<(int First, int Last)>();
                var twins = new List<int>();
                foreach (KeyValuePair<int, BridgeRow> pair in map.RowByWorksheetRow)
                {
                    if (pair.Value.IoSourceRow > 0) Extend(ioRuns, pair.Key);
                    if (pair.Value.Csv != null) Extend(csvRuns, pair.Key);
                    if (pair.Value.Csv != null && pair.Value.Csv.Virtual) twins.Add(pair.Key);
                }

                if (ioOpen >= 0)
                    foreach ((int first, int last) in ioRuns)
                        Box(sheet.Range(sheet.Cells(first, session.WorksheetColumnOf(ioOpen)),
                            sheet.Cells(last, session.WorksheetColumnOf(ioOpen))), DoubleClickBlue, XlThin);
                if (csvOpen >= 0)
                    foreach ((int first, int last) in csvRuns)
                        Box(sheet.Range(sheet.Cells(first, session.WorksheetColumnOf(csvOpen)),
                            sheet.Cells(last, session.WorksheetColumnOf(csvOpen))), DoubleClickBlue, XlThin);

                if (csvFirst >= 0)
                    foreach (int row in twins)
                    {
                        dynamic block = sheet.Range(
                            sheet.Cells(row, session.WorksheetColumnOf(csvFirst)),
                            sheet.Cells(row, session.WorksheetColumnOf(csvLast)));
                        block.BorderAround(XlContinuous, XlMedium, Type.Missing, VirtualRed);
                    }
            }
            catch { }
        }

        /// <summary>Every border of the range (edges and the lines between cells) in one color.</summary>
        private static void Box(dynamic range, int color, int weight)
        {
            dynamic borders = range.Borders;
            borders.LineStyle = XlContinuous;
            borders.Weight = weight;
            borders.Color = color;
        }

        /// <summary>Rows arrive sorted by the dictionary's insertion order (top to bottom); adjacent
        /// rows extend the current run, a gap starts a new one.</summary>
        private static void Extend(List<(int First, int Last)> runs, int row)
        {
            if (runs.Count > 0 && runs[runs.Count - 1].Last == row - 1)
                runs[runs.Count - 1] = (runs[runs.Count - 1].First, row);
            else
                runs.Add((row, row));
        }

        /// <summary>Orange over the CSV columns of a header row — the IO/CSV division reads at a glance.</summary>
        internal static void PaintCsvHeaderBlock(dynamic sheet, IoBridgeSession session, int headerRow)
        {
            try
            {
                int first = 0, last = -1;
                for (int c = 0; c < session.Layout.Columns.Count; c++)
                {
                    if (session.Layout.Columns[c].Def.Side != "CSV")
                    {
                        if (last >= first)
                            Paint(sheet, headerRow, session.WorksheetColumnOf(first), session.WorksheetColumnOf(last));
                        first = c + 1;
                        last = -1;
                        continue;
                    }
                    if (last < first) first = c;
                    last = c;
                }
                if (last >= first)
                    Paint(sheet, headerRow, session.WorksheetColumnOf(first), session.WorksheetColumnOf(last));
            }
            catch { }

            void Paint(dynamic ws, int row, int fromCol, int toCol) =>
                ws.Range(ws.Cells(row, fromCol), ws.Cells(row, toCol)).Interior.Color = CsvHeaderOrange;
        }

        /// <summary>
        /// Real Excel outline groups per configured span; hidden-by-default groups collapse. The
        /// spacer column after each group (see the session's column mapping) is what keeps adjacent
        /// groups from MERGING into one — Excel fuses same-level neighbours — and is where the +/-
        /// button lands with summary-on-right.
        /// </summary>
        private static void ApplyOutlineGroups(dynamic sheet, IoBridgeSession session)
        {
            try
            {
                try { sheet.Cells.ClearOutline(); } catch { }
                sheet.Outline.SummaryColumn = Xl.xlSummaryOnRight;

                foreach (GroupSpan span in session.Layout.Spans)
                {
                    if (span.Name.Length == 0) continue;   // the implicit trailing group stays ungrouped
                    dynamic columns = sheet.Range(
                        sheet.Columns(session.WorksheetColumnOf(span.FirstIndex)),
                        sheet.Columns(session.WorksheetColumnOf(span.LastIndex)));
                    columns.Columns.Group();
                    if (span.Hidden) columns.EntireColumn.Hidden = true;
                }
            }
            catch
            {
                // Outline support is best-effort: without it every column simply stays visible.
                try { ExcelApp.App.StatusBar = "IO Bridge: column grouping unavailable on this sheet."; } catch { }
            }
        }

        private static void RenderColumnsSheet(dynamic workbook, IoBridgeSession session)
        {
            dynamic sheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets(workbook.Worksheets.Count));
            sheet.Name = UniqueName(workbook, IoBridgeWorkspace.ColumnsSheetName);

            List<string[]> rows = IoBridgeColumnLayout.RenderColumnsSheet(session.Config, session.KnownColumns);
            var data = new object[rows.Count, IoBridgeColumnLayout.ColumnsSheetHeaders.Length];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < rows[r].Length && c < data.GetLength(1); c++)
                    data[r, c] = rows[r][c];

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows.Count, data.GetLength(1)));
            range.NumberFormat = "@";
            range.Value = data;

            dynamic header = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, data.GetLength(1)));
            header.Font.Bold = true;
            header.Interior.Color = HeaderBlue;

            // Below the table, first cell EMPTY so the parser ignores it: the grouping semantics.
            int noteRow = rows.Count + 2;
            sheet.Cells(noteRow, 2).Value =
                "Columns keep their SOURCE positions; a group covers each contiguous run of its columns. " +
                "Excel merges two groups that touch — leave a column ungrouped between them for separate " +
                "+/- buttons. Run \"IO Bridge: column config\" to apply.";
            sheet.Cells(noteRow, 2).Font.Italic = true;

            // Used columns only — a whole-sheet Columns.AutoFit walks all 16,384 and costs seconds.
            try
            {
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows.Count, data.GetLength(1))).Columns.AutoFit();
            }
            catch { }
        }

        internal static void WriteBlock(dynamic sheet, int firstRow, int firstCol, object[,] data)
        {
            int rows = data.GetLength(0), cols = data.GetLength(1);
            if (rows == 0 || cols == 0) return;
            dynamic range = sheet.Range(
                sheet.Cells(firstRow, firstCol), sheet.Cells(firstRow + rows - 1, firstCol + cols - 1));
            range.Value = data;
        }

        /// <summary>
        /// Sheet-name sanitize + dedupe. NOT CsvWorkspaceLogic.UniqueSheetName, which treats its input
        /// as a file path — "R1.S1.C1" would lose ".C1" as an extension.
        /// </summary>
        private static string UniqueName(dynamic workbook, string wanted)
        {
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "History" };
            try
            {
                foreach (dynamic ws in workbook.Worksheets) existing.Add((string)ws.Name);
            }
            catch { }

            string stem = wanted ?? "";
            foreach (char c in new[] { '[', ']', ':', '*', '?', '/', '\\' }) stem = stem.Replace(c, '_');
            stem = stem.Trim().Trim('\'');
            if (stem.Length == 0) stem = "Sheet";
            if (stem.Length > 31) stem = stem.Substring(0, 31);

            if (!existing.Contains(stem)) return stem;
            for (int n = 2; ; n++)
            {
                string suffix = " (" + n + ")";
                string candidate = (stem.Length + suffix.Length > 31 ? stem.Substring(0, 31 - suffix.Length) : stem) + suffix;
                if (!existing.Contains(candidate)) return candidate;
            }
        }
    }
}
