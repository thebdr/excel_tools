using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.Commands
{
    /// <summary>Bridges <see cref="CsvRecord"/>s and an Excel worksheet (COM). Never lets Excel coerce values.</summary>
    internal static class CsvSheet
    {
        private const string RecordMarkerPrefix = "//PT.Record:";
        internal const string DamagedBoundaryRecovery =
            "The hidden CSV workspace boundary is missing or damaged. Copy any worksheet edits you need, " +
            "delete this source sheet, then reopen the file. Sheet and helper changes will not be retained.";

        /// <summary>
        /// Writes records to <paramref name="sheet"/> starting at A1 — every row (comment rows included) is
        /// split across columns. The ENTIRE sheet is formatted as Text first, so neither the written values
        /// nor anything the user later types/pastes is coerced (leading zeros, long IDs, dates, formulas all
        /// stay literal). <paramref name="dataWidth"/> is the widest record — remembered so Save re-emits the
        /// same column count.
        /// </summary>
        public static void WriteRecords(
            dynamic sheet,
            IReadOnlyList<CsvRecord> records,
            out int rowCount,
            out int dataWidth,
            out string recordMarkerToken,
            string preservedRecordMarkerToken = null)
        {
            rowCount = records.Count;
            dataWidth = 1;
            recordMarkerToken = null;
            foreach (var r in records)
                if (r.FieldCount > dataWidth) dataWidth = r.FieldCount;

            sheet.Cells.NumberFormat = "@";   // Text — the whole sheet, before any value lands
            if (rowCount > 0)
            {
                var data = new object[rowCount, dataWidth];
                for (int r = 0; r < rowCount; r++)
                {
                    CsvRecord rec = records[r];
                    for (int c = 0; c < dataWidth; c++)
                        data[r, c] = c < rec.FieldCount ? (object)rec.Fields[c] : null;
                }

                dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(rowCount, dataWidth));
                range.Value = data;
            }

            // Excel otherwise collapses blank/trailing-blank CSV records into unused rows. A hidden token
            // column gives the workspace an independent logical row boundary. The token is session-owned,
            // so moving/inserting/clearing row 1 cannot make this metadata look like source data.
            int markerColumn = dataWidth + 2; // one blank spacer keeps CurrentRegion/data tools away
            if (preservedRecordMarkerToken != null && !IsRecordMarkerToken(preservedRecordMarkerToken))
                throw new InvalidOperationException("The existing CSV workspace identity is invalid.");
            recordMarkerToken = preservedRecordMarkerToken ??
                RecordMarkerPrefix + Guid.NewGuid().ToString("N");
            // One end sentinel is enough: inserting/deleting a worksheet row moves it automatically, so its
            // row minus one remains the logical record count (including blank and trailing-blank records).
            sheet.Cells(rowCount + 1, markerColumn).Value2 = recordMarkerToken;
            try
            {
                sheet.Columns(markerColumn - 1).Hidden = true;
                sheet.Columns(markerColumn).Hidden = true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not reserve the hidden CSV workspace boundary.", ex);
            }
            if (!IsReservedColumnPair(sheet, markerColumn))
                throw new InvalidOperationException("Excel did not preserve the hidden CSV workspace boundary.");
        }

        /// <summary>Logical record/data-column dimensions, excluding workspace and view-only helper columns.</summary>
        public static void Measure(dynamic sheet, out int rows, out int dataColumns)
        {
            rows = 0;
            dataColumns = 0;

            if (TryGetWorkspaceLayout(
                (object)sheet, out int ignoredMarker, out dataColumns, out rows, out string ignoredToken)) return;

            long lastRow = RangeHelpers.LastRow(sheet);
            long lastCol = RangeHelpers.LastColumn(sheet);
            if (lastRow < 1 || lastCol < 1) return;

            object first = null;
            try { first = sheet.Cells(1, 1).Value2; } catch { }
            if (lastRow == 1 && lastCol == 1 && first == null) return;

            rows = (int)lastRow;
            HashSet<int> helpers = HelperColumns(sheet, (int)lastCol);
            for (int c = 1; c <= lastCol; c++)
                if (!helpers.Contains(c)) dataColumns++;
            if (dataColumns == 0) dataColumns = 1;
        }

        /// <summary>
        /// Sheet columns excluded from source data. In a workspace, the hidden spacer/marker establish a
        /// hard boundary: columns to their right are view-only helpers. On an ordinary exported sheet, the
        /// legacy row-1 "//" convention is retained.
        /// </summary>
        internal static HashSet<int> HelperColumns(dynamic sheet, int lastColumn)
        {
            var helpers = new HashSet<int>();
            if (lastColumn < 1) return helpers;

            int markerColumn = FindRecordMarkerColumn(sheet, lastColumn);
            if (markerColumn > 1)
            {
                for (int c = markerColumn - 1; c <= lastColumn; c++) helpers.Add(c);
                return helpers;
            }

            object raw = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, lastColumn)).Value2;
            if (raw is object[,] grid)
            {
                int rb = grid.GetLowerBound(0), cb = grid.GetLowerBound(1);
                for (int c = 0; c < grid.GetLength(1); c++)
                {
                    if (!(grid[rb, cb + c] is string marker) ||
                        !marker.StartsWith("//", StringComparison.Ordinal)) continue;
                    helpers.Add(c + 1);
                }
            }
            else if (raw is string marker && marker.StartsWith("//", StringComparison.Ordinal))
            {
                helpers.Add(1);
            }
            return helpers;
        }

        /// <summary>True when a workspace has user/view content to the right of its internal boundary.</summary>
        internal static bool HasViewOnlyContent(dynamic sheet)
        {
            try
            {
                if (!TryGetWorkspaceLayout(
                    (object)sheet, out int markerColumn, out int ignoredColumns,
                    out int ignoredRows, out string ignoredToken)) return false;
                if (markerColumn >= 16384) return false;

                // Excel persists Find's LookIn/LookAt settings globally. Search formulas explicitly so a
                // helper formula whose current value is "" is still protected from Reload's sheet rebuild.
                dynamic viewCells = sheet.Range(
                    sheet.Cells(1, markerColumn + 1), sheet.Cells(1048576, 16384));
                dynamic found = viewCells.Find(
                    "*", Type.Missing, Xl.xlFormulas, Xl.xlPart, Xl.xlByRows, Xl.xlNext);
                if (found != null) return true;
                try { if (Convert.ToInt32(viewCells.Comments.Count) > 0) return true; } catch { }
                try { if (Convert.ToInt32(viewCells.CommentsThreaded.Count) > 0) return true; } catch { }
                return false;
            }
            catch { return true; } // uncertainty must prompt before Reload clears the sheet
        }

        /// <summary>
        /// Detects generated marker metadata even when its hidden boundary was partly damaged. The runtime
        /// session is the primary identity; this fallback is for copied/reopened sheets with no live session.
        /// </summary>
        internal static bool HasWorkspaceMetadataCandidate(dynamic sheet)
        {
            try
            {
                bool persistedWorkspace = !string.IsNullOrEmpty(CsvWorkspace.TryGetWorkspaceId(sheet.Parent));
                dynamic cells = sheet.Cells;
                dynamic found = cells.Find(
                    RecordMarkerPrefix + "*", Type.Missing, Xl.xlFormulas, Xl.xlWhole,
                    Xl.xlByRows, Xl.xlNext, false, false, false);
                if (found == null) return false;
                string firstAddress = (string)found.Address;

                for (int guard = 0; guard < 16384 && found != null; guard++)
                {
                    string marker = found.Value2 as string;
                    int column = (int)found.Column;
                    if (IsRecordMarkerToken(marker))
                    {
                        bool hasHiddenStructure = false;
                        try { hasHiddenStructure = Convert.ToBoolean(sheet.Columns(column).Hidden); } catch { }
                        if (!hasHiddenStructure && column > 1)
                            try { hasHiddenStructure = Convert.ToBoolean(sheet.Columns(column - 1).Hidden); } catch { }
                        if (persistedWorkspace || hasHiddenStructure) return true;
                    }

                    found = cells.Find(
                        RecordMarkerPrefix + "*", found, Xl.xlFormulas, Xl.xlWhole,
                        Xl.xlByRows, Xl.xlNext, false, false, false);
                    if (found == null || string.Equals((string)found.Address, firstAddress, StringComparison.Ordinal))
                        break;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Reads the one structurally valid internal marker without consulting <see cref="CsvSession"/>.
        /// This is the non-recursive source of truth used while resolving a session from a worksheet.
        /// Multiple valid markers are treated as damaged/ambiguous metadata.
        /// </summary>
        internal static bool TryReadSingleInternalRecordMarker(
            dynamic sheet, out string markerToken, out int markerColumn)
        {
            markerToken = null;
            markerColumn = -1;
            if (sheet == null) return false;

            try
            {
                // Scan columns rather than Find-ing across the sheet: the sentinel's row may be hidden by
                // a filter, and Find cannot see into a hidden row.
                var markers = ScanReservedMarkers(sheet, RecordMarkerPrefix + "*");
                if (markers.Count != 1) return false;   // 0 = none, >1 = ambiguous/damaged metadata

                markerColumn = markers[0].Key;
                markerToken = markers[0].Value;
                return true;
            }
            catch
            {
                markerToken = null;
                markerColumn = -1;
                return false;
            }
        }

        internal static bool TryValidateRecordMarker(
            dynamic sheet, string expectedToken, out int markerColumn)
        {
            markerColumn = -1;
            string actualToken = null;
            int actualColumn = -1;
            if (!IsRecordMarkerToken(expectedToken) ||
                !TryReadSingleInternalRecordMarker(sheet, out actualToken, out actualColumn) ||
                !string.Equals(actualToken, expectedToken, StringComparison.Ordinal)) return false;
            markerColumn = actualColumn;
            return true;
        }

        internal static string TryGetRecordMarkerId(string markerToken)
        {
            if (string.IsNullOrEmpty(markerToken) ||
                !markerToken.StartsWith(RecordMarkerPrefix, StringComparison.Ordinal)) return null;
            string id = markerToken.Substring(RecordMarkerPrefix.Length);
            return Guid.TryParseExact(id, "N", out _) ? id.ToLowerInvariant() : null;
        }

        /// <summary>Absolute worksheet column of this add-in's hidden logical-record marker, or -1.</summary>
        internal static int FindRecordMarkerColumn(dynamic sheet, int lastColumn)
        {
            CsvSession session = CsvSession.Get(sheet);
            if (!string.IsNullOrEmpty(session?.RecordMarkerToken))
            {
                if (TryValidateRecordMarker(sheet, session.RecordMarkerToken, out int column))
                {
                    session.RecordMarkerColumn = column;
                    return column;
                }

                // With a single end sentinel, its absence makes the logical row count unknowable. Never
                // guess from a remembered column, or trailing blank records could be silently dropped.
                return -1;
            }

            // Compatibility fallback for a copied sheet or reopened workbook without a runtime session:
            // the end sentinel is below the data, so search the sheet rather than assuming it is in row 1.
            return TryReadSingleInternalRecordMarker(sheet, out string ignoredToken, out int markerColumn)
                ? markerColumn : -1;
        }

        /// <summary>
        /// Finds the source boundary and logical record count for a live workspace sheet. The last logical
        /// row is the greater of the last source value and last session marker; helper-only rows do not count.
        /// </summary>
        internal static bool TryGetWorkspaceLayout(
            dynamic sheet,
            out int markerColumn,
            out int dataColumns,
            out int logicalRows,
            out string markerToken)
        {
            return TryGetWorkspaceLayout(
                sheet, out markerColumn, out dataColumns, out logicalRows,
                out markerToken, out int ignoredSentinelRow);
        }

        /// <summary>The workspace layout plus the current physical row of its unique end sentinel.</summary>
        internal static bool TryGetWorkspaceLayout(
            dynamic sheet,
            out int markerColumn,
            out int dataColumns,
            out int logicalRows,
            out string markerToken,
            out int sentinelRow)
        {
            markerColumn = -1;
            dataColumns = 0;
            logicalRows = 0;
            markerToken = null;
            sentinelRow = 0;

            CsvSession session = CsvSession.Get(sheet);
            // UsedRange, not RangeHelpers.LastColumn: the latter is Find-based, so a filtered-away
            // sentinel row would hide the reserved pair from the column bound entirely.
            markerColumn = FindRecordMarkerColumn(sheet, UsedLastColumn(sheet));
            if (markerColumn < 3) return false;

            markerToken = session?.RecordMarkerToken;
            if (string.IsNullOrEmpty(markerToken))
            {
                try
                {
                    int tokenRow = MatchRowInColumn(sheet, markerColumn, RecordMarkerPrefix + "*");
                    if (tokenRow <= 0) return false;
                    markerToken = sheet.Cells(tokenRow, markerColumn).Value2 as string;
                }
                catch { return false; }
            }
            if (!IsRecordMarkerToken(markerToken)) return false;

            dataColumns = markerColumn - 2;
            // Find-based and so filter-sensitive, but harmless: the sentinel below is authoritative and
            // logicalRows takes the greater of the two.
            int lastDataRow = LastValueRow(
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(1048576, dataColumns)));
            if (!TryFindSingleExactRowInColumn(sheet, markerColumn, markerToken, out sentinelRow))
                return false;
            logicalRows = Math.Max(lastDataRow, Math.Max(0, sentinelRow - 1));
            return true;
        }

        /// <summary>Places the movable end sentinel immediately after the last logical record.</summary>
        internal static void SyncRecordMarkers(
            dynamic sheet,
            CsvSession session,
            int recordCount,
            bool preserveManualEditTrackingShadow = false)
        {
            if (session == null || string.IsNullOrEmpty(session.RecordMarkerToken)) return;
            int markerColumn = FindRecordMarkerColumn(sheet, (int)RangeHelpers.LastColumn(sheet));
            if (markerColumn < 3)
                throw new InvalidOperationException(DamagedBoundaryRecovery);
            if (!TryFindSingleExactRowInColumn(
                sheet, markerColumn, session.RecordMarkerToken, out int oldSentinelRow))
                throw new InvalidOperationException(DamagedBoundaryRecovery);

            int newSentinelRow = recordCount + 1;
            if (oldSentinelRow == newSentinelRow)
            {
                session.RecordMarkerColumn = markerColumn;
                session.DataWidth = markerColumn - 2;
                session.RowCount = recordCount;
                session.ColumnCount = markerColumn - 2;
                return;
            }

            using (new ExcelStateGuard(
                suppressEvents: true,
                suppressManualEditTracking: !preserveManualEditTrackingShadow))
            {
                sheet.Cells(oldSentinelRow, markerColumn).ClearContents();
                sheet.Cells(newSentinelRow, markerColumn).Value2 = session.RecordMarkerToken;
                sheet.Columns(markerColumn - 1).Hidden = true;
                sheet.Columns(markerColumn).Hidden = true;
            }

            session.RecordMarkerColumn = markerColumn;
            session.DataWidth = markerColumn - 2;
            session.RowCount = recordCount;
            session.ColumnCount = markerColumn - 2;
        }

        private static int LastValueRow(dynamic range)
        {
            try
            {
                dynamic found = range.Find(
                    "*", Type.Missing, Xl.xlFormulas, Xl.xlPart, Xl.xlByRows, Xl.xlPrevious);
                return found == null ? 0 : (int)found.Row;
            }
            catch { return 0; }
        }

        // ---- visibility-proof column probes -----------------------------------------------------------
        //
        // Range.Find SKIPS rows hidden by an AutoFilter. xlFormulas buys visibility into hidden COLUMNS
        // (which the reserved pair needs) but NOT into hidden ROWS, and the end sentinel sits one row past
        // the last record — empty in every data column, so any filter hides it. Locating the boundary with
        // Find therefore made a filtered sheet look like it had no marker at all: the session was dropped
        // from SheetMap, the Index reported "(sheet deleted)" and saving became impossible until the filter
        // was cleared. WorksheetFunction ignores visibility entirely, so every boundary lookup uses it.

        /// <summary>Non-empty cell count of one whole column, hidden rows included. -1 on failure.</summary>
        private static int CountNonEmptyInColumn(dynamic sheet, int column)
        {
            try { return Convert.ToInt32(ExcelApp.App.WorksheetFunction.CountA(sheet.Columns(column))); }
            catch { return -1; }
        }

        /// <summary>
        /// Row of the first cell in <paramref name="column"/> matching an Excel wildcard pattern, hidden
        /// rows included; 0 when there is no match. MATCH raises on no-match, hence the catch.
        /// </summary>
        private static int MatchRowInColumn(dynamic sheet, int column, string pattern)
        {
            try { return Convert.ToInt32(ExcelApp.App.WorksheetFunction.Match(pattern, sheet.Columns(column), 0)); }
            catch { return 0; }
        }

        /// <summary>
        /// Rightmost column of the used range. Unlike RangeHelpers.LastColumn (Find-based) this still
        /// reports the reserved pair when the sentinel's row is filtered away. It may overstate when stray
        /// formatting exists, which is harmless for a search bound.
        /// </summary>
        private static int UsedLastColumn(dynamic sheet)
        {
            try
            {
                dynamic used = sheet.UsedRange;
                return (int)used.Column + (int)used.Columns.Count - 1;
            }
            catch { return 0; }
        }

        /// <summary>
        /// Row of the one cell in <paramref name="column"/> holding exactly <paramref name="value"/>,
        /// hidden rows included. False when absent or present more than once — a duplicated sentinel is
        /// ambiguous metadata. Other content in the column is tolerated, matching the Find-based rule this
        /// replaced: uniqueness is required of the TOKEN, not of the column.
        /// </summary>
        private static bool TryFindSingleExactRowInColumn(dynamic sheet, int column, string value, out int row)
        {
            row = 0;
            if (string.IsNullOrEmpty(value)) return false;

            string escaped = EscapeWildcards(value);
            int hit = MatchRowInColumn(sheet, column, escaped);
            if (hit <= 0) return false;
            if (CountIfInColumn(sheet, column, escaped) != 1) return false;

            row = hit;
            return true;
        }

        /// <summary>
        /// How many cells of one whole column equal <paramref name="criteria"/>, hidden rows included.
        /// Case-insensitive, like the Find calls this replaced (they passed MatchCase:=false). -1 on failure.
        /// </summary>
        private static int CountIfInColumn(dynamic sheet, int column, string criteria)
        {
            try { return Convert.ToInt32(ExcelApp.App.WorksheetFunction.CountIf(sheet.Columns(column), criteria)); }
            catch { return -1; }
        }

        /// <summary>Neutralizes Excel's MATCH wildcards so a token is compared literally.</summary>
        private static string EscapeWildcards(string value) =>
            value.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?");

        private static bool IsColumnHidden(dynamic sheet, int column)
        {
            try { return Convert.ToBoolean(sheet.Columns(column).Hidden); }
            catch { return false; }
        }

        /// <summary>
        /// Every valid internal marker on the sheet, left to right. Only a hidden column whose left
        /// neighbour is also hidden can be the reserved pair, so the walk reads each column's Hidden state
        /// once and probes just those few candidates. Row visibility never enters into it.
        /// </summary>
        private static List<KeyValuePair<int, string>> ScanReservedMarkers(dynamic sheet, string pattern)
        {
            var found = new List<KeyValuePair<int, string>>();
            int lastColumn = Math.Min(UsedLastColumn(sheet), 16384);
            if (lastColumn < 2) return found;

            bool previousHidden = IsColumnHidden(sheet, 1);
            for (int column = 2; column <= lastColumn; column++)
            {
                bool hidden = IsColumnHidden(sheet, column);
                if (hidden && previousHidden)
                {
                    int row = MatchRowInColumn(sheet, column, pattern);
                    if (row > 0)
                    {
                        string token = sheet.Cells(row, column).Value2 as string;
                        if (IsInternalRecordMarker(sheet, column, token))
                            found.Add(new KeyValuePair<int, string>(column, token));
                    }
                }
                previousHidden = hidden;
            }
            return found;
        }

        private static int FindValidMarkerColumn(dynamic sheet, string what)
        {
            try
            {
                // Column scan for the same reason as TryReadSingleInternalRecordMarker: the sentinel's row
                // can be hidden by a filter, which Find cannot see past. First valid pair wins.
                var markers = ScanReservedMarkers(sheet, what);
                if (markers.Count > 0) return markers[0].Key;
            }
            catch { }
            return -1;
        }

        private static bool IsReservedColumnPair(dynamic sheet, int markerColumn)
        {
            if (markerColumn < 2 || markerColumn > 16384) return false;
            try
            {
                return Convert.ToBoolean(sheet.Columns(markerColumn).Hidden) &&
                    Convert.ToBoolean(sheet.Columns(markerColumn - 1).Hidden);
            }
            catch { return false; }
        }

        private static bool IsInternalBoundary(dynamic sheet, int markerColumn)
        {
            if (!IsReservedColumnPair(sheet, markerColumn)) return false;
            // CountA, not a Find-based scan: content parked in a filtered-away row must still disqualify
            // the spacer, or a damaged boundary could pass this gate on a filtered sheet.
            return CountNonEmptyInColumn(sheet, markerColumn - 1) == 0;
        }

        private static bool IsInternalRecordMarker(dynamic sheet, int markerColumn, string marker)
        {
            if (markerColumn <= 1 || !IsRecordMarkerToken(marker)) return false;

            // The generated marker and its spacer are both hidden. Requiring those structural traits means
            // a legitimate helper/source header such as "//PT.Record" cannot hide the adjacent data column.
            try
            {
                return IsInternalBoundary(sheet, markerColumn);
            }
            catch { return false; }
        }

        private static bool IsRecordMarkerToken(string marker)
            => TryGetRecordMarkerId(marker) != null;

        /// <summary>
        /// Reads a worksheet back into records. Non-string cells are read via their displayed <c>.Text</c> so
        /// error values (#REF!) and formatted numbers survive; doubles are stringified invariant/integer-aware
        /// (no locale separator, no scientific notation). Every row — comment rows included — is emitted at
        /// <paramref name="forcedWidth"/> (or the widest non-empty extent), matching how Excel itself pads
        /// short rows when saving a CSV. A workspace uses an explicit hidden boundary; an ordinary exported
        /// sheet retains the legacy rule that ROW-1 cells starting with "//" mark view-only columns.
        /// </summary>
        public static List<CsvRecord> ReadRecords(dynamic sheet, int forcedWidth)
        {
            var records = new List<CsvRecord>();

            bool workspaceLayout = TryGetWorkspaceLayout(
                (object)sheet,
                out int ignoredMarker,
                out int workspaceDataColumns,
                out int workspaceRows,
                out string ignoredToken);
            CsvSession liveSession = CsvSession.Get((object)sheet);
            if (!workspaceLayout)
            {
                bool hasWorkspaceMetadata = !string.IsNullOrEmpty(liveSession?.RecordMarkerToken) ||
                    HasWorkspaceMetadataCandidate((object)sheet);
                if (hasWorkspaceMetadata)
                    throw new InvalidOperationException(DamagedBoundaryRecovery);
            }
            long lastRow = workspaceLayout ? workspaceRows : RangeHelpers.LastRow(sheet);
            long lastCol = workspaceLayout ? workspaceDataColumns : RangeHelpers.LastColumn(sheet);
            if (lastRow < 1) return records;

            // Find-based helpers report A1 for a genuinely empty (but formatted-as-Text) worksheet.
            // Preserve an empty source as zero records instead of turning it into one blank record on save.
            if (!workspaceLayout && lastRow == 1 && lastCol == 1)
            {
                try { if (sheet.Cells(1, 1).Value2 == null) return records; } catch { }
            }

            dynamic range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
            object raw = range.Value2;

            object[,] grid;
            int rows, cols, rBase, cBase;
            if (raw is object[,] g)
            {
                grid = g;
                rows = g.GetLength(0); cols = g.GetLength(1);
                rBase = g.GetLowerBound(0); cBase = g.GetLowerBound(1);
            }
            else
            {
                grid = new object[1, 1]; grid[0, 0] = raw;
                rows = 1; cols = 1; rBase = 0; cBase = 0;
            }

            // A live workspace range already ends at its explicit source boundary. Ordinary exported sheets
            // retain the legacy row-1 "//" helper convention.
            var kept = new List<int>(cols);
            HashSet<int> helperColumns = workspaceLayout
                ? new HashSet<int>()
                : HelperColumns(sheet, cols);
            for (int c = 0; c < cols; c++)
            {
                if (helperColumns.Contains(c + 1)) continue;
                kept.Add(c);
            }
            if (kept.Count == 0) kept.Add(0);

            // The marker's position is the live source width, so inserted/deleted data columns move the
            // boundary with them. Without a workspace marker, preserve the original minimum-width behavior.
            int width = workspaceLayout
                ? workspaceDataColumns
                : Math.Max(forcedWidth, EffectiveWidth(grid, rows, kept, rBase, cBase));
            if (width < 1) width = 1;

            for (int r = 0; r < rows; r++)
            {
                int sheetRow = r + 1;
                var row = new string[kept.Count];
                for (int j = 0; j < kept.Count; j++)
                    row[j] = CellText(grid[rBase + r, cBase + kept[j]], sheet, sheetRow, kept[j] + 1);

                records.Add(CsvRecord.Data(Resize(row, width)));
            }

            return records;
        }

        private static string[] Resize(string[] row, int width)
        {
            if (row.Length == width) return row;
            var outRow = new string[width];
            for (int c = 0; c < width; c++)
                outRow[c] = c < row.Length ? row[c] : "";
            return outRow;
        }

        private static int EffectiveWidth(object[,] grid, int rows, List<int> kept, int rBase, int cBase)
        {
            int width = 1;
            for (int r = 0; r < rows; r++)
            {
                int lastNonEmpty = 0;
                for (int j = 0; j < kept.Count; j++)
                    if (!IsEmpty(grid[rBase + r, cBase + kept[j]])) lastNonEmpty = j + 1;
                if (lastNonEmpty > width) width = lastNonEmpty;
            }
            return width;
        }

        private static bool IsEmpty(object v) => v == null || (v is string s && s.Length == 0);

        /// <summary>Faithful cell text: exact for strings, displayed .Text for errors/formatted cells, invariant for doubles.</summary>
        internal static string CellText(object v, dynamic sheet, int sheetRow, int sheetCol)
        {
            if (v == null) return "";
            if (v is string s) return s;
            if (v is bool b) return b ? "TRUE" : "FALSE";
            if (v is double d) return FormatDouble(d);

            // CVErr (an Excel error like #REF!) or an unexpected type — use the displayed value.
            try { object t = sheet.Cells(sheetRow, sheetCol).Text; return t == null ? "" : t.ToString(); }
            catch { return Convert.ToString(v, CultureInfo.InvariantCulture); }
        }

        private static string FormatDouble(double d)
        {
            if (!double.IsInfinity(d) && !double.IsNaN(d) && d == Math.Floor(d) && Math.Abs(d) < 9.0e18)
                return ((long)d).ToString(CultureInfo.InvariantCulture);           // integer, no exponent
            return d.ToString("0.###############", CultureInfo.InvariantCulture);   // no locale sep, no scientific
        }
    }
}
