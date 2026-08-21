using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// Headless indexer for SWP diagnostic sheets. Streams each requested worksheet with
    /// OpenXmlReader (a CHUTE sheet alone is ~35 MB of XML — never DOM-load it) and reads CACHED
    /// formula values: identity column A ("DB_CABINETS_DIAG.Alarm1"), C (table entry) and D (bit)
    /// are formulas whose saved results Excel always writes. A formula identity cell without a
    /// cached value means the file was last saved by something other than Excel — that is refused
    /// with a remedy, because nothing else in the file can be trusted either.
    /// </summary>
    public static class SwpSheetReader
    {
        private static readonly Regex TagTypePattern =
            new Regex(@"\.(State|Alarm\d+|Warning\d+)$", RegexOptions.Compiled);

        /// <summary>Thrown when the SWP cannot be indexed at all (foreign save, no sheets…).</summary>
        public sealed class SwpUnreadableException : Exception
        {
            public SwpUnreadableException(string message) : base(message) { }
        }

        public static SwpWorkbookIndex Read(byte[] xlsxBytes, IEnumerable<string> targetSheets,
            List<SwpLogEntry> log, bool includeAllDiagnosticSheets = false)
        {
            var index = new SwpWorkbookIndex
            {
                Sheets = new Dictionary<string, SwpSheetIndex>(StringComparer.OrdinalIgnoreCase),
            };
            var wanted = new HashSet<string>(targetSheets ?? new string[0], StringComparer.OrdinalIgnoreCase);

            using (var stream = new MemoryStream(xlsxBytes, writable: false))
            {
                // Not a `using`: the dispose is defensive. Every sheet is parsed from a COPY of
                // its part (below), so by dispose time nothing of ours depends on the package.
                SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
                try
                {
                    WorkbookPart workbook = document.WorkbookPart;
                    if (workbook?.Workbook?.Sheets == null)
                        throw new SwpUnreadableException("The SWP workbook holds no sheets.");

                    IReadOnlyList<string> sharedStrings = LoadSharedStrings(workbook);

                    var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (Sheet sheet in workbook.Workbook.Sheets.Elements<Sheet>())
                    {
                        string name = sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : null;
                        if (string.IsNullOrEmpty(name)) continue;
                        index.AllSheetNames.Add(name);
                        if (!(workbook.GetPartById(sheet.Id.Value) is WorksheetPart part)) continue;

                        bool want = wanted.Contains(name);
                        if (!want && !includeAllDiagnosticSheets) continue;

                        // One deterministic pass over the part's zip stream: copy it out fully
                        // and parse the copy. The packaging layer (WindowsBase on net48) keeps
                        // per-part stream state that a partial read or a mid-stream error can
                        // leave half-dead — its dispose then throws "Cannot access a closed
                        // Stream" long after the real cause (seen in the field 2026-08-21).
                        MemoryStream partXml = CopyPart(part, name, required: want, log);
                        if (partXml == null) continue;
                        using (partXml)
                        {
                            // Each reader gets its own zero-copy view: the SDK reader may close
                            // the stream it was handed, and the buffer must survive the scan.
                            if (!want) want = IsDiagnosticSheet(part, View(partXml), sharedStrings);
                            if (!want) continue;

                            found.Add(name);
                            index.Sheets[name] = ReadSheet(part, View(partXml), name, sharedStrings, log);
                        }
                    }

                    foreach (string name in wanted)
                        if (!found.Contains(name)) index.MissingSheets.Add(name);
                    index.MissingSheets.Sort(StringComparer.OrdinalIgnoreCase);
                }
                finally
                {
                    DisposeDefensively(document, log);
                }
            }
            return index;
        }

        /// <summary>
        /// The part's raw XML, fully drained from the package in one go. A copy that fails on a
        /// REQUIRED sheet makes the whole read fail with the real cause (a truncated OneDrive
        /// sync, a damaged zip); a scan-only sheet just logs and skips.
        /// </summary>
        private static MemoryStream CopyPart(WorksheetPart part, string name, bool required,
            List<SwpLogEntry> log)
        {
            try
            {
                using (Stream source = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    var copy = new MemoryStream();
                    source.CopyTo(copy);
                    copy.Position = 0;
                    return copy;
                }
            }
            catch (Exception ex)
            {
                if (required)
                    throw new SwpUnreadableException(
                        "Sheet '" + name + "' could not be read from the workbook (" + ex.Message +
                        ") — the file may be mid-sync or damaged. Open it in Excel, save it once, and retry.");
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "sheet '" + name + "' could not be read and was skipped: " + ex.Message));
                return null;
            }
        }

        /// <summary>A fresh read-only stream over the copy's buffer — no allocation, no sharing.</summary>
        private static MemoryStream View(MemoryStream copy) =>
            new MemoryStream(copy.GetBuffer(), 0, (int)copy.Length, writable: false);

        /// <summary>A typed reader over copied part XML — the part supplies only its features.</summary>
        private static OpenXmlReader ReaderOver(WorksheetPart part, MemoryStream xml) =>
            new OpenXmlPartReader(xml, part.Features, new OpenXmlPartReaderOptions());

        /// <summary>
        /// Read-only package teardown. With every part already copied out, a dispose-time
        /// packaging error is pure zip bookkeeping over a broken source — noted, never fatal.
        /// </summary>
        private static void DisposeDefensively(SpreadsheetDocument document, List<SwpLogEntry> log)
        {
            try { document.Dispose(); }
            catch (Exception ex)
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the workbook package closed uncleanly (" + ex.GetType().Name +
                    ") — harmless, every sheet was already copied out"));
            }
        }

        /// <summary>The "# Sheet parameters" marker in A2 — the diagnostic-sheet family badge.</summary>
        private static bool IsDiagnosticSheet(WorksheetPart part, MemoryStream partXml,
            IReadOnlyList<string> sharedStrings)
        {
            try
            {
                using (OpenXmlReader reader = ReaderOver(part, partXml))
                {
                    while (reader.Read())
                    {
                        if (reader.ElementType != typeof(Row) || !reader.IsStartElement) continue;
                        var row = (Row)reader.LoadCurrentElement();
                        if (row.RowIndex == null) continue;
                        int rowNumber = (int)row.RowIndex.Value;
                        if (rowNumber > 2) return false;
                        if (rowNumber != 2) continue;
                        foreach (Cell cell in row.Elements<Cell>())
                            if (ColumnOf(cell.CellReference?.Value) == 1)
                                return Text(cell, sharedStrings).TrimStart().StartsWith("#", StringComparison.Ordinal);
                        return false;
                    }
                }
            }
            catch { }
            return false;
        }

        private static SwpSheetIndex ReadSheet(WorksheetPart part, MemoryStream partXml, string name,
            IReadOnlyList<string> sharedStrings, List<SwpLogEntry> log)
        {
            var sheet = new SwpSheetIndex { SheetName = name };

            string blockFUnit = "", blockLocation = "", blockLabel = "";
            int blockHeaderRow = 0;

            using (OpenXmlReader reader = ReaderOver(part, partXml))
            {
                while (reader.Read())
                {
                    if (reader.ElementType != typeof(Row) || !reader.IsStartElement) continue;
                    var row = (Row)reader.LoadCurrentElement();
                    if (row.RowIndex == null) continue;
                    int rowNumber = (int)row.RowIndex.Value;

                    // Columns A..U (the audited block) + AA; collect them once per row.
                    var byColumn = new Cell[28];
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        int column = ColumnOf(cell.CellReference?.Value);
                        if (column >= 1 && column <= 21) byColumn[column] = cell;
                        else if (column == 27) byColumn[27] = cell;
                    }
                    Cell a = byColumn[1], b = byColumn[2], c = byColumn[3], d = byColumn[4],
                         e = byColumn[5], f = byColumn[6], g = byColumn[7], h = byColumn[8],
                         i = byColumn[9], j = byColumn[10], l = byColumn[12], q = byColumn[17],
                         r = byColumn[18], t = byColumn[20], u = byColumn[21], aa = byColumn[27];

                    if (rowNumber == 2)
                    {
                        sheet.IsDiagnosticSheet = Text(a, sharedStrings).TrimStart()
                            .StartsWith("#", StringComparison.Ordinal);
                        continue;
                    }
                    if (rowNumber == 4) { sheet.DdeTable = Text(b, sharedStrings); continue; }
                    if (rowNumber == 6) { sheet.Description = Text(b, sharedStrings); continue; }

                    // Device-block header: "Table entry:" in B, the entry number in C, the
                    // electrical label / FUnit / Location in F / H / I.
                    string bText = Text(b, sharedStrings);
                    if (bText == "Table entry:")
                    {
                        blockLabel = Text(f, sharedStrings);
                        blockFUnit = Text(h, sharedStrings);
                        blockLocation = Text(i, sharedStrings);
                        blockHeaderRow = rowNumber;
                        sheet.BlockCount++;

                        var header = new SwpBlockHeader
                        {
                            RowNumber = rowNumber,
                            LabelF = blockLabel,
                            FUnitH = blockFUnit,
                            LocationI = blockLocation,
                            LocationIFormula = i?.CellFormula?.Text ?? "",
                        };
                        sheet.BlockHeaders.Add(header);
                        sheet.BlockHeaderByRow[rowNumber] = header;
                        continue;
                    }

                    // A column-A formula with no saved result means the file skipped Excel's
                    // calculation on save — no cached value anywhere can be trusted then.
                    if (a?.CellFormula != null && a.CellValue == null)
                        throw new SwpUnreadableException(
                            "Sheet '" + name + "' row " + rowNumber + ": a formula cell has no cached " +
                            "value — the file was saved by a tool other than Excel. Open it in Excel, " +
                            "save it once, and retry.");

                    string tagType = Text(a, sharedStrings);
                    if (tagType.Length == 0 || !TagTypePattern.IsMatch(tagType)) continue;

                    if (!TryIntCached(c, sharedStrings, out int tableEntry)) continue;
                    if (!TryIntCached(d, sharedStrings, out int bit)) continue;

                    var snapshot = new SwpRowSnapshot
                    {
                        RowNumber = rowNumber,
                        TagType = tagType,
                        TableEntry = tableEntry,
                        Bit = bit,
                        ExportB = bText,
                        PriorityE = Text(e, sharedStrings),
                        PriorityIsText = IsTextTyped(e),
                        AddressG = Text(g, sharedStrings),
                        DeviceJ = Text(j, sharedStrings),
                        TsRefL = Text(l, sharedStrings),
                        Desc1T = Text(t, sharedStrings),
                        Desc2U = Text(u, sharedStrings),
                        CachedQ = Text(q, sharedStrings),
                        CachedR = Text(r, sharedStrings),
                        StampAA = Text(aa, sharedStrings),
                        FormulaQ = q?.CellFormula?.Text ?? "",
                        FormulaR = r?.CellFormula?.Text ?? "",
                        FormulaQShared = IsShared(q),
                        FormulaRShared = IsShared(r),
                        ExportBIsFormula = b?.CellFormula != null,
                        FormulaH = h?.CellFormula?.Text ?? "",
                        FormulaI = i?.CellFormula?.Text ?? "",
                        FormulaHShared = IsShared(h),
                        FormulaIShared = IsShared(i),
                        HasCellH = h != null,
                        HasCellI = i != null,
                        CachedH = Text(h, sharedStrings),
                        CachedI = Text(i, sharedStrings),
                        BlockFUnit = blockFUnit,
                        BlockLocation = blockLocation,
                        BlockLabel = blockLabel,
                        BlockHeaderRow = blockHeaderRow,
                    };

                    // The audit's raw material: which of columns A..U hold formulas, which of
                    // those are shared, and the texts of external-reference formulas.
                    for (int column = 1; column <= 21; column++)
                    {
                        Cell cell = byColumn[column];
                        if (cell?.CellFormula == null) continue;
                        snapshot.FormulaMask |= 1u << (column - 1);
                        if (IsShared(cell)) snapshot.SharedMask |= 1u << (column - 1);
                        string formulaText = cell.CellFormula.Text ?? "";
                        if (formulaText.IndexOf('[') >= 0)
                        {
                            if (snapshot.ExternalFormulas == null)
                                snapshot.ExternalFormulas = new Dictionary<int, string>();
                            snapshot.ExternalFormulas[column] = formulaText;
                        }
                    }

                    sheet.BitRows.Add(snapshot);
                    string key = snapshot.RowKey;
                    if (sheet.RowsByKey.ContainsKey(key))
                    {
                        // The legacy scanner wrote to EVERY matching row; a duplicate triple means
                        // a broken SWP. The first row wins and the anomaly is surfaced.
                        if (!sheet.DuplicateKeys.Contains(key)) sheet.DuplicateKeys.Add(key);
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "SWP sheet '" + name + "' rows " + sheet.RowsByKey[key].RowNumber + " and " +
                            rowNumber + " both carry " + key + " — the first row is used"));
                    }
                    else
                    {
                        sheet.RowsByKey[key] = snapshot;
                    }
                }
            }
            return sheet;
        }

        private static IReadOnlyList<string> LoadSharedStrings(WorkbookPart workbook)
        {
            var strings = new List<string>();
            SharedStringTable table = workbook.SharedStringTablePart?.SharedStringTable;
            if (table == null) return strings;
            foreach (OpenXmlElement item in table.Elements())
                strings.Add(item.InnerText ?? "");
            return strings;
        }

        private static string Text(Cell cell, IReadOnlyList<string> sharedStrings)
        {
            if (cell == null) return "";
            string raw = cell.CellValue?.Text ?? "";
            if (cell.DataType == null) return raw;

            if (cell.DataType.Value == CellValues.SharedString)
            {
                return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx)
                       && idx >= 0 && idx < sharedStrings.Count
                    ? sharedStrings[idx]
                    : "";
            }
            if (cell.DataType.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? cell.InnerText ?? "";
            if (cell.DataType.Value == CellValues.Boolean)
                return raw == "1" ? "TRUE" : "FALSE";
            return raw; // str / error / date — the cached text as written
        }

        private static bool TryIntCached(Cell cell, IReadOnlyList<string> sharedStrings, out int value)
        {
            value = 0;
            string text = Text(cell, sharedStrings).Trim();
            if (text.Length == 0) return false;
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double d)) return false;
            value = (int)Math.Round(d);
            return true;
        }

        /// <summary>A string-typed cell — a "2" stored as text is invisible to the loaders.</summary>
        private static bool IsTextTyped(Cell cell) =>
            cell?.DataType != null &&
            (cell.DataType.Value == CellValues.SharedString ||
             cell.DataType.Value == CellValues.InlineString ||
             cell.DataType.Value == CellValues.String);

        private static bool IsShared(Cell cell) =>
            cell?.CellFormula?.FormulaType != null &&
            cell.CellFormula.FormulaType.Value == CellFormulaValues.Shared;

        /// <summary>1-based column index from a cell reference ("AA17" → 27); 0 when unreadable.</summary>
        private static int ColumnOf(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            int column = 0;
            foreach (char ch in cellReference)
            {
                if (ch >= 'A' && ch <= 'Z') column = column * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') column = column * 26 + (ch - 'a' + 1);
                else break;
            }
            return column;
        }
    }
}
