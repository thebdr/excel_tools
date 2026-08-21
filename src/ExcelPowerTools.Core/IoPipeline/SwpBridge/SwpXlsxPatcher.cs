using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The headless apply core: patches SWP cells directly in the .xlsx bytes. Each affected
    /// worksheet part is REWRITTEN as a stream (reader → writer → part swap) so a 35 MB CHUTE
    /// sheet never becomes a DOM; untouched sheets are not rewritten at all. Literal writes use
    /// inline strings (no sharedStrings bookkeeping; Excel normalizes on its next save); writing
    /// over a template formula (Export B) drops the formula; Q/R formula edits carry a best-effort
    /// cached value so the file displays correctly even before Excel recalculates. The
    /// calculation chain is deleted and fullCalcOnLoad set, so the first open recalculates
    /// everything. Every action passes a three-way check against a fresh read of the same bytes;
    /// the caller owns TOCTOU against the disk file, verification, and the atomic write.
    /// </summary>
    public static class SwpXlsxPatcher
    {
        public sealed class Outcome
        {
            public SwpAction Action;
            public SwpActionStatus Result = SwpActionStatus.Pending;
            public string Note = "";
        }

        public sealed class PatchResult
        {
            /// <summary>The patched workbook — null when no action survived the three-way check.</summary>
            public byte[] PatchedBytes;
            public List<Outcome> Outcomes = new List<Outcome>();
            public int AppliedCount;
        }

        private sealed class CellMutation
        {
            public int Column;
            public bool IsFormula;
            public string Value = "";
            public string CachedValue;   // null = computed in the per-row second pass (Q/R)
        }

        private static bool IsFormulaField(SwpField field) =>
            field == SwpField.FormulaQ || field == SwpField.FormulaR ||
            field == SwpField.FormulaH || field == SwpField.FormulaI ||
            field == SwpField.HeaderFormulaI;

        /// <summary>
        /// <paramref name="stampText"/> is the AA edit stamp ("YYYY-MM-DD.HH:MM:SS", legacy
        /// format) written on every row that received at least one applied action — the caller
        /// supplies it so this stays clock-free and testable.
        /// </summary>
        public static PatchResult Apply(byte[] xlsxBytes, IReadOnlyList<SwpAction> actions,
            string stampText, List<SwpLogEntry> log)
        {
            if (xlsxBytes == null || xlsxBytes.Length == 0) throw new ArgumentException("No SWP bytes.", nameof(xlsxBytes));
            var result = new PatchResult();
            if (actions == null || actions.Count == 0) return result;

            var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SwpAction action in actions) sheetNames.Add(action.Sheet);

            SwpWorkbookIndex index = SwpSheetReader.Read(xlsxBytes, sheetNames, log);

            // sheet → row number → mutations (built only from actions that pass the three-way check)
            var plan = new Dictionary<string, Dictionary<int, List<CellMutation>>>(StringComparer.OrdinalIgnoreCase);
            var newDescriptions = new Dictionary<string, string>();   // sheet|row|col → freshly written T/U value

            foreach (SwpAction action in actions)
            {
                var outcome = new Outcome { Action = action };
                result.Outcomes.Add(outcome);

                if (!index.Sheets.TryGetValue(action.Sheet, out SwpSheetIndex sheet))
                {
                    outcome.Result = SwpActionStatus.Conflict;
                    outcome.Note = "sheet not found";
                    continue;
                }

                int targetRow;
                string current;
                string plannedCache = null;
                SwpRowSnapshot row = null;

                if (action.CoordinateKey.StartsWith("header>", StringComparison.Ordinal))
                {
                    // A block-header repair: the pseudo-coordinate names the row directly.
                    if (!int.TryParse(action.CoordinateKey.Substring("header>".Length), out int headerRow) ||
                        !sheet.BlockHeaderByRow.TryGetValue(headerRow, out SwpBlockHeader header))
                    {
                        outcome.Result = SwpActionStatus.Conflict;
                        outcome.Note = "block header not found";
                        continue;
                    }
                    targetRow = header.RowNumber;
                    current = action.Field == SwpField.HeaderFormulaI ? header.LocationIFormula : header.LocationI;
                    if (action.Field == SwpField.HeaderFormulaI)
                    {
                        // Best-effort cache: what MID(F,4,12) would show, from the label's cache.
                        string label = header.LabelF ?? "";
                        plannedCache = label.Length > 3
                            ? label.Substring(3, Math.Min(12, label.Length - 3))
                            : "";
                    }
                }
                else
                {
                    if (!sheet.RowsByKey.TryGetValue(action.CoordinateKey, out row))
                    {
                        outcome.Result = SwpActionStatus.Conflict;
                        outcome.Note = "target row not found";
                        continue;
                    }
                    targetRow = row.RowNumber;
                    current = CurrentOf(row, action.Field);
                    if (action.Field == SwpField.FormulaH) plannedCache = row.BlockFUnit;
                    else if (action.Field == SwpField.FormulaI) plannedCache = row.BlockLocation;
                }

                if (SwpFieldCompare.Same(action.Field, current, action.NewValue))
                {
                    outcome.Result = SwpActionStatus.AppliedExternally;
                    outcome.Note = "the file already holds the value";
                    continue;
                }
                if (!SwpFieldCompare.Same(action.Field, current, action.OldValue))
                {
                    outcome.Result = SwpActionStatus.Conflict;
                    outcome.Note = "the file now holds \"" + current + "\"";
                    continue;
                }
                if (row != null &&
                    ((action.Field == SwpField.FormulaQ && row.FormulaQShared) ||
                     (action.Field == SwpField.FormulaR && row.FormulaRShared) ||
                     (action.Field == SwpField.FormulaH && row.FormulaHShared) ||
                     (action.Field == SwpField.FormulaI && row.FormulaIShared)))
                {
                    outcome.Result = SwpActionStatus.Conflict;
                    outcome.Note = "shared formula — " +
                        (action.Field == SwpField.FormulaH || action.Field == SwpField.FormulaI
                            ? "the template is intact, the repair is moot"
                            : "not editable headlessly");
                    continue;
                }

                outcome.Result = SwpActionStatus.Applied;
                result.AppliedCount++;

                if (!plan.TryGetValue(action.Sheet, out Dictionary<int, List<CellMutation>> rows))
                    plan[action.Sheet] = rows = new Dictionary<int, List<CellMutation>>();
                if (!rows.TryGetValue(targetRow, out List<CellMutation> mutations))
                    rows[targetRow] = mutations = new List<CellMutation>();

                int column = ColumnOf(action.Field);
                if (action.Field == SwpField.Desc1T || action.Field == SwpField.Desc2U)
                    newDescriptions[action.Sheet + "|" + targetRow + "|" + column] = action.NewValue;

                mutations.Add(new CellMutation
                {
                    Column = column,
                    IsFormula = IsFormulaField(action.Field),
                    Value = action.NewValue,
                    CachedValue = plannedCache,
                });
            }

            if (result.AppliedCount == 0) return result;

            // Formula caches and edit stamps need the whole per-row picture, so a second pass.
            foreach (var sheetPlan in plan)
            {
                SwpSheetIndex sheet = index.Sheets[sheetPlan.Key];
                var byRow = new Dictionary<int, SwpRowSnapshot>();
                foreach (SwpRowSnapshot row in sheet.BitRows) byRow[row.RowNumber] = row;

                foreach (var rowPlan in sheetPlan.Value)
                {
                    if (byRow.TryGetValue(rowPlan.Key, out SwpRowSnapshot row))
                    {
                        foreach (CellMutation mutation in rowPlan.Value)
                        {
                            // Only Q/R still need a cache — theirs depends on the same batch's T/U.
                            if (!mutation.IsFormula || mutation.CachedValue != null) continue;
                            bool isQ = mutation.Column == ColumnOf(SwpField.FormulaQ);
                            bool isR = mutation.Column == ColumnOf(SwpField.FormulaR);
                            if (!isQ && !isR) { mutation.CachedValue = ""; continue; }

                            string descKey = sheetPlan.Key + "|" + rowPlan.Key + "|" +
                                ColumnOf(isQ ? SwpField.Desc1T : SwpField.Desc2U);
                            string descriptionValue = newDescriptions.TryGetValue(descKey, out string written)
                                ? written
                                : (isQ ? row.Desc1T : row.Desc2U);
                            // Exact for the append form; the wholesale CONCAT recalculates on open.
                            mutation.CachedValue = mutation.Value.StartsWith("CONCAT", StringComparison.Ordinal)
                                ? ""
                                : (isQ ? row.CachedQ : row.CachedR) + " " + descriptionValue;
                        }
                    }

                    rowPlan.Value.Add(new CellMutation
                    {
                        Column = ColumnOf(SwpField.StampAA),
                        Value = stampText ?? "",
                    });
                    rowPlan.Value.Sort((a, b) => a.Column.CompareTo(b.Column));
                }
            }

            result.PatchedBytes = Rewrite(xlsxBytes, plan);
            return result;
        }

        private static byte[] Rewrite(byte[] xlsxBytes,
            Dictionary<string, Dictionary<int, List<CellMutation>>> plan)
        {
            using (var stream = new MemoryStream())
            {
                stream.Write(xlsxBytes, 0, xlsxBytes.Length);
                stream.Position = 0;

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
                {
                    WorkbookPart workbook = document.WorkbookPart;

                    foreach (Sheet sheet in workbook.Workbook.Sheets.Elements<Sheet>().ToList())
                    {
                        string name = sheet.Name != null && sheet.Name.HasValue ? sheet.Name.Value : null;
                        if (name == null || !plan.TryGetValue(name, out Dictionary<int, List<CellMutation>> rows))
                            continue;
                        RewriteSheetPart(workbook, sheet, rows);
                    }

                    // Excel rebuilds the (23.9 MB) calculation chain; forcing a full calculation
                    // on load also recomputes the formulas whose text was just edited.
                    if (workbook.CalculationChainPart != null)
                        workbook.DeletePart(workbook.CalculationChainPart);
                    CalculationProperties calc = workbook.Workbook.CalculationProperties;
                    if (calc == null)
                        workbook.Workbook.CalculationProperties = calc = new CalculationProperties();
                    calc.FullCalculationOnLoad = true;
                    workbook.Workbook.Save();
                }

                return stream.ToArray();
            }
        }

        private static void RewriteSheetPart(WorkbookPart workbook, Sheet sheet,
            Dictionary<int, List<CellMutation>> rows)
        {
            var oldPart = (WorksheetPart)workbook.GetPartById(sheet.Id.Value);
            WorksheetPart newPart = workbook.AddNewPart<WorksheetPart>();

            // The new part must keep every child the old one had (printer settings, drawings…)
            // under the SAME relationship ids — the copied XML references them literally.
            foreach (IdPartPair child in oldPart.Parts.ToList())
                newPart.AddPart(child.OpenXmlPart, child.RelationshipId);
            foreach (ExternalRelationship external in oldPart.ExternalRelationships.ToList())
                newPart.AddExternalRelationship(external.RelationshipType, external.Uri, external.Id);
            foreach (HyperlinkRelationship hyperlink in oldPart.HyperlinkRelationships.ToList())
                newPart.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal, hyperlink.Id);

            var columnStyles = new Dictionary<int, uint>();

            // The source XML is drained out of the package in one deterministic pass and read
            // from the copy — a reader directly over the part leaves per-part zip stream state
            // (WindowsBase on net48) that can detonate at package dispose.
            var oldXml = new MemoryStream();
            using (Stream source = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                source.CopyTo(oldXml);
            oldXml.Position = 0;

            using (OpenXmlReader reader = new OpenXmlPartReader(oldXml, oldPart.Features, new OpenXmlPartReaderOptions()))
            using (OpenXmlWriter writer = OpenXmlWriter.Create(newPart))
            {
                bool inSheetData = false;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Worksheet))
                    {
                        if (reader.IsStartElement)
                            writer.WriteStartElement(new Worksheet(), reader.Attributes, reader.NamespaceDeclarations);
                        else
                            writer.WriteEndElement();
                    }
                    else if (reader.ElementType == typeof(SheetData))
                    {
                        if (reader.IsStartElement)
                        {
                            writer.WriteStartElement(new SheetData(), reader.Attributes);
                            inSheetData = true;
                        }
                        else
                        {
                            writer.WriteEndElement();
                            inSheetData = false;
                        }
                    }
                    else if (inSheetData && reader.ElementType == typeof(Row) && reader.IsStartElement)
                    {
                        var row = (Row)reader.LoadCurrentElement();
                        if (row.RowIndex != null &&
                            rows.TryGetValue((int)row.RowIndex.Value, out List<CellMutation> mutations))
                            MutateRow(row, mutations, columnStyles);
                        writer.WriteElement(row);
                    }
                    else if (reader.IsStartElement)
                    {
                        OpenXmlElement element = reader.LoadCurrentElement();
                        if (element is Columns columns) CaptureColumnStyles(columns, columnStyles);
                        writer.WriteElement(element);
                    }
                }
            }

            string relationshipId = workbook.GetIdOfPart(oldPart);
            workbook.DeletePart(oldPart);
            workbook.ChangeIdOfPart(newPart, relationshipId);
        }

        private static void MutateRow(Row row, List<CellMutation> mutations, Dictionary<int, uint> columnStyles)
        {
            row.Spans = null;   // the hint would lie after inserting cells; Excel recomputes it

            foreach (CellMutation mutation in mutations)
            {
                Cell cell = null, insertBefore = null;
                foreach (Cell candidate in row.Elements<Cell>())
                {
                    int column = ColumnIndexOf(candidate.CellReference?.Value);
                    if (column == mutation.Column) { cell = candidate; break; }
                    if (column > mutation.Column) { insertBefore = candidate; break; }
                }

                if (cell == null)
                {
                    cell = new Cell { CellReference = ColumnLetters(mutation.Column) + row.RowIndex.Value };
                    if (columnStyles.TryGetValue(mutation.Column, out uint style))
                        cell.StyleIndex = style;
                    if (insertBefore != null) row.InsertBefore(cell, insertBefore);
                    else row.AppendChild(cell);
                }

                if (mutation.IsFormula)
                {
                    cell.CellFormula = new CellFormula(mutation.Value);
                    cell.DataType = CellValues.String;
                    cell.InlineString = null;
                    cell.CellValue = new CellValue(mutation.CachedValue ?? "");
                }
                else
                {
                    cell.CellFormula = null;
                    cell.CellValue = null;
                    cell.DataType = CellValues.InlineString;
                    var text = new XlsxText(mutation.Value);
                    if (mutation.Value.Length > 0 &&
                        (char.IsWhiteSpace(mutation.Value[0]) ||
                         char.IsWhiteSpace(mutation.Value[mutation.Value.Length - 1])))
                        text.Space = SpaceProcessingModeValues.Preserve;
                    cell.InlineString = new InlineString(text);
                }
            }
        }

        private static void CaptureColumnStyles(Columns columns, Dictionary<int, uint> styles)
        {
            foreach (Column column in columns.Elements<Column>())
            {
                if (column.Style == null || column.Min == null || column.Max == null) continue;
                for (uint c = column.Min.Value; c <= column.Max.Value && c <= 40; c++)
                    styles[(int)c] = column.Style.Value;
            }
        }

        private static string CurrentOf(SwpRowSnapshot row, SwpField field)
        {
            switch (field)
            {
                case SwpField.ExportB: return row.ExportB;
                case SwpField.AddressG: return row.AddressG;
                case SwpField.DeviceJ: return row.DeviceJ;
                case SwpField.TsRefL: return row.TsRefL;
                case SwpField.Desc1T: return row.Desc1T;
                case SwpField.Desc2U: return row.Desc2U;
                case SwpField.FormulaQ: return row.FormulaQ;
                case SwpField.FormulaR: return row.FormulaR;
                case SwpField.StampAA: return row.StampAA;
                case SwpField.FormulaH: return row.FormulaH;
                case SwpField.FormulaI: return row.FormulaI;
                default: return "";
            }
        }

        internal static int ColumnOf(SwpField field)
        {
            switch (field)
            {
                case SwpField.ExportB: return 2;
                case SwpField.AddressG: return 7;
                case SwpField.DeviceJ: return 10;
                case SwpField.TsRefL: return 12;
                case SwpField.FormulaQ: return 17;
                case SwpField.FormulaR: return 18;
                case SwpField.Desc1T: return 20;
                case SwpField.Desc2U: return 21;
                case SwpField.StampAA: return 27;
                case SwpField.FormulaH: return 8;
                case SwpField.FormulaI: return 9;
                case SwpField.HeaderLocationI: return 9;
                case SwpField.HeaderFormulaI: return 9;
                default: throw new ArgumentOutOfRangeException(nameof(field));
            }
        }

        private static int ColumnIndexOf(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            int column = 0;
            foreach (char ch in cellReference)
            {
                if (ch >= 'A' && ch <= 'Z') column = column * 26 + (ch - 'A' + 1);
                else break;
            }
            return column;
        }

        private static string ColumnLetters(int column)
        {
            string letters = "";
            while (column > 0)
            {
                int remainder = (column - 1) % 26;
                letters = (char)('A' + remainder) + letters;
                column = (column - 1) / 26;
            }
            return letters;
        }
    }
}
