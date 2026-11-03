using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.Comparison.RevisionBridge
{
    public sealed class RevisionComparisonReviewSource
    {
        internal RevisionComparisonReviewSource(int index, string path, string contentSha256, long byteLength)
        {
            Index = index;
            Path = path;
            ContentSha256 = contentSha256;
            ByteLength = byteLength;
        }

        public int Index { get; }
        public string Path { get; }
        public string ContentSha256 { get; }
        public long ByteLength { get; }
    }

    public sealed class RevisionComparisonReviewSelection
    {
        private readonly IReadOnlyList<RevisionComparisonReviewSource> _sources;
        private readonly IReadOnlyList<string> _selectedActionIds;

        internal RevisionComparisonReviewSelection(
            IEnumerable<RevisionComparisonReviewSource> sources,
            RevisionComparisonPlan plan,
            string planFingerprint,
            IEnumerable<string> selectedActionIds)
        {
            _sources = new ReadOnlyCollection<RevisionComparisonReviewSource>(sources.OrderBy(item => item.Index).ToList());
            Plan = plan;
            PlanFingerprint = planFingerprint;
            _selectedActionIds = new ReadOnlyCollection<string>(selectedActionIds.ToList());
        }

        public IReadOnlyList<RevisionComparisonReviewSource> Sources => _sources;
        public RevisionComparisonPlan Plan { get; }
        public string PlanFingerprint { get; }
        public IReadOnlyList<string> SelectedActionIds => _selectedActionIds;
    }

    /// <summary>
    /// Writes and reads the deliberately small, macro-free review workbook used by the keyed revision comparison.
    /// The visible workbook is a selection surface only. Apply always re-reads and re-compares every source before
    /// trusting an action ID from this file.
    /// </summary>
    public static class RevisionComparisonReviewFile
    {
        public const string SummarySheetName = "Summary";
        public const string AddedSheetName = "Added";
        public const string RemovedSheetName = "Removed";
        public const string ChangedSheetName = "Changed";
        public const string DuplicateKeysSheetName = "Duplicate Keys";

        private const string MetadataSheetName = "_PT_Meta";
        private const string Marker = "PT_REVISION_COMPARISON_REVIEW";
        private const string SchemaVersion = "1";
        private const int MaximumMetadataRows = 10000;
        private const int MaximumSelectionRows = 1000000;

        public static void Write(string outputPath, RevisionComparisonResult result)
        {
            if (result == null) throw new ArgumentNullException(nameof(result));
            string destination = NormalizeOutput(outputPath);
            foreach (RevisionComparisonSourceInspection source in result.Inspection.Sources)
            {
                if (string.Equals(destination, source.Path, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The review workbook cannot overwrite a source revision.");
            }

            string temporary = destination + ".writing-" + Guid.NewGuid().ToString("N");
            try
            {
                CreateWorkbook(temporary, result);
                RevisionComparisonReviewSelection reopened = ReadSelection(temporary);
                if (!string.Equals(reopened.PlanFingerprint, result.PlanFingerprint, StringComparison.Ordinal) ||
                    reopened.SelectedActionIds.Count != 0)
                    throw new InvalidDataException("The generated comparison review did not pass its reopen check.");
                Publish(temporary, destination);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }

        public static RevisionComparisonReviewSelection ReadSelection(string reviewPath)
        {
            if (string.IsNullOrWhiteSpace(reviewPath))
                throw new ArgumentException("A comparison review path is required.", nameof(reviewPath));
            string fullPath = System.IO.Path.GetFullPath(reviewPath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Comparison review not found.", fullPath);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fullPath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart ??
                    throw new InvalidDataException("The review workbook has no workbook part.");
                IReadOnlyList<IReadOnlyList<string>> metadata = ReadRows(
                    workbookPart,
                    MetadataSheetName,
                    MaximumMetadataRows);
                if (metadata.Count == 0 ||
                    Cell(metadata[0], 0) != Marker ||
                    Cell(metadata[0], 1) != SchemaVersion ||
                    string.IsNullOrWhiteSpace(Cell(metadata[0], 2)))
                    throw new InvalidDataException("This is not a supported Excel Power Tools revision-comparison review.");

                string fingerprint = Cell(metadata[0], 2);
                var sources = new Dictionary<int, RevisionComparisonReviewSource>();
                var order = new SortedDictionary<int, int>();
                var pairs = new List<RevisionComparisonPairSpec>();
                int? target = null;
                for (int rowIndex = 1; rowIndex < metadata.Count; rowIndex++)
                {
                    IReadOnlyList<string> row = metadata[rowIndex];
                    string kind = Cell(row, 0);
                    if (kind.Length == 0) continue;
                    switch (kind)
                    {
                        case "SOURCE":
                            int sourceIndex = ParseInt(Cell(row, 1), "source index", 0, RevisionComparisonEngine.MaximumRevisionCount - 1);
                            if (sources.ContainsKey(sourceIndex))
                                throw new InvalidDataException("The review metadata repeats a source index.");
                            string sourcePath = Cell(row, 2);
                            string sourceHash = Cell(row, 3);
                            long sourceLength = ParseLong(Cell(row, 4), "source byte length", 0, long.MaxValue);
                            if (string.IsNullOrWhiteSpace(sourcePath) || sourceHash.Length != 64)
                                throw new InvalidDataException("The review contains invalid source metadata.");
                            sources.Add(sourceIndex, new RevisionComparisonReviewSource(
                                sourceIndex,
                                System.IO.Path.GetFullPath(sourcePath),
                                sourceHash,
                                sourceLength));
                            break;

                        case "ORDER":
                            int position = ParseInt(Cell(row, 1), "revision position", 0, RevisionComparisonEngine.MaximumRevisionCount - 1);
                            int orderedSource = ParseInt(Cell(row, 2), "ordered source index", 0, RevisionComparisonEngine.MaximumRevisionCount - 1);
                            if (order.ContainsKey(position))
                                throw new InvalidDataException("The review metadata repeats a revision position.");
                            order.Add(position, orderedSource);
                            break;

                        case "PAIR":
                            pairs.Add(new RevisionComparisonPairSpec(
                                ParseInt(Cell(row, 1), "before revision index", 0, RevisionComparisonEngine.MaximumRevisionCount - 1),
                                ParseInt(Cell(row, 2), "after revision index", 0, RevisionComparisonEngine.MaximumRevisionCount - 1),
                                Cell(row, 3),
                                Cell(row, 4),
                                ParseInt(Cell(row, 5), "before header row", 1, 1048576),
                                ParseInt(Cell(row, 6), "after header row", 1, 1048576),
                                DecodeList(Cell(row, 7)),
                                DecodeList(Cell(row, 8))));
                            break;

                        case "TARGET":
                            if (target.HasValue) throw new InvalidDataException("The review metadata repeats the apply target.");
                            target = ParseInt(Cell(row, 1), "apply target", 0, RevisionComparisonEngine.MaximumRevisionCount - 1);
                            break;

                        default:
                            throw new InvalidDataException("The review metadata contains an unknown record: " + kind);
                    }
                }

                if (sources.Count < 2 || sources.Count > RevisionComparisonEngine.MaximumRevisionCount ||
                    sources.Keys.OrderBy(item => item).Where((item, index) => item != index).Any())
                    throw new InvalidDataException("The review source list is incomplete or non-contiguous.");
                if (order.Count != sources.Count || order.Keys.Where((item, index) => item != index).Any() ||
                    order.Values.Distinct().Count() != sources.Count ||
                    order.Values.Any(item => !sources.ContainsKey(item)))
                    throw new InvalidDataException("The review revision order is invalid.");
                if (pairs.Count == 0 || !target.HasValue)
                    throw new InvalidDataException("The review comparison plan is incomplete.");

                var plan = new RevisionComparisonPlan(order.Values, pairs, target.Value);
                var selected = new HashSet<string>(StringComparer.Ordinal);
                CollectSelected(workbookPart, AddedSheetName, selected);
                CollectSelected(workbookPart, ChangedSheetName, selected);
                return new RevisionComparisonReviewSelection(
                    sources.Values,
                    plan,
                    fingerprint,
                    selected.OrderBy(item => item, StringComparer.Ordinal));
            }
        }

        private static void CreateWorkbook(string path, RevisionComparisonResult result)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateStylesheet();
                stylesPart.Stylesheet.Save();
                var sheets = new Sheets();
                workbookPart.Workbook.Append(sheets);
                uint sheetId = 1;

                AddSheet(workbookPart, sheets, sheetId++, SummarySheetName, SummaryRows(result), false, false);
                AddSheet(workbookPart, sheets, sheetId++, AddedSheetName, AddedRows(result), true, true);
                AddSheet(workbookPart, sheets, sheetId++, RemovedSheetName, RemovedRows(result), false, false);
                AddSheet(workbookPart, sheets, sheetId++, ChangedSheetName, ChangedRows(result), true, true);
                AddSheet(workbookPart, sheets, sheetId++, DuplicateKeysSheetName, DuplicateRows(result), false, false);
                AddSheet(workbookPart, sheets, sheetId, MetadataSheetName, MetadataRows(result), false, false, true);

                workbookPart.Workbook.CalculationProperties = new CalculationProperties
                {
                    CalculationMode = CalculateModeValues.Manual,
                    FullCalculationOnLoad = false,
                    ForceFullCalculation = false,
                };
                workbookPart.Workbook.Save();
            }
        }

        private static IEnumerable<RowData> SummaryRows(RevisionComparisonResult result)
        {
            yield return RowData.Header("Keyed revision comparison");
            yield return RowData.Data("Use", "Review Added, Removed, Changed and Duplicate Keys. Type Yes in Apply only for actions you want copied to a new workbook.");
            yield return RowData.Data("Safety", "Apply re-reads every source, rejects external changes, writes a new workbook and verifies it before publication.");
            yield return RowData.Data("Changed cells", result.ChangedCellCount.ToString(CultureInfo.InvariantCulture));
            yield return RowData.Data("Added rows", result.AddedRowCount.ToString(CultureInfo.InvariantCulture));
            yield return RowData.Data("Removed rows", result.RemovedRows.Count.ToString(CultureInfo.InvariantCulture));
            yield return RowData.Data("Blank/duplicate key findings", result.DuplicateKeys.Count.ToString(CultureInfo.InvariantCulture));
            yield return RowData.Data("");
            yield return RowData.Header("Revision order", "File", "SHA-256");
            for (int position = 0; position < result.Plan.OrderedSourceIndexes.Count; position++)
            {
                RevisionComparisonSourceInspection source = result.Inspection.Sources[result.Plan.OrderedSourceIndexes[position]];
                yield return RowData.Data(
                    (position + 1).ToString(CultureInfo.InvariantCulture),
                    source.Path,
                    source.ContentSha256);
            }
            yield return RowData.Data("");
            yield return RowData.Header(
                "Transition",
                "Sheet pair",
                "Before rows",
                "Before blank keys",
                "Before duplicate rows",
                "After rows",
                "After blank keys",
                "After duplicate rows");
            foreach (RevisionComparisonKeyPreview preview in result.KeyPreviews)
            {
                yield return RowData.Data(
                    preview.Transition,
                    preview.SheetPair,
                    preview.BeforeRows.ToString(CultureInfo.InvariantCulture),
                    preview.BeforeBlankKeys.ToString(CultureInfo.InvariantCulture),
                    preview.BeforeDuplicateRows.ToString(CultureInfo.InvariantCulture),
                    preview.AfterRows.ToString(CultureInfo.InvariantCulture),
                    preview.AfterBlankKeys.ToString(CultureInfo.InvariantCulture),
                    preview.AfterDuplicateRows.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static IEnumerable<RowData> AddedRows(RevisionComparisonResult result)
        {
            yield return RowData.Header("Action ID", "Apply", "Transition", "Sheet", "Key", "Source row", "Values and formulas");
            foreach (RevisionComparisonAction action in result.Actions.Where(item => item.Kind == RevisionComparisonActionKind.AddedRow))
            {
                yield return RowData.Data(
                    action.ActionId,
                    "No",
                    action.Transition,
                    action.AfterSheetName,
                    action.KeyDisplay,
                    action.AfterRow.ToString(CultureInfo.InvariantCulture),
                    DescribeRow(action.RowCells));
            }
        }

        private static IEnumerable<RowData> RemovedRows(RevisionComparisonResult result)
        {
            yield return RowData.Header("Transition", "Sheet", "Key", "Source row", "Apply support");
            foreach (RevisionComparisonRemovedRow row in result.RemovedRows)
            {
                yield return RowData.Data(
                    row.Transition,
                    row.Sheet,
                    row.KeyDisplay,
                    row.SourceRow.ToString(CultureInfo.InvariantCulture),
                    "Review only in this first working version");
            }
        }

        private static IEnumerable<RowData> ChangedRows(RevisionComparisonResult result)
        {
            yield return RowData.Header(
                "Action ID",
                "Apply",
                "Transition",
                "Sheet",
                "Key",
                "Column",
                "Before value",
                "After value",
                "Before formula",
                "After formula",
                "Before row",
                "After row");
            foreach (RevisionComparisonAction action in result.Actions.Where(item => item.Kind == RevisionComparisonActionKind.CellChange))
            {
                yield return RowData.Data(
                    action.ActionId,
                    "No",
                    action.Transition,
                    action.AfterSheetName,
                    action.KeyDisplay,
                    action.AfterColumnHeader,
                    RevisionComparisonEngine.ValueDisplay(action.BeforeState),
                    RevisionComparisonEngine.ValueDisplay(action.AfterState),
                    action.BeforeState?.HasFormula == true ? "=" + action.BeforeState.FormulaText : "",
                    action.AfterState?.HasFormula == true ? "=" + action.AfterState.FormulaText : "",
                    action.BeforeRow.ToString(CultureInfo.InvariantCulture),
                    action.AfterRow.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static IEnumerable<RowData> DuplicateRows(RevisionComparisonResult result)
        {
            yield return RowData.Header("Transition", "Side", "Sheet", "Key", "Rows", "Effect");
            foreach (RevisionComparisonDuplicateKey item in result.DuplicateKeys)
            {
                yield return RowData.Data(
                    item.Transition,
                    item.Side,
                    item.Sheet,
                    item.KeyDisplay,
                    string.Join(", ", item.Rows.Select(row => row.ToString(CultureInfo.InvariantCulture))),
                    "Excluded from automatic matching and apply");
            }
        }

        private static IEnumerable<RowData> MetadataRows(RevisionComparisonResult result)
        {
            yield return RowData.Data(Marker, SchemaVersion, result.PlanFingerprint);
            for (int index = 0; index < result.Inspection.Sources.Count; index++)
            {
                RevisionComparisonSourceInspection source = result.Inspection.Sources[index];
                yield return RowData.Data(
                    "SOURCE",
                    index.ToString(CultureInfo.InvariantCulture),
                    source.Path,
                    source.ContentSha256,
                    source.ByteLength.ToString(CultureInfo.InvariantCulture));
            }
            for (int position = 0; position < result.Plan.OrderedSourceIndexes.Count; position++)
            {
                yield return RowData.Data(
                    "ORDER",
                    position.ToString(CultureInfo.InvariantCulture),
                    result.Plan.OrderedSourceIndexes[position].ToString(CultureInfo.InvariantCulture));
            }
            foreach (RevisionComparisonPairSpec pair in result.Plan.Pairs)
            {
                yield return RowData.Data(
                    "PAIR",
                    pair.BeforeRevisionIndex.ToString(CultureInfo.InvariantCulture),
                    pair.AfterRevisionIndex.ToString(CultureInfo.InvariantCulture),
                    pair.BeforeSheetName,
                    pair.AfterSheetName,
                    pair.BeforeHeaderRow.ToString(CultureInfo.InvariantCulture),
                    pair.AfterHeaderRow.ToString(CultureInfo.InvariantCulture),
                    EncodeList(pair.BeforeKeyHeaders),
                    EncodeList(pair.AfterKeyHeaders));
            }
            yield return RowData.Data("TARGET", result.Plan.ApplyTargetRevisionIndex.ToString(CultureInfo.InvariantCulture));
        }

        private static void AddSheet(
            WorkbookPart workbookPart,
            Sheets sheets,
            uint sheetId,
            string name,
            IEnumerable<RowData> sourceRows,
            bool hideFirstColumn,
            bool hasApplyColumn,
            bool veryHidden = false)
        {
            List<RowData> rows = sourceRows.ToList();
            int columnCount = Math.Max(1, rows.Count == 0 ? 1 : rows.Max(item => item.Values.Count));
            var worksheet = new Worksheet();
            worksheet.Append(new SheetViews(new SheetView(
                new Pane
                {
                    VerticalSplit = 1D,
                    TopLeftCell = "A2",
                    ActivePane = PaneValues.BottomLeft,
                    State = PaneStateValues.Frozen,
                })
            {
                WorkbookViewId = 0U,
            }));

            var columns = new Columns();
            for (int column = 1; column <= columnCount; column++)
            {
                double width = rows.Select(item => item.Values.Count >= column ? item.Values[column - 1]?.Length ?? 0 : 0)
                    .DefaultIfEmpty(0)
                    .Max();
                width = Math.Min(80D, Math.Max(10D, width * 1.15D + 3D));
                columns.Append(new Column
                {
                    Min = (uint)column,
                    Max = (uint)column,
                    Width = width,
                    CustomWidth = true,
                    Hidden = hideFirstColumn && column == 1,
                });
            }
            worksheet.Append(columns);

            var sheetData = new SheetData();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                RowData source = rows[rowIndex];
                uint excelRow = checked((uint)rowIndex + 1U);
                var row = new Row { RowIndex = excelRow };
                for (int column = 0; column < source.Values.Count; column++)
                {
                    row.Append(TextCell(
                        ColumnName(column + 1) + excelRow.ToString(CultureInfo.InvariantCulture),
                        source.Values[column],
                        source.IsHeader ? 1U : 0U));
                }
                sheetData.Append(row);
            }
            worksheet.Append(sheetData);
            if (rows.Count > 1 && columnCount > 1)
            {
                worksheet.Append(new AutoFilter
                {
                    Reference = "A1:" + ColumnName(columnCount) + rows.Count.ToString(CultureInfo.InvariantCulture),
                });
            }
            if (hasApplyColumn && rows.Count > 1)
            {
                var validation = new DataValidation
                {
                    Type = DataValidationValues.List,
                    AllowBlank = false,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:B1048576" },
                };
                validation.Append(new Formula1("\"No,Yes\""));
                worksheet.Append(new DataValidations(validation) { Count = 1U });
            }

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name,
            };
            if (veryHidden) sheet.State = SheetStateValues.VeryHidden;
            sheets.Append(sheet);
        }

        private static Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet(
                new Fonts(
                    new Font(),
                    new Font(new Bold(), new Color { Rgb = "FFFFFFFF" })),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                    new Fill(new PatternFill(
                        new ForegroundColor { Rgb = "FF4472C4" },
                        new BackgroundColor { Indexed = 64U })
                    {
                        PatternType = PatternValues.Solid,
                    })),
                new Borders(new Border()),
                new CellStyleFormats(new CellFormat()),
                new CellFormats(
                    new CellFormat(),
                    new CellFormat
                    {
                        FontId = 1U,
                        FillId = 2U,
                        BorderId = 0U,
                        FormatId = 0U,
                        ApplyFont = true,
                        ApplyFill = true,
                    }),
                new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }));
            return stylesheet;
        }

        private static Cell TextCell(string reference, string value, uint styleIndex)
        {
            var text = new DocumentFormat.OpenXml.Spreadsheet.Text(Truncate(value ?? ""));
            if (text.Text.Length != text.Text.Trim().Length) text.Space = SpaceProcessingModeValues.Preserve;
            return new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                StyleIndex = styleIndex,
                InlineString = new InlineString(text),
            };
        }

        private static string DescribeRow(IEnumerable<RevisionComparisonCellPayload> cells) =>
            Truncate(string.Join(" | ", cells.Select(cell =>
            {
                string value = RevisionComparisonEngine.ValueDisplay(cell.State);
                string formula = cell.State?.HasFormula == true ? "=" + cell.State.FormulaText : "";
                string content = formula.Length == 0 ? value : formula + (value.Length == 0 ? "" : " -> " + value);
                return cell.Header + "=" + content;
            })));

        private static void CollectSelected(WorkbookPart workbookPart, string sheetName, ISet<string> selected)
        {
            IReadOnlyList<IReadOnlyList<string>> rows = ReadRows(workbookPart, sheetName, MaximumSelectionRows);
            if (rows.Count == 0 ||
                !string.Equals(Cell(rows[0], 0), "Action ID", StringComparison.Ordinal) ||
                !string.Equals(Cell(rows[0], 1), "Apply", StringComparison.Ordinal))
                throw new InvalidDataException("The " + sheetName + " sheet does not have the expected comparison columns.");
            for (int index = 1; index < rows.Count; index++)
            {
                if (!IsSelected(Cell(rows[index], 1))) continue;
                string actionId = Cell(rows[index], 0);
                if (string.IsNullOrWhiteSpace(actionId))
                    throw new InvalidDataException("A selected " + sheetName + " row has no action ID.");
                selected.Add(actionId);
            }
        }

        private static IReadOnlyList<IReadOnlyList<string>> ReadRows(
            WorkbookPart workbookPart,
            string sheetName,
            int maximumRows)
        {
            Sheet sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(item =>
                string.Equals(item.Name?.Value, sheetName, StringComparison.Ordinal));
            if (sheet?.Id == null) throw new InvalidDataException("Required review sheet is missing: " + sheetName);
            WorksheetPart part = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
            if (part?.Worksheet == null) throw new InvalidDataException("Review sheet is unreadable: " + sheetName);

            List<SharedStringItem> sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>().ToList() ?? new List<SharedStringItem>();
            var result = new List<IReadOnlyList<string>>();
            foreach (Row row in part.Worksheet.Descendants<Row>())
            {
                if (result.Count >= maximumRows)
                    throw new InvalidDataException("The review sheet exceeds its safe row limit: " + sheetName);
                var values = new Dictionary<int, string>();
                int fallbackColumn = 1;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    int column = ColumnIndex(cell.CellReference?.Value);
                    if (column <= 0) column = fallbackColumn;
                    fallbackColumn = column + 1;
                    values[column] = ReadCellText(cell, sharedStrings);
                }
                int maximumColumn = values.Count == 0 ? 0 : values.Keys.Max();
                var ordered = new string[maximumColumn];
                foreach (KeyValuePair<int, string> item in values) ordered[item.Key - 1] = item.Value;
                result.Add(new ReadOnlyCollection<string>(ordered));
            }
            return new ReadOnlyCollection<IReadOnlyList<string>>(result);
        }

        private static string ReadCellText(Cell cell, IReadOnlyList<SharedStringItem> sharedStrings)
        {
            if (cell == null) return "";
            if (cell.DataType?.Value == CellValues.InlineString) return cell.InlineString?.InnerText ?? "";
            string raw = cell.CellValue?.Text ?? cell.InnerText ?? "";
            if (cell.DataType?.Value != CellValues.SharedString) return raw;
            if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int index) ||
                index < 0 || index >= sharedStrings.Count)
                throw new InvalidDataException("The review contains an invalid shared-string reference.");
            return sharedStrings[index].InnerText ?? "";
        }

        private static bool IsSelected(string value)
        {
            string normalized = (value ?? "").Trim();
            return string.Equals(normalized, "yes", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "y", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "x", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "true", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "apply", StringComparison.OrdinalIgnoreCase) ||
                   normalized == "1";
        }

        private static string EncodeList(IEnumerable<string> values) => string.Join(",", values.Select(value =>
            Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? ""))));

        private static IReadOnlyList<string> DecodeList(string encoded)
        {
            if (string.IsNullOrEmpty(encoded)) return Array.Empty<string>();
            try
            {
                return new ReadOnlyCollection<string>(encoded.Split(',')
                    .Select(item => Encoding.UTF8.GetString(Convert.FromBase64String(item)))
                    .ToList());
            }
            catch (FormatException exception)
            {
                throw new InvalidDataException("The review contains invalid encoded key columns.", exception);
            }
        }

        private static string NormalizeOutput(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A review output path is required.", nameof(path));
            string full = System.IO.Path.GetFullPath(path);
            if (!string.Equals(System.IO.Path.GetExtension(full), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("The comparison review must be saved as a macro-free .xlsx workbook.");
            string directory = System.IO.Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException("The comparison review folder does not exist: " + directory);
            return full;
        }

        private static void Publish(string temporary, string destination)
        {
            if (File.Exists(destination)) File.Replace(temporary, destination, null, true);
            else File.Move(temporary, destination);
        }

        private static int ParseInt(string value, string label, int minimum, int maximum)
        {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) ||
                parsed < minimum || parsed > maximum)
                throw new InvalidDataException("The review contains an invalid " + label + ".");
            return parsed;
        }

        private static long ParseLong(string value, string label, long minimum, long maximum)
        {
            if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) ||
                parsed < minimum || parsed > maximum)
                throw new InvalidDataException("The review contains an invalid " + label + ".");
            return parsed;
        }

        private static int ColumnIndex(string reference)
        {
            if (string.IsNullOrEmpty(reference)) return 0;
            int result = 0;
            int index = 0;
            while (index < reference.Length && reference[index] >= 'A' && reference[index] <= 'Z')
            {
                result = checked(result * 26 + reference[index] - 'A' + 1);
                index++;
            }
            return result;
        }

        private static string ColumnName(int column)
        {
            if (column < 1 || column > 16384) throw new ArgumentOutOfRangeException(nameof(column));
            var value = new StringBuilder();
            while (column > 0)
            {
                column--;
                value.Insert(0, (char)('A' + column % 26));
                column /= 26;
            }
            return value.ToString();
        }

        private static string Cell(IReadOnlyList<string> row, int index) =>
            row != null && index >= 0 && index < row.Count ? row[index] ?? "" : "";

        private static string Truncate(string value) =>
            value != null && value.Length > 32000 ? value.Substring(0, 31980) + " ... [truncated]" : value ?? "";

        private sealed class RowData
        {
            private RowData(bool isHeader, IEnumerable<string> values)
            {
                IsHeader = isHeader;
                Values = new ReadOnlyCollection<string>(values.Select(Truncate).ToList());
            }

            internal bool IsHeader { get; }
            internal IReadOnlyList<string> Values { get; }
            internal static RowData Header(params string[] values) => new RowData(true, values);
            internal static RowData Data(params string[] values) => new RowData(false, values);
        }
    }
}
