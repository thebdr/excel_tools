using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    public static partial class ReviewWorkbookFile
    {
        private const int MaxStreamingCellCount = 8_000_000;
        private const int MaxStreamingHyperlinkCount = 100_000;
        private const int MaxStreamingSharedStringCount = MaxStreamingCellCount;
        private const long MaxStreamingSharedStringSpoolBytes = 256L * 1024 * 1024;

        private static string ComputeProjectionSha256(
            WorkbookPart workbookPart,
            string projectionDigestVersion,
            ReviewWorkbookProjectionValidationMetrics metrics = null)
        {
            if (string.Equals(
                    projectionDigestVersion,
                    LegacyProjectionDigestVersion,
                    StringComparison.Ordinal))
            {
                ValidateLegacyProjectionSafety(workbookPart);
                return ComputeLegacyProjectionSha256(workbookPart);
            }
            if (!string.Equals(
                    projectionDigestVersion,
                    StreamingProjectionDigestVersion,
                    StringComparison.Ordinal))
                throw new InvalidDataException("The review projection digest version is unsupported.");
            return ComputeStreamingProjectionSha256(workbookPart, metrics);
        }

        private static void ValidateLegacyProjectionSafety(WorkbookPart workbookPart)
        {
            SharedStringTablePart sharedPart = workbookPart.SharedStringTablePart;
            if (sharedPart != null)
            {
                using (Stream stream = sharedPart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    if (stream.Length > MaxSpreadsheetPartCharacters)
                        throw new InvalidDataException(
                            "A legacy review workbook shared-string part exceeds its safe limit.");
                }
                SharedStringTable table = sharedPart.SharedStringTable;
                if (table == null || table.Descendants<RunProperties>().Any(HasUnsupportedRunProperties))
                    throw new InvalidDataException(
                        "A comparison review workbook contains unsafe shared-string presentation properties.");
            }

            // Legacy projection digests retain their exact v1 ordering and rich-text semantics, but the
            // passive security boundary is version-independent. Run the bounded raw worksheet grammar and
            // presentation pass first so legacy files cannot hide formula-bearing validation prompts, panes
            // or far viewports in content that the v1 semantic digest intentionally did not cover.
            using (StreamingSharedStringStore sharedStrings =
                   ReadStreamingSharedStrings(sharedPart, metrics: null))
            using (SHA256 safetyHash = SHA256.Create())
            {
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                {
                    using (Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        if (stream.Length > MaxSpreadsheetPartCharacters)
                            throw new InvalidDataException(
                                "A legacy review workbook worksheet exceeds its 64 MiB safe limit.");
                    }
                    StreamingWorksheetScan scan = ScanStreamingWorksheet(worksheetPart);
                    HashStreamingWorksheetCells(
                        worksheetPart,
                        scan,
                        sharedStrings,
                        safetyHash);
                }
                safetyHash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            }

            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
            {
                using (Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    if (stream.Length > MaxSpreadsheetPartCharacters)
                        throw new InvalidDataException(
                            "A legacy review workbook worksheet exceeds its 64 MiB safe limit.");
                }
                Worksheet worksheet = worksheetPart.Worksheet;
                if (worksheet == null)
                    throw new InvalidDataException("A legacy review worksheet is empty.");
                if (worksheet.Descendants<CellFormula>().Any())
                    throw new InvalidDataException(
                        "A comparison review workbook must not contain executable worksheet formulas.");
                if (HasUnsupportedPresentation(worksheet))
                    throw new InvalidDataException(
                        "A comparison review workbook contains presentation changes that can conceal or " +
                        "relabel protected evidence.");
            }
        }

        private static string ComputeStreamingProjectionSha256(
            WorkbookPart workbookPart,
            ReviewWorkbookProjectionValidationMetrics metrics)
        {
            Sheets sheets = workbookPart.Workbook?.Sheets;
            if (sheets == null)
                throw new InvalidDataException("The review workbook has no sheets collection.");
            List<Sheet> actualSheets = sheets.Elements<Sheet>().ToList();
            if (actualSheets.Count != ExpectedSheetNames.Count)
                throw new InvalidDataException("The review workbook sheet count is invalid.");

            using (StreamingSharedStringStore sharedStrings =
                   ReadStreamingSharedStrings(workbookPart.SharedStringTablePart, metrics))
            using (SHA256 sha256 = SHA256.Create())
            {
                HashProjectionValue(sha256, StreamingProjectionDigestVersion);
                for (int index = 0; index < actualSheets.Count; index++)
                {
                    Sheet sheet = actualSheets[index];
                    string name = sheet.Name?.Value ?? "";
                    if (!string.Equals(name, ExpectedSheetNames[index], StringComparison.Ordinal))
                        throw new InvalidDataException("The review workbook sheet order or name is invalid.");
                    if (sheet.State != null && sheet.State.Value != SheetStateValues.Visible)
                        throw new InvalidDataException("Every review workbook sheet must remain visible.");
                    if (sheet.Id == null || string.IsNullOrEmpty(sheet.Id.Value))
                        throw new InvalidDataException("A review workbook sheet relationship is missing.");

                    WorksheetPart worksheetPart;
                    try
                    {
                        worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    }
                    catch (Exception exception) when (
                        exception is ArgumentOutOfRangeException ||
                        exception is KeyNotFoundException ||
                        exception is InvalidOperationException)
                    {
                        throw new InvalidDataException(
                            "A review workbook sheet relationship is invalid.",
                            exception);
                    }
                    if (worksheetPart == null)
                        throw new InvalidDataException("A review workbook sheet is not a worksheet.");
                    using (Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        if (stream.Length > MaxStreamingWorksheetCharacters)
                            throw new InvalidDataException(
                                "A streaming review worksheet exceeds its safe expansion limit.");
                    }

                    HashProjectionValue(sha256, name);
                    StreamingWorksheetScan scan = ScanStreamingWorksheet(worksheetPart);
                    if (metrics != null)
                    {
                        metrics.WorksheetPassCount += 2;
                        metrics.CellCount += scan.CellCount;
                        metrics.HyperlinkCount += scan.Hyperlinks.Count;
                    }
                    HashProjectionValue(
                        sha256,
                        scan.CellCount.ToString(CultureInfo.InvariantCulture));
                    HashStreamingWorksheetCells(
                        worksheetPart,
                        scan,
                        sharedStrings,
                        sha256);
                    HashProjectionValue(
                        sha256,
                        scan.Hyperlinks.Count.ToString(CultureInfo.InvariantCulture));
                    foreach (StreamingHyperlink hyperlink in scan.Hyperlinks)
                    {
                        string display = hyperlink.Display;
                        if (hyperlink.VisibleCellText != null &&
                            string.Equals(display, hyperlink.VisibleCellText, StringComparison.Ordinal))
                            display = "";
                        HashProjectionValue(sha256, hyperlink.Reference);
                        HashProjectionValue(sha256, hyperlink.Location);
                        HashProjectionValue(sha256, display);
                        HashProjectionValue(sha256, hyperlink.Tooltip);
                    }
                }

                sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return BytesToHex(sha256.Hash);
            }
        }

        private static StreamingWorksheetScan ScanStreamingWorksheet(WorksheetPart worksheetPart)
        {
            var hyperlinks = new List<StreamingHyperlink>();
            StreamingCellReference? priorHyperlink = null;
            int cellCount = 0;
            using (Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(stream, MaxStreamingWorksheetCharacters))
            {
                if (reader.MoveToContent() != XmlNodeType.Element ||
                    reader.LocalName != "worksheet" || reader.NamespaceURI != SpreadsheetNamespace)
                    throw new InvalidDataException("A streaming review worksheet has an invalid root element.");
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.NamespaceURI != SpreadsheetNamespace)
                        continue;
                    if (reader.LocalName == "c")
                    {
                        if (cellCount == MaxStreamingCellCount)
                            throw new InvalidDataException(
                                "A streaming review worksheet exceeds the safe cell-count limit.");
                        cellCount++;
                    }
                    else if (reader.LocalName == "hyperlink")
                    {
                        if (hyperlinks.Count == MaxStreamingHyperlinkCount)
                            throw new InvalidDataException(
                                "A streaming review worksheet exceeds the safe hyperlink-count limit.");
                        string reference = reader.GetAttribute("ref") ?? "";
                        StreamingCellReference parsed = ParseStreamingCellReference(reference);
                        if (priorHyperlink.HasValue &&
                            CompareStreamingReferences(priorHyperlink.Value, parsed) >= 0)
                            throw new InvalidDataException(
                                "Review workbook hyperlink references must be unique and monotonic.");
                        priorHyperlink = parsed;
                        string relationshipId = reader.GetAttribute(
                            "id",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ?? "";
                        if (relationshipId.Length != 0)
                            throw new InvalidDataException("Review workbook hyperlinks must be internal.");
                        ValidateStreamingHyperlinkAttributes(reader);
                        hyperlinks.Add(new StreamingHyperlink(
                            parsed,
                            reference,
                            reader.GetAttribute("location") ?? "",
                            reader.GetAttribute("display") ?? "",
                            reader.GetAttribute("tooltip") ?? ""));
                    }
                }
            }
            return new StreamingWorksheetScan(
                cellCount,
                new ReadOnlyCollection<StreamingHyperlink>(hyperlinks));
        }

        private static void HashStreamingWorksheetCells(
            WorksheetPart worksheetPart,
            StreamingWorksheetScan scan,
            StreamingSharedStringStore sharedStrings,
            HashAlgorithm sha256)
        {
            int sheetDataCount = 0;
            int sheetDataDepth = -1;
            int rowDepth = -1;
            int currentRow = 0;
            int priorRow = 0;
            int priorColumn = 0;
            int cellCount = 0;
            int hyperlinkIndex = 0;

            using (Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(stream, MaxStreamingWorksheetCharacters))
            {
                if (reader.MoveToContent() != XmlNodeType.Element ||
                    reader.LocalName != "worksheet" || reader.NamespaceURI != SpreadsheetNamespace)
                    throw new InvalidDataException("A streaming review worksheet has an invalid root element.");
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.EndElement &&
                        reader.NamespaceURI == SpreadsheetNamespace)
                    {
                        if (reader.LocalName == "row" && reader.Depth == rowDepth)
                        {
                            rowDepth = -1;
                            currentRow = 0;
                            priorColumn = 0;
                        }
                        else if (reader.LocalName == "sheetData" && reader.Depth == sheetDataDepth)
                        {
                            sheetDataDepth = -1;
                        }
                        continue;
                    }
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.NamespaceURI != SpreadsheetNamespace)
                        continue;

                    switch (reader.LocalName)
                    {
                        case "sheetData":
                            if (++sheetDataCount != 1 || rowDepth >= 0)
                                throw new InvalidDataException(
                                    "A streaming review worksheet must contain one sheetData element.");
                            sheetDataDepth = reader.Depth;
                            break;
                        case "row":
                            if (sheetDataDepth < 0 || reader.Depth != sheetDataDepth + 1 || rowDepth >= 0)
                                throw new InvalidDataException(
                                    "A review workbook row is outside canonical sheetData order.");
                            currentRow = ParseCanonicalRowIndex(reader.GetAttribute("r"));
                            if (currentRow <= priorRow)
                                throw new InvalidDataException(
                                    "Review workbook row references must be unique and monotonic.");
                            priorRow = currentRow;
                            priorColumn = 0;
                            rowDepth = reader.Depth;
                            ValidateStreamingRowPresentation(reader);
                            if (reader.IsEmptyElement)
                            {
                                rowDepth = -1;
                                currentRow = 0;
                            }
                            break;
                        case "c":
                            if (rowDepth < 0 || reader.Depth != rowDepth + 1)
                                throw new InvalidDataException(
                                    "A review workbook cell is outside canonical row order.");
                            string reference = reader.GetAttribute("r") ?? "";
                            StreamingCellReference parsed = ParseStreamingCellReference(reference);
                            if (parsed.Row != currentRow || parsed.Column <= priorColumn)
                                throw new InvalidDataException(
                                    "Review workbook cell references must be canonical, unique and monotonic.");
                            priorColumn = parsed.Column;
                            ValidateStreamingCellAttributes(reader);
                            ValidateStreamingCellPresentation(reader);
                            StreamingProjectionText value = ReadStreamingCell(reader, sharedStrings);
                            HashProjectionValue(sha256, reference);
                            HashProjectionValue(sha256, value.Semantic);
                            while (hyperlinkIndex < scan.Hyperlinks.Count &&
                                   CompareStreamingReferences(
                                       scan.Hyperlinks[hyperlinkIndex].Coordinate,
                                       parsed) < 0)
                                hyperlinkIndex++;
                            if (hyperlinkIndex < scan.Hyperlinks.Count &&
                                CompareStreamingReferences(
                                    scan.Hyperlinks[hyperlinkIndex].Coordinate,
                                    parsed) == 0)
                                scan.Hyperlinks[hyperlinkIndex].VisibleCellText = value.Visible;
                            cellCount++;
                            break;
                        case "f":
                        case "formula1":
                        case "formula2":
                            throw new InvalidDataException(
                                "A comparison review workbook must not contain executable worksheet formulas.");
                        case "dataValidations":
                        case "dataValidation":
                            throw new InvalidDataException(
                                "A comparison review workbook must not contain data-validation formulas or prompts.");
                        case "sheetFormatPr":
                            ValidateStreamingSheetFormat(reader);
                            break;
                        case "sheetViews":
                            if (reader.HasAttributes)
                                throw UnsafeWorksheetPresentation();
                            break;
                        case "sheetView":
                            ValidateStreamingSheetView(reader);
                            break;
                        case "selection":
                            ValidateStreamingSelection(reader);
                            break;
                        case "pane":
                            throw UnsafeWorksheetPresentation();
                        case "col":
                            ValidateStreamingColumnPresentation(reader);
                            break;
                        case "mergeCell":
                        case "conditionalFormatting":
                        case "sheetProtection":
                        case "protectedRanges":
                        case "autoFilter":
                        case "sortState":
                        case "customSheetViews":
                        case "scenarios":
                        case "dataConsolidate":
                        case "drawing":
                        case "legacyDrawing":
                        case "legacyDrawingHF":
                        case "picture":
                        case "oleObjects":
                        case "controls":
                        case "tableParts":
                        case "extLst":
                            throw new InvalidDataException(
                                "A comparison review workbook contains presentation changes that can conceal " +
                                "or relabel protected evidence.");
                    }
                }
            }

            if (sheetDataCount != 1 || cellCount != scan.CellCount)
                throw new InvalidDataException(
                    "A streaming review worksheet has inconsistent sheet-data or cell counts.");
        }

        private static StreamingProjectionText ReadStreamingCell(
            XmlReader cellReader,
            StreamingSharedStringStore sharedStrings)
        {
            string type = cellReader.GetAttribute("t") ?? "";
            string raw = "";
            StreamingProjectionText inline = null;
            bool valueSeen = false;
            bool inlineSeen = false;
            int rootDepth;
            using (XmlReader reader = cellReader.ReadSubtree())
            {
                reader.Read();
                rootDepth = reader.Depth;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element)
                        continue;
                    if (reader.NamespaceURI == SpreadsheetNamespace && reader.LocalName == "f")
                        throw new InvalidDataException(
                            "A comparison review workbook must not contain executable worksheet formulas.");
                    if (reader.Depth != rootDepth + 1 ||
                        reader.NamespaceURI != SpreadsheetNamespace)
                        continue;
                    if (reader.LocalName == "v")
                    {
                        if (valueSeen)
                            throw new InvalidDataException("A review workbook cell repeats its value element.");
                        valueSeen = true;
                        raw = ReadBoundedElementText(reader, MaxExcelCellTextCharacters);
                    }
                    else if (reader.LocalName == "is")
                    {
                        if (inlineSeen)
                            throw new InvalidDataException(
                                "A review workbook cell repeats its inline-string element.");
                        inlineSeen = true;
                        inline = ReadStreamingProjectionText(reader);
                    }
                    else
                    {
                        throw new InvalidDataException(
                            "A review workbook cell contains unsupported or non-canonical child content.");
                    }
                }
            }

            switch (type)
            {
                case "s":
                    if (!valueSeen || inlineSeen)
                        throw InvalidCellTypeGrammar();
                    if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int index) ||
                        index < 0 || index >= sharedStrings.Count ||
                        !string.Equals(index.ToString(CultureInfo.InvariantCulture), raw, StringComparison.Ordinal))
                        throw new InvalidDataException("A review workbook shared-string index is invalid.");
                    StreamingProjectionText shared = sharedStrings.Get(index);
                    return new StreamingProjectionText(
                        "text:" + shared.Semantic,
                        shared.Visible);
                case "inlineStr":
                    if (!inlineSeen || valueSeen)
                        throw InvalidCellTypeGrammar();
                    inline = inline ?? new StreamingProjectionText("", "");
                    return new StreamingProjectionText("text:" + inline.Semantic, inline.Visible);
                case "str":
                    if (!valueSeen || inlineSeen)
                        throw InvalidCellTypeGrammar();
                    return new StreamingProjectionText("text:" + raw, raw);
                case "":
                case "n":
                case "b":
                case "e":
                case "d":
                    if (!valueSeen || inlineSeen)
                        throw InvalidCellTypeGrammar();
                    return new StreamingProjectionText("scalar:" + type + ":" + raw, null);
                default:
                    throw InvalidCellTypeGrammar();
            }
        }

        private static InvalidDataException InvalidCellTypeGrammar() => new InvalidDataException(
            "A review workbook cell type and child elements are inconsistent or unsupported.");

        private static StreamingSharedStringStore ReadStreamingSharedStrings(
            SharedStringTablePart sharedPart,
            ReviewWorkbookProjectionValidationMetrics metrics)
        {
            var strings = new StreamingSharedStringStore(
                MaxStreamingSharedStringCount,
                MaxStreamingSharedStringSpoolBytes,
                metrics);
            if (sharedPart == null) return strings;
            try
            {
                using (Stream stream = sharedPart.GetStream(FileMode.Open, FileAccess.Read))
                using (XmlReader reader = CreateBoundedXmlReader(
                           stream,
                           MaxStreamingSharedStringCharacters))
                {
                    if (stream.Length > MaxStreamingSharedStringCharacters ||
                        reader.MoveToContent() != XmlNodeType.Element ||
                        reader.LocalName != "sst" || reader.NamespaceURI != SpreadsheetNamespace)
                        throw new InvalidDataException("The review workbook shared-string table is invalid.");
                    int rootDepth = reader.Depth;
                    while (reader.Read())
                    {
                        if (reader.NodeType != XmlNodeType.Element ||
                            reader.NamespaceURI != SpreadsheetNamespace ||
                            reader.Depth != rootDepth + 1 || reader.LocalName != "si")
                            continue;
                        strings.Add(ReadStreamingProjectionText(reader));
                    }
                }
                strings.Seal();
                return strings;
            }
            catch
            {
                strings.Dispose();
                throw;
            }
        }

        private static StreamingProjectionText ReadStreamingProjectionText(XmlReader containerReader)
        {
            if (containerReader.HasAttributes)
                throw new InvalidDataException(
                    "A review workbook inline or shared string has unsupported attributes.");
            var styles = new List<string>();
            var texts = new List<StringBuilder>();
            int visibleLength = 0;
            using (XmlReader reader = containerReader.ReadSubtree())
            {
                reader.Read();
                int rootDepth = reader.Depth;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.NamespaceURI != SpreadsheetNamespace ||
                        reader.Depth != rootDepth + 1)
                        continue;
                    string style;
                    string text;
                    if (reader.LocalName == "t")
                    {
                        style = "";
                        text = ReadStreamingTextElement(
                            reader,
                            MaxExcelCellTextCharacters - visibleLength);
                    }
                    else if (reader.LocalName == "r")
                    {
                        StreamingRun run = ReadStreamingRun(reader, visibleLength);
                        style = run.Style;
                        text = run.Text;
                    }
                    else
                    {
                        if (reader.LocalName != "rPh" && reader.LocalName != "phoneticPr")
                            throw new InvalidDataException(
                                "A review workbook text value contains unsupported rich-text markup.");
                        style = "element:" + reader.LocalName;
                        text = ReadBoundedDescendantText(
                            reader,
                            MaxExcelCellTextCharacters - visibleLength);
                    }
                    visibleLength = checked(visibleLength + text.Length);
                    if (visibleLength > MaxExcelCellTextCharacters)
                        throw new InvalidDataException(
                            "A projected worksheet text value exceeds Excel's cell-text limit.");
                    AppendStreamingTextSegment(styles, texts, style, text);
                }
            }

            string visible = string.Concat(texts.Select(text => text.ToString()));
            if (styles.All(string.IsNullOrEmpty))
                return new StreamingProjectionText(visible, visible);
            var semantic = new StringBuilder("rich-v1");
            for (int index = 0; index < styles.Count; index++)
            {
                AppendProjectionToken(semantic, styles[index]);
                AppendProjectionToken(semantic, texts[index].ToString());
            }
            return new StreamingProjectionText(semantic.ToString(), visible);
        }

        private static StreamingRun ReadStreamingRun(XmlReader runReader, int existingVisibleLength)
        {
            if (runReader.HasAttributes)
                throw new InvalidDataException("A review workbook rich-text run has unsupported attributes.");
            string bold = "";
            string strike = "";
            string rgb = "";
            bool propertiesSeen = false;
            var text = new StringBuilder();
            using (XmlReader reader = runReader.ReadSubtree())
            {
                reader.Read();
                int rootDepth = reader.Depth;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.NamespaceURI != SpreadsheetNamespace ||
                        reader.Depth != rootDepth + 1)
                        continue;
                    if (reader.LocalName == "rPr")
                    {
                        if (propertiesSeen)
                            throw new InvalidDataException(
                                "A review workbook rich-text run repeats its properties.");
                        propertiesSeen = true;
                        StreamingRunStyle properties = ReadStreamingRunProperties(reader);
                        bold = properties.Bold;
                        strike = properties.Strike;
                        rgb = properties.Rgb;
                    }
                    else if (reader.LocalName == "t")
                    {
                        string value = ReadStreamingTextElement(
                            reader,
                            MaxExcelCellTextCharacters - existingVisibleLength - text.Length);
                        text.Append(value);
                    }
                    else
                    {
                        throw new InvalidDataException(
                            "A review workbook rich-text run contains unsupported content.");
                    }
                }
            }

            string style = "";
            if (bold.Length != 0 || strike.Length != 0 || rgb.Length != 0)
            {
                var signature = new StringBuilder();
                AppendProjectionToken(signature, bold);
                AppendProjectionToken(signature, strike);
                AppendProjectionToken(signature, rgb);
                AppendProjectionToken(signature, "");
                AppendProjectionToken(signature, "");
                style = signature.ToString();
            }
            return new StreamingRun(style, text.ToString());
        }

        private static StreamingRunStyle ReadStreamingRunProperties(XmlReader propertiesReader)
        {
            if (propertiesReader.HasAttributes)
                throw new InvalidDataException(
                    "A review workbook contains unsafe rich-text presentation properties.");
            int boldCount = 0;
            int strikeCount = 0;
            int colorCount = 0;
            int fontCount = 0;
            int sizeCount = 0;
            string bold = "";
            string strike = "";
            string rgb = "";
            using (XmlReader reader = propertiesReader.ReadSubtree())
            {
                reader.Read();
                int rootDepth = reader.Depth;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.NamespaceURI != SpreadsheetNamespace ||
                        reader.Depth != rootDepth + 1)
                        continue;
                    switch (reader.LocalName)
                    {
                        case "b":
                            if (++boldCount > 1 || !HasOnlyStreamingAttributes(reader, "val"))
                                throw UnsafeRichText();
                            if (StreamingBooleanEnabled(reader.GetAttribute("val"))) bold = "bold";
                            break;
                        case "strike":
                            if (++strikeCount > 1 || !HasOnlyStreamingAttributes(reader, "val"))
                                throw UnsafeRichText();
                            if (StreamingBooleanEnabled(reader.GetAttribute("val"))) strike = "strike";
                            break;
                        case "color":
                            rgb = (reader.GetAttribute("rgb") ?? "").ToUpperInvariant();
                            if (++colorCount > 1 || !HasOnlyStreamingAttributes(reader, "rgb") ||
                                !AllowedReviewRunRgb.Contains(rgb))
                                throw UnsafeRichText();
                            break;
                        case "rFont":
                            if (++fontCount > 1 || !HasOnlyStreamingAttributes(reader, "val") ||
                                !AllowedVisibleFontNames.Contains(reader.GetAttribute("val") ?? ""))
                                throw UnsafeRichText();
                            break;
                        case "sz":
                            if (++sizeCount > 1 || !HasOnlyStreamingAttributes(reader, "val") ||
                                !double.TryParse(
                                    reader.GetAttribute("val"),
                                    NumberStyles.Float,
                                    CultureInfo.InvariantCulture,
                                    out double size) || size < MinVisibleFontSizePoints)
                                throw UnsafeRichText();
                            break;
                        default:
                            throw UnsafeRichText();
                    }
                }
            }
            return new StreamingRunStyle(bold, strike, rgb);
        }

        private static InvalidDataException UnsafeRichText() => new InvalidDataException(
            "A comparison review workbook contains unsafe rich-text presentation properties.");

        private static bool HasOnlyStreamingAttributes(XmlReader reader, params string[] allowedNames)
        {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            if (!reader.HasAttributes) return true;
            while (reader.MoveToNextAttribute())
            {
                if (!string.IsNullOrEmpty(reader.NamespaceURI) || !allowed.Contains(reader.LocalName))
                {
                    reader.MoveToElement();
                    return false;
                }
            }
            reader.MoveToElement();
            return true;
        }

        private static bool StreamingBooleanEnabled(string value)
        {
            if (string.IsNullOrEmpty(value) || value == "1" ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "on", StringComparison.OrdinalIgnoreCase))
                return true;
            if (value == "0" ||
                string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "off", StringComparison.OrdinalIgnoreCase))
                return false;
            throw new InvalidDataException(
                "A review workbook boolean presentation value is invalid.");
        }

        private static void AppendStreamingTextSegment(
            IList<string> styles,
            IList<StringBuilder> texts,
            string style,
            string value)
        {
            if (styles.Count > 0 &&
                string.Equals(styles[styles.Count - 1], style, StringComparison.Ordinal))
            {
                texts[texts.Count - 1].Append(value);
                return;
            }
            styles.Add(style);
            texts.Add(new StringBuilder(value));
        }

        private static string ReadBoundedElementText(XmlReader elementReader, int maximumCharacters)
        {
            if (maximumCharacters < 0)
                throw new InvalidDataException(
                    "A projected worksheet text value exceeds Excel's cell-text limit.");
            var text = new StringBuilder();
            using (XmlReader reader = elementReader.ReadSubtree())
            {
                reader.Read();
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element)
                        throw new InvalidDataException(
                            "A worksheet text element contains unsupported nested markup.");
                    if (reader.NodeType == XmlNodeType.Text ||
                        reader.NodeType == XmlNodeType.CDATA ||
                        reader.NodeType == XmlNodeType.SignificantWhitespace ||
                        reader.NodeType == XmlNodeType.Whitespace)
                    {
                        if (text.Length > maximumCharacters - reader.Value.Length)
                            throw new InvalidDataException(
                                "A projected worksheet text value exceeds Excel's cell-text limit.");
                        text.Append(reader.Value);
                    }
                }
            }
            return text.ToString();
        }

        private static string ReadBoundedDescendantText(XmlReader elementReader, int maximumCharacters)
        {
            if (maximumCharacters < 0)
                throw new InvalidDataException(
                    "A projected worksheet text value exceeds Excel's cell-text limit.");
            var text = new StringBuilder();
            using (XmlReader reader = elementReader.ReadSubtree())
            {
                reader.Read();
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Text ||
                        reader.NodeType == XmlNodeType.CDATA ||
                        reader.NodeType == XmlNodeType.SignificantWhitespace)
                    {
                        if (text.Length > maximumCharacters - reader.Value.Length)
                            throw new InvalidDataException(
                                "A projected worksheet text value exceeds Excel's cell-text limit.");
                        text.Append(reader.Value);
                    }
                }
            }
            return text.ToString();
        }

        private static int ParseCanonicalRowIndex(string value)
        {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int row) ||
                row < 1 || row > 1_048_576 ||
                !string.Equals(row.ToString(CultureInfo.InvariantCulture), value, StringComparison.Ordinal))
                throw new InvalidDataException("A review workbook row reference is not canonical.");
            return row;
        }

        private static StreamingCellReference ParseStreamingCellReference(string value)
        {
            if (string.IsNullOrEmpty(value))
                throw new InvalidDataException("A review workbook cell reference is missing.");
            int offset = 0;
            int column = 0;
            while (offset < value.Length && value[offset] >= 'A' && value[offset] <= 'Z')
            {
                if (column > (16_384 - (value[offset] - 'A' + 1)) / 26)
                    throw new InvalidDataException("A review workbook cell column is outside Excel's limit.");
                column = checked(column * 26 + value[offset] - 'A' + 1);
                offset++;
            }
            if (column < 1 || column > 16_384 || offset == 0 || offset == value.Length)
                throw new InvalidDataException("A review workbook cell reference is not canonical.");
            string rowText = value.Substring(offset);
            int row = ParseCanonicalRowIndex(rowText);
            if (!string.Equals(ColumnName(column) + rowText, value, StringComparison.Ordinal))
                throw new InvalidDataException("A review workbook cell reference is not canonical.");
            return new StreamingCellReference(row, column);
        }

        private static int CompareStreamingReferences(
            StreamingCellReference left,
            StreamingCellReference right)
        {
            int row = left.Row.CompareTo(right.Row);
            return row != 0 ? row : left.Column.CompareTo(right.Column);
        }

        private static void ValidateStreamingSheetFormat(XmlReader reader)
        {
            if (StreamingBooleanEnabledWhenPresent(reader.GetAttribute("zeroHeight")) ||
                IsBelow(reader.GetAttribute("defaultRowHeight"), MinVisibleRowHeightPoints) ||
                IsBelow(reader.GetAttribute("defaultColWidth"), MinVisibleColumnWidth))
                throw UnsafeWorksheetPresentation();
        }

        private static void ValidateStreamingWorkbookPresentation(Workbook workbook)
        {
            foreach (WorkbookView view in workbook?.BookViews?.Elements<WorkbookView>() ??
                                          Enumerable.Empty<WorkbookView>())
            {
                if (view.Minimized?.Value == true ||
                    (view.ActiveTab != null && view.ActiveTab.Value != 0U) ||
                    (view.FirstSheet != null && view.FirstSheet.Value != 0U) ||
                    (view.Visibility != null &&
                     view.Visibility.Value != VisibilityValues.Visible) ||
                    view.ShowSheetTabs?.Value == false ||
                    view.ShowHorizontalScroll?.Value == false ||
                    view.ShowVerticalScroll?.Value == false ||
                    (view.WindowWidth != null && view.WindowWidth.Value < 1000U) ||
                    (view.WindowHeight != null && view.WindowHeight.Value < 1000U) ||
                    (view.XWindow != null &&
                     (view.XWindow.Value < -100000 || view.XWindow.Value > 100000)) ||
                    (view.YWindow != null &&
                     (view.YWindow.Value < -100000 || view.YWindow.Value > 100000)))
                    throw new InvalidDataException(
                        "A comparison review workbook has a startup view that can conceal its overview.");
            }
        }

        private static void ValidateStreamingSheetView(XmlReader reader)
        {
            var allowed = new HashSet<string>(new[]
            {
                "windowProtection", "showFormulas", "showGridLines", "showRowColHeaders",
                "showZeros", "rightToLeft", "tabSelected", "showRuler", "showOutlineSymbols",
                "defaultGridColor", "showWhiteSpace", "view", "topLeftCell", "colorId",
                "zoomScale", "zoomScaleNormal", "zoomScaleSheetLayoutView",
                "zoomScalePageLayoutView", "workbookViewId",
            }, StringComparer.Ordinal);
            ValidateOnlyUnqualifiedAttributes(reader, allowed, "sheet view");
            if (StreamingBooleanEnabledWhenPresent(reader.GetAttribute("windowProtection")) ||
                StreamingBooleanEnabledWhenPresent(reader.GetAttribute("showFormulas")) ||
                StreamingBooleanDisabledWhenPresent(reader.GetAttribute("showGridLines")) ||
                StreamingBooleanDisabledWhenPresent(reader.GetAttribute("showRowColHeaders")) ||
                StreamingBooleanDisabledWhenPresent(reader.GetAttribute("showZeros")) ||
                StreamingBooleanEnabledWhenPresent(reader.GetAttribute("rightToLeft")) ||
                StreamingBooleanDisabledWhenPresent(reader.GetAttribute("showOutlineSymbols")))
                throw UnsafeWorksheetPresentation();
            string topLeft = reader.GetAttribute("topLeftCell");
            if (topLeft != null && !string.Equals(topLeft, "A1", StringComparison.Ordinal))
                throw UnsafeWorksheetPresentation();
            string view = reader.GetAttribute("view");
            if (view != null && !string.Equals(view, "normal", StringComparison.Ordinal))
                throw UnsafeWorksheetPresentation();
            ValidateZoom(reader.GetAttribute("zoomScale"));
            ValidateZoom(reader.GetAttribute("zoomScaleNormal"));
            ValidateZoom(reader.GetAttribute("zoomScaleSheetLayoutView"));
            ValidateZoom(reader.GetAttribute("zoomScalePageLayoutView"));
        }

        private static void ValidateStreamingSelection(XmlReader reader)
        {
            var allowed = new HashSet<string>(
                new[] { "pane", "activeCell", "activeCellId", "sqref" },
                StringComparer.Ordinal);
            ValidateOnlyUnqualifiedAttributes(reader, allowed, "selection");
            if (reader.GetAttribute("pane") != null ||
                (reader.GetAttribute("activeCell") is string activeCell && activeCell != "A1") ||
                (reader.GetAttribute("activeCellId") is string activeCellId && activeCellId != "0") ||
                (reader.GetAttribute("sqref") is string sqref && sqref != "A1"))
                throw UnsafeWorksheetPresentation();
        }

        private static void ValidateZoom(string value)
        {
            if (value == null) return;
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int zoom) ||
                zoom < 50 || zoom > 400)
                throw UnsafeWorksheetPresentation();
        }

        private static void ValidateStreamingRowPresentation(XmlReader reader)
        {
            if (StreamingBooleanEnabledWhenPresent(reader.GetAttribute("hidden")) ||
                StreamingBooleanEnabledWhenPresent(reader.GetAttribute("collapsed")) ||
                IsNonZeroUnsigned(reader.GetAttribute("s")) ||
                IsBelow(reader.GetAttribute("ht"), MinVisibleRowHeightPoints))
                throw UnsafeWorksheetPresentation();
        }

        private static void ValidateStreamingColumnPresentation(XmlReader reader)
        {
            if (StreamingBooleanEnabledWhenPresent(reader.GetAttribute("hidden")) ||
                StreamingBooleanEnabledWhenPresent(reader.GetAttribute("collapsed")) ||
                IsNonZeroUnsigned(reader.GetAttribute("style")) ||
                IsBelow(reader.GetAttribute("width"), MinVisibleColumnWidth))
                throw UnsafeWorksheetPresentation();
        }

        private static void ValidateStreamingCellPresentation(XmlReader reader)
        {
            if (IsNonZeroUnsigned(reader.GetAttribute("s")))
                throw UnsafeWorksheetPresentation();
        }

        private static void ValidateStreamingCellAttributes(XmlReader reader)
        {
            var allowed = new HashSet<string>(new[] { "r", "t", "s" }, StringComparer.Ordinal);
            if (!reader.HasAttributes) return;
            while (reader.MoveToNextAttribute())
            {
                if (!string.IsNullOrEmpty(reader.NamespaceURI) || !allowed.Contains(reader.LocalName))
                {
                    reader.MoveToElement();
                    throw new InvalidDataException(
                        "A review workbook cell contains an unsupported attribute.");
                }
            }
            reader.MoveToElement();
        }

        private static void ValidateStreamingHyperlinkAttributes(XmlReader reader)
        {
            var allowed = new HashSet<string>(
                new[] { "ref", "location", "display", "tooltip" },
                StringComparer.Ordinal);
            const string relationshipsNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            if (!reader.HasAttributes) return;
            while (reader.MoveToNextAttribute())
            {
                bool allowedRelationshipId =
                    reader.LocalName == "id" && reader.NamespaceURI == relationshipsNamespace;
                if (!allowedRelationshipId &&
                    (!string.IsNullOrEmpty(reader.NamespaceURI) || !allowed.Contains(reader.LocalName)))
                {
                    reader.MoveToElement();
                    throw new InvalidDataException(
                        "A review workbook hyperlink contains an unsupported attribute.");
                }
            }
            reader.MoveToElement();
        }

        private static void ValidateStreamingTextAttributes(XmlReader reader)
        {
            if (!reader.HasAttributes) return;
            while (reader.MoveToNextAttribute())
            {
                if (reader.LocalName != "space" || reader.NamespaceURI != XmlNamespace ||
                    (reader.Value != "preserve" && reader.Value != "default"))
                {
                    reader.MoveToElement();
                    throw new InvalidDataException(
                        "A review workbook text element contains an unsupported attribute.");
                }
            }
            reader.MoveToElement();
        }

        private static string ReadStreamingTextElement(XmlReader reader, int maximumCharacters)
        {
            string whitespaceMode = reader.GetAttribute("space", XmlNamespace);
            ValidateStreamingTextAttributes(reader);
            string value = ReadBoundedElementText(reader, maximumCharacters);
            if (value.Length > 0 &&
                (IsXmlWhitespace(value[0]) || IsXmlWhitespace(value[value.Length - 1])) &&
                !string.Equals(whitespaceMode, "preserve", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "A review workbook text value with edge whitespace must preserve that whitespace explicitly.");
            return value;
        }

        private static bool IsXmlWhitespace(char value) =>
            value == ' ' || value == '\t' || value == '\r' || value == '\n';

        private static bool StreamingBooleanEnabledWhenPresent(string value) =>
            value != null && StreamingBooleanEnabled(value);

        private static bool StreamingBooleanDisabledWhenPresent(string value) =>
            value != null && !StreamingBooleanEnabled(value);

        private static void ValidateOnlyUnqualifiedAttributes(
            XmlReader reader,
            ISet<string> allowed,
            string label)
        {
            if (!reader.HasAttributes) return;
            while (reader.MoveToNextAttribute())
            {
                if (!string.IsNullOrEmpty(reader.NamespaceURI) || !allowed.Contains(reader.LocalName))
                {
                    reader.MoveToElement();
                    throw new InvalidDataException(
                        "A review workbook " + label + " contains an unsupported attribute.");
                }
            }
            reader.MoveToElement();
        }

        private static bool IsNonZeroUnsigned(string value)
        {
            if (value == null) return false;
            return !uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed) ||
                   parsed != 0;
        }

        private static bool IsBelow(string value, double minimum)
        {
            if (value == null) return false;
            return !double.TryParse(
                       value,
                       NumberStyles.Float,
                       CultureInfo.InvariantCulture,
                       out double parsed) || parsed < minimum;
        }

        private static InvalidDataException UnsafeWorksheetPresentation() => new InvalidDataException(
            "A comparison review workbook contains presentation changes that can conceal or relabel " +
            "protected evidence.");

        private readonly struct StreamingCellReference
        {
            public StreamingCellReference(int row, int column)
            {
                Row = row;
                Column = column;
            }

            public int Row { get; }
            public int Column { get; }
        }

        private sealed class StreamingHyperlink
        {
            public StreamingHyperlink(
                StreamingCellReference coordinate,
                string reference,
                string location,
                string display,
                string tooltip)
            {
                Coordinate = coordinate;
                Reference = reference;
                Location = location;
                Display = display;
                Tooltip = tooltip;
            }

            public StreamingCellReference Coordinate { get; }
            public string Reference { get; }
            public string Location { get; }
            public string Display { get; }
            public string Tooltip { get; }
            public string VisibleCellText { get; set; }
        }

        private sealed class StreamingWorksheetScan
        {
            public StreamingWorksheetScan(
                int cellCount,
                IReadOnlyList<StreamingHyperlink> hyperlinks)
            {
                CellCount = cellCount;
                Hyperlinks = hyperlinks;
            }

            public int CellCount { get; }
            public IReadOnlyList<StreamingHyperlink> Hyperlinks { get; }
        }

        /// <summary>
        /// A fixed-width offset index plus bounded variable-length payload spool avoids retaining millions of
        /// shared-string objects in a 32-bit Excel process. Both handles are delete-on-close; sequential index
        /// access (Excel's normal first-use ordering) avoids a seek per cell, while arbitrary valid indices
        /// remain supported with bounded random access.
        /// </summary>
        private sealed class StreamingSharedStringStore : IDisposable
        {
            private readonly int _maximumCount;
            private readonly long _maximumPayloadBytes;
            private readonly ReviewWorkbookProjectionValidationMetrics _metrics;
            private FileStream _index;
            private FileStream _payload;
            private readonly UTF8Encoding _utf8 = new UTF8Encoding(false, true);
            private readonly string _tempToken = Guid.NewGuid().ToString("N");
            private readonly Dictionary<int, StreamingSharedStringCacheEntry> _cache =
                new Dictionary<int, StreamingSharedStringCacheEntry>();
            private readonly LinkedList<int> _cacheRecency = new LinkedList<int>();
            private bool _sealed;
            private int _lastReadIndex = -1;
            private int _lastDiskReadIndex = -1;
            private int _cachedCharacters;
            private StreamingProjectionText _lastReadValue;

            public StreamingSharedStringStore(
                int maximumCount,
                long maximumPayloadBytes,
                ReviewWorkbookProjectionValidationMetrics metrics)
            {
                _maximumCount = maximumCount;
                _maximumPayloadBytes = maximumPayloadBytes;
                _metrics = metrics;
            }

            public int Count { get; private set; }

            public void Add(StreamingProjectionText value)
            {
                if (_sealed) throw new InvalidOperationException("The shared-string spool is sealed.");
                if (Count == _maximumCount)
                    throw new InvalidDataException(
                        "The review workbook shared-string table exceeds the safe item-count limit.");
                EnsureSpoolOpen();
                byte[] semantic = _utf8.GetBytes(value?.Semantic ?? "");
                byte[] visible = _utf8.GetBytes(value?.Visible ?? "");
                long recordLength = checked(8L + semantic.Length + visible.Length);
                if (_payload.Position > _maximumPayloadBytes - recordLength)
                    throw new InvalidDataException(
                        "The review workbook shared-string semantic spool exceeds its safe disk budget.");
                WriteInt64(_index, _payload.Position);
                WriteInt32(_payload, semantic.Length);
                WriteInt32(_payload, visible.Length);
                _payload.Write(semantic, 0, semantic.Length);
                _payload.Write(visible, 0, visible.Length);
                Count++;
                if (_metrics != null) _metrics.SharedStringItemCount++;
            }

            public void Seal()
            {
                if (_sealed) return;
                _index?.Flush();
                _payload?.Flush();
                _sealed = true;
                if (_index != null) _index.Position = 0;
                if (_payload != null) _payload.Position = 0;
            }

            public StreamingProjectionText Get(int index)
            {
                if (!_sealed) throw new InvalidOperationException("The shared-string spool is not sealed.");
                if (index < 0 || index >= Count)
                    throw new InvalidDataException("A review workbook shared-string index is invalid.");
                if (_metrics != null) _metrics.SharedStringLookupCount++;
                if (index == _lastReadIndex)
                {
                    if (_metrics != null) _metrics.SharedStringCacheHitCount++;
                    return _lastReadValue;
                }
                if (_cache.TryGetValue(index, out StreamingSharedStringCacheEntry cached))
                {
                    _cacheRecency.Remove(cached.Recency);
                    _cacheRecency.AddFirst(cached.Recency);
                    _lastReadIndex = index;
                    _lastReadValue = cached.Value;
                    if (_metrics != null) _metrics.SharedStringCacheHitCount++;
                    return cached.Value;
                }
                if (_index == null || _payload == null)
                    throw new InvalidDataException("The bounded shared-string spool is unavailable.");
                if (index != _lastDiskReadIndex + 1)
                {
                    _index.Position = checked((long)index * 8L);
                    _payload.Position = ReadInt64(_index);
                    if (_metrics != null) _metrics.SharedStringRandomSeekCount++;
                }
                else if (_lastDiskReadIndex < 0)
                {
                    _payload.Position = 0;
                }

                int semanticLength = ReadInt32(_payload);
                int visibleLength = ReadInt32(_payload);
                if (semanticLength < 0 || visibleLength < 0 ||
                    semanticLength > MaxStreamingSharedStringSpoolBytes ||
                    visibleLength > MaxStreamingSharedStringSpoolBytes - semanticLength)
                    throw new InvalidDataException("The bounded shared-string spool is corrupt.");
                string semantic = _utf8.GetString(ReadExactly(_payload, semanticLength));
                string visible = _utf8.GetString(ReadExactly(_payload, visibleLength));
                _lastReadIndex = index;
                _lastDiskReadIndex = index;
                _lastReadValue = new StreamingProjectionText(semantic, visible);
                if (_metrics != null) _metrics.SharedStringDiskReadCount++;
                Cache(index, _lastReadValue);
                return _lastReadValue;
            }

            public void Dispose()
            {
                _payload?.Dispose();
                _index?.Dispose();
            }

            private void EnsureSpoolOpen()
            {
                if (_index != null) return;
                string directory = Path.GetTempPath();
                _index = OpenDeleteOnCloseTemp(Path.Combine(
                    directory,
                    "ExcelPowerTools-sst-index-" + _tempToken + ".tmp"));
                try
                {
                    _payload = OpenDeleteOnCloseTemp(Path.Combine(
                        directory,
                        "ExcelPowerTools-sst-payload-" + _tempToken + ".tmp"));
                }
                catch
                {
                    _index.Dispose();
                    _index = null;
                    throw;
                }
            }

            private void Cache(int index, StreamingProjectionText value)
            {
                int characters = checked(value.Semantic.Length + (value.Visible?.Length ?? 0));
                if (characters > MaxStreamingSharedStringCacheCharacters) return;
                while (_cache.Count >= MaxStreamingSharedStringCacheEntries ||
                       _cachedCharacters > MaxStreamingSharedStringCacheCharacters - characters)
                {
                    LinkedListNode<int> oldest = _cacheRecency.Last;
                    if (oldest == null) break;
                    StreamingSharedStringCacheEntry removed = _cache[oldest.Value];
                    _cache.Remove(oldest.Value);
                    _cacheRecency.RemoveLast();
                    _cachedCharacters -= removed.Characters;
                }
                var recency = new LinkedListNode<int>(index);
                _cacheRecency.AddFirst(recency);
                _cache.Add(
                    index,
                    new StreamingSharedStringCacheEntry(value, recency, characters));
                _cachedCharacters += characters;
                if (_metrics != null)
                {
                    _metrics.PeakSharedStringCacheEntries = Math.Max(
                        _metrics.PeakSharedStringCacheEntries,
                        _cache.Count);
                    _metrics.PeakSharedStringCacheCharacters = Math.Max(
                        _metrics.PeakSharedStringCacheCharacters,
                        _cachedCharacters);
                }
            }

            private static FileStream OpenDeleteOnCloseTemp(string path) => new FileStream(
                path,
                FileMode.CreateNew,
                FileAccess.ReadWrite,
                FileShare.None,
                64 * 1024,
                FileOptions.DeleteOnClose | FileOptions.RandomAccess);

            private static void WriteInt32(Stream stream, int value)
            {
                byte[] bytes =
                {
                    (byte)value,
                    (byte)(value >> 8),
                    (byte)(value >> 16),
                    (byte)(value >> 24),
                };
                stream.Write(bytes, 0, bytes.Length);
            }

            private static void WriteInt64(Stream stream, long value)
            {
                var bytes = new byte[8];
                for (int index = 0; index < bytes.Length; index++)
                    bytes[index] = (byte)(value >> (index * 8));
                stream.Write(bytes, 0, bytes.Length);
            }

            private static int ReadInt32(Stream stream)
            {
                byte[] bytes = ReadExactly(stream, 4);
                return bytes[0] | bytes[1] << 8 | bytes[2] << 16 | bytes[3] << 24;
            }

            private static long ReadInt64(Stream stream)
            {
                byte[] bytes = ReadExactly(stream, 8);
                long value = 0;
                for (int index = 0; index < bytes.Length; index++)
                    value |= (long)bytes[index] << (index * 8);
                return value;
            }

            private static byte[] ReadExactly(Stream stream, int count)
            {
                var bytes = new byte[count];
                int offset = 0;
                while (offset < count)
                {
                    int read = stream.Read(bytes, offset, count - offset);
                    if (read == 0)
                        throw new InvalidDataException("The bounded shared-string spool is truncated.");
                    offset += read;
                }
                return bytes;
            }

            private sealed class StreamingSharedStringCacheEntry
            {
                public StreamingSharedStringCacheEntry(
                    StreamingProjectionText value,
                    LinkedListNode<int> recency,
                    int characters)
                {
                    Value = value;
                    Recency = recency;
                    Characters = characters;
                }

                public StreamingProjectionText Value { get; }
                public LinkedListNode<int> Recency { get; }
                public int Characters { get; }
            }
        }

        private sealed class StreamingProjectionText
        {
            public StreamingProjectionText(string semantic, string visible)
            {
                Semantic = semantic ?? "";
                Visible = visible;
            }

            public string Semantic { get; }
            public string Visible { get; }
        }

        private readonly struct StreamingRun
        {
            public StreamingRun(string style, string text)
            {
                Style = style;
                Text = text;
            }

            public string Style { get; }
            public string Text { get; }
        }

        private readonly struct StreamingRunStyle
        {
            public StreamingRunStyle(string bold, string strike, string rgb)
            {
                Bold = bold;
                Strike = strike;
                Rgb = rgb;
            }

            public string Bold { get; }
            public string Strike { get; }
            public string Rgb { get; }
        }
    }

    /// <summary>
    /// Deterministic work counters from projection validation. These are diagnostics, not integrity inputs;
    /// they make scale regressions observable without exposing Open XML implementation types.
    /// </summary>
    public sealed class ReviewWorkbookProjectionValidationMetrics
    {
        internal ReviewWorkbookProjectionValidationMetrics(string digestVersion)
        {
            DigestVersion = digestVersion ?? "";
        }

        public string DigestVersion { get; }
        public long WorksheetPassCount { get; internal set; }
        public long CellCount { get; internal set; }
        public long HyperlinkCount { get; internal set; }
        public long SharedStringItemCount { get; internal set; }
        public long SharedStringLookupCount { get; internal set; }
        public long SharedStringCacheHitCount { get; internal set; }
        public long SharedStringDiskReadCount { get; internal set; }
        public long SharedStringRandomSeekCount { get; internal set; }
        public int PeakSharedStringCacheEntries { get; internal set; }
        public int PeakSharedStringCacheCharacters { get; internal set; }
    }
}
