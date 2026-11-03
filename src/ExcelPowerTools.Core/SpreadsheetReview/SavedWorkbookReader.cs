using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public sealed class SavedReviewWorkbook
    {
        internal SavedReviewWorkbook(string id, string path, string hash, long length, DateTime modified, bool date1904,
            IEnumerable<ReviewSheet> sheets, IEnumerable<ReviewIssue> issues, IDictionary<string, long> storedCounts)
        {
            ArtifactId = id; SourcePath = path; Sha256 = hash; ByteLength = length; LastWriteTimeUtc = modified; Date1904 = date1904;
            Sheets = Array.AsReadOnly(sheets.ToArray()); Issues = Array.AsReadOnly(issues.ToArray());
            StoredCellCounts = new ReadOnlyDictionary<string, long>(new Dictionary<string, long>(storedCounts));
        }
        public string ArtifactId { get; }
        public string SourcePath { get; }
        public string Sha256 { get; }
        public long ByteLength { get; }
        public DateTime LastWriteTimeUtc { get; }
        public bool Date1904 { get; }
        public ReadOnlyCollection<ReviewSheet> Sheets { get; }
        public ReadOnlyCollection<ReviewIssue> Issues { get; }
        public IReadOnlyDictionary<string, long> StoredCellCounts { get; }
    }

    public sealed class ReviewReadProgress
    {
        internal ReviewReadProgress(string path, string sheet, long cells) { SourcePath = path; Sheet = sheet; StoredCellsRead = cells; }
        public string SourcePath { get; }
        public string Sheet { get; }
        public long StoredCellsRead { get; }
    }

    /// <summary>
    /// Read-only saved-package reader for the first XLSX/XLSM trial. No Excel automation, recalculation, extraction,
    /// external relationship following or workbook save occurs. Unsupported evidence is explicit, never a blank fallback.
    /// </summary>
    public sealed class SavedWorkbookReader
    {
        private const long MaximumSourceBytes = 256L * 1024 * 1024;
        private const long MaximumExpandedBytes = 2L * 1024 * 1024 * 1024;
        private const long MaximumPartBytes = 512L * 1024 * 1024;
        private const long MaximumStoredCells = 10000000;
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public SavedReviewWorkbook Read(string sourcePath, string artifactId = null,
            Action<ReviewReadProgress> progress = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            string path = Path.GetFullPath(sourcePath ?? throw new ArgumentNullException(nameof(sourcePath)));
            string extension = Path.GetExtension(path);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("The first saved-package reader supports .xlsx and .xlsm. XLSB and separated-text readers are still pending; no conversion was attempted.");
            artifactId = artifactId ?? Guid.NewGuid().ToString("N");
            byte[] bytes; DateTime modified;
            // Denying concurrent writers makes the captured bytes coherent. An already writable Excel handle may require closing the source.
            using (var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (input.Length > MaximumSourceBytes) throw new InvalidDataException("Source exceeds the current 256 MiB saved-byte capture limit. No source was changed: " + path);
                modified = File.GetLastWriteTimeUtc(path);
                using (var capture = new MemoryStream((int)input.Length))
                {
                    var buffer = new byte[81920]; int read;
                    while ((read = input.Read(buffer, 0, buffer.Length)) != 0)
                    { cancellationToken.ThrowIfCancellationRequested(); capture.Write(buffer, 0, read); }
                    bytes = capture.ToArray();
                }
            }
            string hash;
            using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            using (var memory = new MemoryStream(bytes, false))
            using (var zip = new ZipArchive(memory, ZipArchiveMode.Read))
            {
                if (zip.Entries.Count > 50000 || zip.Entries.Any(e => e.Length > MaximumPartBytes) || zip.Entries.Sum(e => e.Length) > MaximumExpandedBytes)
                    throw new InvalidDataException("Workbook package exceeds the bounded reader entry/expanded-size limits. No content was truncated: " + path);
                if (zip.Entries.GroupBy(e => e.FullName, StringComparer.Ordinal).Any(g => g.Count() > 1))
                    throw new InvalidDataException("Workbook contains duplicate package parts: " + path);
                var rootRelationships = Relationships(zip, "");
                var office = rootRelationships.Values.SingleOrDefault(r => r.Type.EndsWith("/officeDocument", StringComparison.Ordinal) && !r.External);
                if (office == null) throw new InvalidDataException("A saved spreadsheet workbook part was not found: " + path);
                string workbookPath = ResolvePart("", office.Target);
                var workbook = ReadXml(zip, workbookPath, 16 * 1024 * 1024);
                if (workbook.Name != S + "workbook") throw new NotSupportedException("This trial reader does not yet support the workbook XML namespace: " + workbook.Name.NamespaceName);
                var relationships = Relationships(zip, workbookPath);
                var issues = new List<ReviewIssue>(); var seenIssues = new HashSet<string>(StringComparer.Ordinal);
                void Issue(string code, string message, string sheet, string address)
                {
                    // Aggregate repeated reader failures by sheet/code, retaining counts through affected cell states.
                    if (!seenIssues.Add(code + "\0" + sheet)) return;
                    issues.Add(ReviewIssue.Create(code, message, new Dictionary<string, string> { ["path"] = path,
                        ["artifactId"] = artifactId, ["sha256"] = hash, ["sheet"] = sheet ?? "", ["firstAddress"] = address ?? "" }));
                }
                XElement Related(string suffix) => relationships.Values.Where(r => !r.External && r.Type.EndsWith(suffix, StringComparison.Ordinal))
                    .Select(r => ReadXml(zip, ResolvePart(workbookPath, r.Target), 128 * 1024 * 1024)).SingleOrDefault();
                var stringTable = Related("/sharedStrings");
                if (stringTable != null && stringTable.Name != S + "sst") throw new InvalidDataException("Unexpected shared-string table root.");
                var strings = stringTable?.Elements(S + "si").ToArray() ?? new XElement[0];
                var styles = new SavedStyleReader(Related("/styles"), Related("/theme"));
                var sheets = new List<ReviewSheet>(); var counts = new Dictionary<string, long>(StringComparer.Ordinal);
                var sheetList = workbook.Element(S + "sheets") ?? throw new InvalidDataException("The workbook sheet list is missing.");
                foreach (var sheet in sheetList.Elements(S + "sheet"))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string name = (string)sheet.Attribute("name"), relId = (string)sheet.Attribute(R + "id");
                    if (string.IsNullOrWhiteSpace(name) || relId == null || !relationships.TryGetValue(relId, out var rel) || rel.External)
                        throw new InvalidDataException("A sheet has an invalid or external relationship in " + path);
                    if (!rel.Type.EndsWith("/worksheet", StringComparison.Ordinal))
                    { Issue("SR.READ.NON_WORKSHEET", "A non-grid sheet was not compared. Inspect its source separately.", name, null); continue; }
                    string partPath = ResolvePart(workbookPath, rel.Target);
                    var cells = new List<ReviewCell>(); var cellStyles = new List<StyleSpan>();
                    var rowStyles = new Dictionary<int, int>(); var columnStyles = new List<ColumnStyle>();
                    int row = 0, column = 0; long stored = 0;
                    using (var part = RequiredEntry(zip, partPath).Open())
                    using (var reader = XmlReader.Create(part, XmlSettings(MaximumPartBytes)))
                    {
                        reader.MoveToContent();
                        if (reader.LocalName != "worksheet" || reader.NamespaceURI != S.NamespaceName)
                            throw new InvalidDataException("Unexpected worksheet root in " + partPath);
                        while (reader.Read())
                        {
                            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != S.NamespaceName) continue;
                            if (reader.LocalName == "row")
                            {
                                int nextRow = ParseInt(reader.GetAttribute("r"), row + 1);
                                if (nextRow <= row || nextRow > 1048576) throw new InvalidDataException("Invalid/non-increasing row index in " + partPath);
                                row = nextRow; column = 0;
                                if (reader.GetAttribute("s") != null && reader.GetAttribute("customFormat") != "0") rowStyles.Add(row, ParseInt(reader.GetAttribute("s"), 0));
                            }
                            else if (reader.LocalName == "col")
                            {
                                if (reader.GetAttribute("style") != null)
                                    columnStyles.Add(new ColumnStyle { First = ParseInt(reader.GetAttribute("min"), 1), Last = ParseInt(reader.GetAttribute("max"), 1), Style = ParseInt(reader.GetAttribute("style"), 0) });
                            }
                            else if (reader.LocalName == "c")
                            {
                                stored++;
                                if (stored > MaximumStoredCells) throw new InvalidDataException("Sheet exceeds the ten-million stored-cell trial limit; no rows were truncated: " + path + " / " + name);
                                if ((stored & 4095) == 0) { cancellationToken.ThrowIfCancellationRequested(); progress?.Invoke(new ReviewReadProgress(path, name, stored)); }
                                XElement cell;
                                using (var subtree = reader.ReadSubtree()) cell = XElement.Load(subtree);
                                string address = (string)cell.Attribute("r");
                                int nextColumn = column + 1;
                                if (address != null)
                                {
                                    var coordinate = Address(address);
                                    if (coordinate.Item1 != row) throw new InvalidDataException("Cell and containing row disagree at " + name + "!" + address);
                                    nextColumn = coordinate.Item2;
                                }
                                if (row < 1 || nextColumn <= column || nextColumn > 16384) throw new InvalidDataException("Invalid/non-increasing cell coordinate in " + name);
                                column = nextColumn;
                                if (cell.Attribute("s") != null) AddStyle(cellStyles, row, column, ParseInt((string)cell.Attribute("s"), 0));
                                if (cell.Element(S + "f") == null && cell.Element(S + "v") == null && cell.Element(S + "is") == null) continue;
                                int style = StyleAt(cellStyles, rowStyles, columnStyles, row, column);
                                try
                                {
                                    XElement rich = null; var formula = cell.Element(S + "f");
                                    var scalar = Scalar(cell, strings, out rich);
                                    CellContent content;
                                    if (formula != null)
                                    {
                                        if (formula.Value.Length == 0 || (string)formula.Attribute("t") == "dataTable")
                                        {
                                            string reason = "A shared-formula follower or data-table formula needs formula expansion, which is not implemented in the trial reader. The saved result is not a substitute for its formula.";
                                            Issue("SR.READ.FORMULA_EXPANSION", reason, name, address);
                                            content = CellContent.UnexpandedExpression(formula.ToString(SaveOptions.DisableFormatting), reason,
                                                cell.Element(S + "v") == null ? ChannelEvidence<ReviewScalar>.Unavailable("No saved formula result is stored in the source package.") :
                                                ChannelEvidence<ReviewScalar>.Observed(scalar));
                                        }
                                        else content = CellContent.Expression(formula.Value, cell.Element(S + "v") == null
                                                ? ChannelEvidence<ReviewScalar>.Unavailable("No saved formula result is stored in the source package.")
                                                : ChannelEvidence<ReviewScalar>.Observed(scalar));
                                        rich = null;
                                    }
                                    else content = CellContent.Value(scalar);
                                    var observation = styles.Observe(content, style, rich);
                                    if (!observation.Interior.IsAvailable) Issue("SR.READ.INTERIOR", observation.Interior.Reason, name, address);
                                    if (!observation.Strike.IsAvailable) Issue("SR.READ.STRIKE", observation.Strike.Reason, name, address);
                                    // Empty shared/inline strings carry no content but their manual formatting is retained by the style index.
                                    if (content.HasContent || (!content.IsFormula && scalar.Kind == ReviewScalarKind.Text))
                                        cells.Add(new ReviewCell(row, column, observation));
                                }
                                catch (Exception ex) when (ex is FormatException || ex is ArgumentException || ex is InvalidDataException)
                                {
                                    string reason = "Saved-cell evidence could not be decoded: " + ex.GetType().Name + ": " + ex.Message;
                                    Issue("SR.READ.CELL", reason, name, address);
                                    cells.Add(new ReviewCell(row, column, CellObservation.Missing(CellObservationState.Failed, reason)));
                                }
                            }
                        }
                    }
                    // Captures compact style evidence only, not the source stream or entire used-range rectangle.
                    var blankCache = new Dictionary<int, CellObservation>();
                    sheets.Add(new ReviewSheet(artifactId, name, cells, (r, c) =>
                    {
                        int style = StyleAt(cellStyles, rowStyles, columnStyles, r, c);
                        lock (blankCache)
                        {
                            if (!blankCache.TryGetValue(style, out var blank))
                                blankCache.Add(style, blank = styles.Observe(CellContent.Value(ReviewScalar.Blank), style, null));
                            return blank;
                        }
                    }));
                    counts.Add(name, stored); progress?.Invoke(new ReviewReadProgress(path, name, stored));
                }
                bool date1904 = (string)workbook.Element(S + "workbookPr")?.Attribute("date1904") == "1" ||
                    (string)workbook.Element(S + "workbookPr")?.Attribute("date1904") == "true";
                return new SavedReviewWorkbook(artifactId, path, hash, bytes.Length, modified, date1904, sheets, issues, counts);
            }
        }

        private static ReviewScalar Scalar(XElement cell, XElement[] strings, out XElement rich)
        {
            rich = null; string type = (string)cell.Attribute("t"), value = (string)cell.Element(S + "v");
            switch (type)
            {
                case "s":
                    int index = ParseInt(value, -1);
                    if (index < 0 || index >= strings.Length) throw new InvalidDataException("Shared string index is outside the saved table: " + value);
                    rich = strings[index]; return ReviewScalar.Text(RichText(rich));
                case "inlineStr":
                    rich = cell.Element(S + "is");
                    if (rich == null) throw new InvalidDataException("Inline string payload is absent.");
                    return ReviewScalar.Text(RichText(rich));
                case "str": return ReviewScalar.Text(DecodeText(value ?? ""));
                case "b":
                    if (value != "0" && value != "1") throw new InvalidDataException("Saved Boolean is neither 0 nor 1.");
                    return ReviewScalar.Boolean(value == "1");
                case "e": return ReviewScalar.Error(value);
                case "d": throw new InvalidDataException("ISO-date typed cells require the date-typed reader extension; the value was not coerced into text or a serial number.");
                case null: case "n": return string.IsNullOrEmpty(value) ? ReviewScalar.Blank : ReviewScalar.Number(value);
                default: throw new InvalidDataException("Unsupported saved cell type: " + type);
            }
        }
        internal static string RichText(XElement rich) => string.Concat(rich.Elements().Where(e => e.Name == S + "t" || e.Name == S + "r")
            .Select(RichElementText));
        internal static string RichElementText(XElement element) => DecodeText(element.Name == S + "t" ? element.Value : (string)element.Element(S + "t") ?? "");
        // One pass is important: _x005F_x000D_ is literal "_x000D_", not a carriage return.
        internal static string DecodeText(string value) => Regex.Replace(value, "_x([0-9A-Fa-f]{4})_", m =>
            ((char)int.Parse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());
        private static int ParseInt(string text, int fallback) => text == null ? fallback : int.Parse(text, NumberStyles.None, CultureInfo.InvariantCulture);
        private static Tuple<int, int> Address(string address)
        {
            int col = 0, i = 0;
            while (i < address.Length && address[i] >= 'A' && address[i] <= 'Z') { col = checked(col * 26 + address[i] - 'A' + 1); i++; }
            int row = ParseInt(address.Substring(i), 0);
            if (i == 0 || col < 1 || col > 16384 || row < 1 || row > 1048576) throw new InvalidDataException("Invalid saved cell address: " + address);
            return Tuple.Create(row, col);
        }
        private sealed class Relationship { public string Type; public string Target; public bool External; }
        private static Dictionary<string, Relationship> Relationships(ZipArchive zip, string owner)
        {
            int slash = owner.LastIndexOf('/'); string prefix = slash < 0 ? "" : owner.Substring(0, slash + 1);
            string filename = slash < 0 ? owner : owner.Substring(slash + 1);
            string path = owner.Length == 0 ? "_rels/.rels" : prefix + "_rels/" + filename + ".rels";
            if (zip.GetEntry(path) == null) return new Dictionary<string, Relationship>();
            var xml = ReadXml(zip, path, 16 * 1024 * 1024);
            XNamespace package = "http://schemas.openxmlformats.org/package/2006/relationships";
            if (xml.Name != package + "Relationships") throw new InvalidDataException("Unexpected package relationship namespace/root.");
            return xml.Elements(package + "Relationship").ToDictionary(e => (string)e.Attribute("Id"), e => new Relationship
            { Type = (string)e.Attribute("Type") ?? "", Target = (string)e.Attribute("Target"), External = (string)e.Attribute("TargetMode") == "External" });
        }
        private static string ResolvePart(string owner, string target)
        {
            if (string.IsNullOrEmpty(target) || target.IndexOf('\\') >= 0 || target.IndexOf(':') >= 0 || target.IndexOf('#') >= 0)
                throw new InvalidDataException("Invalid internal package target.");
            var uri = new Uri(new Uri("https://package.invalid/" + owner), target);
            if (uri.Host != "package.invalid" || uri.Query.Length != 0) throw new InvalidDataException("External package targets cannot be followed.");
            return Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/');
        }
        private static ZipArchiveEntry RequiredEntry(ZipArchive zip, string name) => zip.GetEntry(name) ?? throw new InvalidDataException("Required saved package part is missing: " + name);
        private static XmlReaderSettings XmlSettings(long limit) => new XmlReaderSettings
        { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = limit, IgnoreComments = true };
        private static XElement ReadXml(ZipArchive zip, string name, long limit)
        {
            var entry = RequiredEntry(zip, name);
            if (entry.Length > limit) throw new InvalidDataException("Package metadata part exceeds its read limit: " + name);
            using (var stream = entry.Open()) using (var reader = XmlReader.Create(stream, XmlSettings(limit))) return XElement.Load(reader);
        }
        private struct StyleSpan { public int Row; public int First; public int Last; public int Style; }
        private struct ColumnStyle { public int First; public int Last; public int Style; }
        private static void AddStyle(List<StyleSpan> styles, int row, int col, int style)
        {
            if (styles.Count > 0)
            {
                var last = styles[styles.Count - 1];
                if (last.Row == row && last.Last + 1 == col && last.Style == style)
                { last.Last = col; styles[styles.Count - 1] = last; return; }
            }
            styles.Add(new StyleSpan { Row = row, First = col, Last = col, Style = style });
        }
        private static int StyleAt(List<StyleSpan> styles, Dictionary<int, int> rows, List<ColumnStyle> columns, int row, int col)
        {
            int lo = 0, hi = styles.Count - 1;
            while (lo <= hi)
            {
                int mid = lo + (hi - lo) / 2; var span = styles[mid];
                if (span.Row == row && col >= span.First && col <= span.Last) return span.Style;
                if (span.Row < row || (span.Row == row && span.Last < col)) lo = mid + 1; else hi = mid - 1;
            }
            if (rows.TryGetValue(row, out int style)) return style;
            for (int i = columns.Count - 1; i >= 0; i--) if (col >= columns[i].First && col <= columns[i].Last) return columns[i].Style;
            return 0;
        }
    }
}
