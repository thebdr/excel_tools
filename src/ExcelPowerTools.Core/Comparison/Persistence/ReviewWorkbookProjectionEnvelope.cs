using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    public static partial class ReviewWorkbookFile
    {
        private const long MaxOpcControlPartCharacters = 4L * 1024 * 1024;
        private const string ContentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string PackageRelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string OfficeDocumentRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string StrictOfficeDocumentRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
        private const string WorksheetRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string StrictWorksheetRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";
        private const string SharedStringsRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        private const string StrictSharedStringsRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings";
        private const string WorkbookContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string WorksheetContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string SharedStringsContentType =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        private const string StrictSpreadsheetNamespace =
            "http://purl.oclc.org/ooxml/spreadsheetml/main";

        /// <summary>
        /// Finds large-entry candidates from OPC content types and the package/workbook relationship graph.
        /// A filename that merely resembles xl/worksheets/sheetN.xml never earns the larger expansion tier.
        /// </summary>
        private static IReadOnlyDictionary<string, long>
            GetRelationshipBoundLargeSpreadsheetEntries(ZipArchive archive)
        {
            var entries = archive.Entries.ToDictionary(
                entry => (entry.FullName ?? "").Replace('\\', '/'),
                StringComparer.Ordinal);
            XDocument contentTypes = ReadOpcControlXml(
                RequireOpcEntry(entries, "[Content_Types].xml"),
                "content-types");
            var contentTypeMap = ParseContentTypes(contentTypes);

            XDocument packageRelationships = ReadOpcControlXml(
                RequireOpcEntry(entries, "_rels/.rels"),
                "package relationships");
            XElement[] officeDocuments = DirectRelationships(packageRelationships)
                .Where(element =>
                    IsRelationshipType(
                        element.Attribute("Type")?.Value,
                        OfficeDocumentRelationship,
                        StrictOfficeDocumentRelationship) &&
                    !IsExternalRelationship(element))
                .ToArray();
            if (officeDocuments.Length != 1)
                throw new InvalidDataException(
                    "The review workbook package has no unique internal office-document relationship.");
            XElement officeDocument = officeDocuments[0];
            string workbookName = ResolveRelationshipTarget(
                "",
                RequireOpcAttribute(officeDocument, "Target"));
            if (!entries.ContainsKey(workbookName) ||
                !string.Equals(
                    contentTypeMap.GetContentType(workbookName),
                    WorkbookContentType,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The review workbook main part has an invalid content type or relationship target.");

            string workbookRelationshipsName = RelationshipPartName(workbookName);
            XDocument workbookRelationships = ReadOpcControlXml(
                RequireOpcEntry(entries, workbookRelationshipsName),
                "workbook relationships");
            XElement[] relationships = DirectRelationships(workbookRelationships);
            var relationshipsById = new Dictionary<string, XElement>(StringComparer.Ordinal);
            foreach (XElement relationship in relationships)
            {
                string id = RequireOpcAttribute(relationship, "Id");
                if (relationshipsById.ContainsKey(id))
                    throw new InvalidDataException("A workbook relationship ID is duplicated.");
                relationshipsById.Add(id, relationship);
            }

            XDocument workbook = ReadOpcControlXml(
                RequireOpcEntry(entries, workbookName),
                "workbook");
            List<string> sheetRelationshipIds = ReadWorkbookSheetRelationshipIds(workbook);
            var largeParts = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            var referencedWorksheetIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (string relationshipId in sheetRelationshipIds)
            {
                if (!relationshipsById.TryGetValue(relationshipId, out XElement relationship) ||
                    !IsRelationshipType(
                        relationship.Attribute("Type")?.Value,
                        WorksheetRelationship,
                        StrictWorksheetRelationship))
                    throw new InvalidDataException(
                        "A workbook sheet does not reference a worksheet relationship.");
                referencedWorksheetIds.Add(relationshipId);
                AddBoundLargePart(
                    relationship,
                    workbookName,
                    entries,
                    contentTypeMap,
                    WorksheetContentType,
                    MaxStreamingWorksheetCharacters,
                    largeParts,
                    "worksheet");
            }

            XElement[] worksheetRelationships = relationships.Where(relationship =>
                IsRelationshipType(
                    relationship.Attribute("Type")?.Value,
                    WorksheetRelationship,
                    StrictWorksheetRelationship)).ToArray();
            if (worksheetRelationships.Length != referencedWorksheetIds.Count)
                throw new InvalidDataException(
                    "The workbook contains an unreferenced or duplicated worksheet relationship.");

            XElement[] sharedStringRelationships = relationships.Where(relationship =>
                IsRelationshipType(
                    relationship.Attribute("Type")?.Value,
                    SharedStringsRelationship,
                    StrictSharedStringsRelationship)).ToArray();
            if (sharedStringRelationships.Length > 1)
                throw new InvalidDataException(
                    "The workbook contains more than one shared-string relationship.");
            if (sharedStringRelationships.Length == 1)
            {
                AddBoundLargePart(
                    sharedStringRelationships[0],
                    workbookName,
                    entries,
                    contentTypeMap,
                    SharedStringsContentType,
                    MaxStreamingSharedStringCharacters,
                    largeParts,
                    "shared-string table");
            }

            return largeParts;
        }

        private static void AddBoundLargePart(
            XElement relationship,
            string workbookName,
            IReadOnlyDictionary<string, ZipArchiveEntry> entries,
            OpcContentTypeMap contentTypeMap,
            string expectedContentType,
            long limit,
            IDictionary<string, long> largeParts,
            string label)
        {
            if (!IsRelationshipType(
                    relationship.Attribute("Type")?.Value,
                    WorksheetRelationship,
                    StrictWorksheetRelationship) &&
                !IsRelationshipType(
                    relationship.Attribute("Type")?.Value,
                    SharedStringsRelationship,
                    StrictSharedStringsRelationship))
                throw new InvalidDataException(
                    "A large spreadsheet part has an unexpected relationship type.");
            if (IsExternalRelationship(relationship))
                throw new InvalidDataException(
                    "A " + label + " relationship cannot target an external resource.");

            string target = ResolveRelationshipTarget(
                workbookName,
                RequireOpcAttribute(relationship, "Target"));
            if (!entries.ContainsKey(target) ||
                !string.Equals(
                    contentTypeMap.GetContentType(target),
                    expectedContentType,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "A " + label + " relationship has a missing part or invalid content type: " + target);
            if (largeParts.ContainsKey(target))
                throw new InvalidDataException(
                    "A large spreadsheet part is referenced more than once: " + target);
            largeParts.Add(target, limit);
        }

        private static List<string> ReadWorkbookSheetRelationshipIds(XDocument workbook)
        {
            XElement root = workbook.Root;
            if (root == null || root.Name.LocalName != "workbook" ||
                (root.Name.NamespaceName != SpreadsheetNamespace &&
                 root.Name.NamespaceName != StrictSpreadsheetNamespace))
                throw new InvalidDataException("The workbook part has an invalid root element.");
            XNamespace workbookNamespace = root.Name.Namespace;
            XElement[] sheetsContainers = root.Elements()
                .Where(element => element.Name == workbookNamespace + "sheets")
                .ToArray();
            if (sheetsContainers.Length != 1)
                throw new InvalidDataException("The workbook must contain one sheets collection.");
            XElement[] sheets = sheetsContainers[0].Elements()
                .Where(element => element.Name == workbookNamespace + "sheet")
                .ToArray();
            if (sheets.Length != ExpectedSheetNames.Count)
                throw new InvalidDataException("The review workbook sheet count is invalid.");
            var ids = new List<string>(sheets.Length);
            var unique = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement sheet in sheets)
            {
                XAttribute[] relationshipIds = sheet.Attributes().Where(attribute =>
                    attribute.Name.LocalName == "id" &&
                    (attribute.Name.NamespaceName ==
                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ||
                     attribute.Name.NamespaceName ==
                     "http://purl.oclc.org/ooxml/officeDocument/relationships"))
                    .ToArray();
                if (relationshipIds.Length != 1 ||
                    string.IsNullOrWhiteSpace(relationshipIds[0].Value) ||
                    !unique.Add(relationshipIds[0].Value))
                    throw new InvalidDataException(
                        "Workbook sheet relationship IDs must be present and unique.");
                ids.Add(relationshipIds[0].Value);
            }
            return ids;
        }

        private static XElement[] DirectRelationships(XDocument document)
        {
            XNamespace ns = PackageRelationshipsNamespace;
            if (document.Root == null || document.Root.Name != ns + "Relationships")
                throw new InvalidDataException("An OPC relationships part has an invalid root element.");
            return document.Root.Elements(ns + "Relationship").ToArray();
        }

        private static bool IsRelationshipType(string value, string transitional, string strict) =>
            string.Equals(value, transitional, StringComparison.Ordinal) ||
            string.Equals(value, strict, StringComparison.Ordinal);

        private static bool IsExternalRelationship(XElement relationship) =>
            string.Equals(
                relationship.Attribute("TargetMode")?.Value,
                "External",
                StringComparison.OrdinalIgnoreCase);

        private static string RequireOpcAttribute(XElement element, string name)
        {
            string value = element.Attribute(name)?.Value;
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException("An OPC relationship is missing its " + name + " attribute.");
            return value;
        }

        private static ZipArchiveEntry RequireOpcEntry(
            IReadOnlyDictionary<string, ZipArchiveEntry> entries,
            string name)
        {
            if (!entries.TryGetValue(name, out ZipArchiveEntry entry))
                throw new InvalidDataException("The review workbook package is missing " + name + ".");
            if (entry.Length < 0 || entry.Length > MaxOpcControlPartCharacters)
                throw new InvalidDataException("An OPC control part exceeds the safe inspection limit: " + name);
            return entry;
        }

        private static XDocument ReadOpcControlXml(ZipArchiveEntry entry, string label)
        {
            try
            {
                using (Stream stream = entry.Open())
                using (XmlReader reader = CreateBoundedXmlReader(stream, MaxOpcControlPartCharacters))
                    return XDocument.Load(reader, LoadOptions.None);
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is XmlException || exception is InvalidOperationException)
            {
                throw new InvalidDataException(
                    "The review workbook " + label + " part is invalid bounded XML.",
                    exception);
            }
        }

        private static OpcContentTypeMap ParseContentTypes(XDocument document)
        {
            XNamespace ns = ContentTypesNamespace;
            if (document.Root == null || document.Root.Name != ns + "Types")
                throw new InvalidDataException("The OPC content-types part has an invalid root element.");
            var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement element in document.Root.Elements())
            {
                if (element.Name == ns + "Override")
                {
                    string partName = RequireOpcAttribute(element, "PartName");
                    if (!partName.StartsWith("/", StringComparison.Ordinal))
                        throw new InvalidDataException("An OPC content-type override has an invalid part name.");
                    string normalized = NormalizeOpcPath(partName.Substring(1));
                    if (overrides.ContainsKey(normalized))
                        throw new InvalidDataException("An OPC content-type override is duplicated.");
                    overrides.Add(normalized, RequireOpcAttribute(element, "ContentType"));
                }
                else if (element.Name == ns + "Default")
                {
                    string extension = RequireOpcAttribute(element, "Extension");
                    if (extension.IndexOfAny(new[] { '/', '\\', '.', ':', '?' }) >= 0 ||
                        defaults.ContainsKey(extension))
                        throw new InvalidDataException("An OPC default content type is invalid or duplicated.");
                    defaults.Add(extension, RequireOpcAttribute(element, "ContentType"));
                }
                else
                {
                    throw new InvalidDataException("The OPC content-types part contains an unknown element.");
                }
            }
            return new OpcContentTypeMap(overrides, defaults);
        }

        private static string RelationshipPartName(string sourcePartName)
        {
            int slash = sourcePartName.LastIndexOf('/');
            string directory = slash < 0 ? "" : sourcePartName.Substring(0, slash + 1);
            string fileName = slash < 0 ? sourcePartName : sourcePartName.Substring(slash + 1);
            return directory + "_rels/" + fileName + ".rels";
        }

        private static string ResolveRelationshipTarget(string sourcePartName, string target)
        {
            if (string.IsNullOrWhiteSpace(target) || target.IndexOf('\\') >= 0 ||
                target.IndexOf('?') >= 0 || target.IndexOf('#') >= 0 ||
                target.IndexOf(':') >= 0)
                throw new InvalidDataException("An OPC relationship target is not a safe internal part URI.");
            string combined;
            if (target.StartsWith("/", StringComparison.Ordinal))
            {
                combined = target.Substring(1);
            }
            else
            {
                int slash = sourcePartName.LastIndexOf('/');
                string directory = slash < 0 ? "" : sourcePartName.Substring(0, slash + 1);
                combined = directory + target;
            }
            return NormalizeOpcPath(combined);
        }

        private static string NormalizeOpcPath(string value)
        {
            var segments = new List<string>();
            foreach (string segment in value.Split('/'))
            {
                if (segment.Length == 0 || segment == ".") continue;
                if (segment == "..")
                {
                    if (segments.Count == 0)
                        throw new InvalidDataException("An OPC relationship target escapes the package root.");
                    segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                if (segment.IndexOf('\\') >= 0)
                    throw new InvalidDataException("An OPC part name contains an invalid separator.");
                segments.Add(segment);
            }
            if (segments.Count == 0)
                throw new InvalidDataException("An OPC relationship target resolves to an empty part name.");
            return string.Join("/", segments);
        }

        private sealed class OpcContentTypeMap
        {
            private readonly IReadOnlyDictionary<string, string> _overrides;
            private readonly IReadOnlyDictionary<string, string> _defaults;

            public OpcContentTypeMap(
                IReadOnlyDictionary<string, string> overrides,
                IReadOnlyDictionary<string, string> defaults)
            {
                _overrides = overrides;
                _defaults = defaults;
            }

            public string GetContentType(string partName)
            {
                if (_overrides.TryGetValue(partName, out string contentType)) return contentType;
                int dot = partName.LastIndexOf('.');
                string extension = dot < 0 ? "" : partName.Substring(dot + 1);
                return _defaults.TryGetValue(extension, out contentType) ? contentType : "";
            }
        }

        private sealed class ProjectionWriteBudget
        {
            private readonly long _maximum;
            private long _written;

            public ProjectionWriteBudget(long maximum)
            {
                _maximum = maximum;
            }

            public void Reserve(int count)
            {
                if (count < 0 || _written > _maximum - count)
                    throw new InvalidDataException(
                        "The generated review worksheets exceed the safe aggregate expansion budget.");
                _written += count;
            }
        }

        private sealed class BoundedWorksheetWriteStream : Stream
        {
            private readonly Stream _inner;
            private readonly long _maximum;
            private readonly ProjectionWriteBudget _budget;
            private long _written;

            public BoundedWorksheetWriteStream(
                Stream inner,
                long maximum,
                ProjectionWriteBudget budget)
            {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
                _maximum = maximum;
                _budget = budget ?? throw new ArgumentNullException(nameof(budget));
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _written;
            public override long Position
            {
                get => _written;
                set => throw new NotSupportedException();
            }

            public override void Flush() => _inner.Flush();

            public override void Write(byte[] buffer, int offset, int count)
            {
                if (count < 0 || _written > _maximum - count)
                    throw new InvalidDataException(
                        "A generated review worksheet exceeds the safe expansion budget.");
                _budget.Reserve(count);
                _inner.Write(buffer, offset, count);
                _written += count;
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) =>
                throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
        }

        private sealed class BoundedPackageWriteStream : Stream
        {
            private readonly Stream _inner;
            private readonly long _maximum;
            private readonly bool _leaveOpen;

            public BoundedPackageWriteStream(Stream inner, long maximum, bool leaveOpen)
            {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
                _maximum = maximum;
                _leaveOpen = leaveOpen;
                if (!_inner.CanSeek || !_inner.CanWrite)
                    throw new InvalidDataException(
                        "The generated review package requires a seekable bounded output stream.");
            }

            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => true;
            public override bool CanWrite => true;
            public override long Length => _inner.Length;
            public override long Position
            {
                get => _inner.Position;
                set
                {
                    if (value < 0 || value > _maximum)
                        throw new InvalidDataException(
                            "The generated review package exceeds the safe disk budget.");
                    _inner.Position = value;
                }
            }

            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) =>
                _inner.Read(buffer, offset, count);

            public override long Seek(long offset, SeekOrigin origin)
            {
                long position = _inner.Seek(offset, origin);
                if (position < 0 || position > _maximum)
                    throw new InvalidDataException(
                        "The generated review package exceeds the safe disk budget.");
                return position;
            }

            public override void SetLength(long value)
            {
                if (value < 0 || value > _maximum)
                    throw new InvalidDataException(
                        "The generated review package exceeds the safe disk budget.");
                _inner.SetLength(value);
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                if (count < 0 || Position > _maximum - count)
                    throw new InvalidDataException(
                        "The generated review package exceeds the safe disk budget.");
                _inner.Write(buffer, offset, count);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing && !_leaveOpen) _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
