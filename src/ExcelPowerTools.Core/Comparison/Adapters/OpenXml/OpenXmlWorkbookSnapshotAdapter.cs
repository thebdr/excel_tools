using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Core.Comparison.Adapters.OpenXml
{
    /// <summary>
    /// Read-only OOXML workbook adapter. It opens an owned in-memory copy with editing disabled, never starts
    /// Excel, never evaluates formulas, and never resolves external package relationships.
    /// </summary>
    public static class OpenXmlWorkbookSnapshotAdapter
    {
        public const int MaxSourcePackageBytes = 256 * 1024 * 1024;
        public const int MaxPackageEntries = 10000;
        public const long MaxCharactersInPart = 128L * 1024 * 1024;
        public const int MaxTableFormulaObservations = 100000;
        public const long MaxPhysicalCellElementsPerWorkbook = 2000000L;

        private const long MaxCentralDirectoryBytes = 16L * 1024 * 1024;
        private const long MaxDeclaredUncompressedBytes = 1024L * 1024 * 1024;

        public static OpenXmlWorkbookSnapshot FromBytes(
            string revisionId,
            byte[] exactSourceBytes,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            FromBytesCore(
                revisionId,
                exactSourceBytes,
                OpenXmlFormulaMetadataLimits.Default,
                MaxPhysicalCellElementsPerWorkbook,
                cancellationToken);

        internal static OpenXmlWorkbookSnapshot FromBytesForTesting(
            string revisionId,
            byte[] exactSourceBytes,
            OpenXmlFormulaMetadataLimits metadataLimits,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            FromBytesCore(
                revisionId,
                exactSourceBytes,
                metadataLimits,
                MaxPhysicalCellElementsPerWorkbook,
                cancellationToken);

        internal static OpenXmlWorkbookSnapshot FromBytesWithCellLimitForTesting(
            string revisionId,
            byte[] exactSourceBytes,
            long maxPhysicalCellElements,
            CancellationToken cancellationToken = default(CancellationToken)) =>
            FromBytesCore(
                revisionId,
                exactSourceBytes,
                OpenXmlFormulaMetadataLimits.Default,
                maxPhysicalCellElements,
                cancellationToken);

        private static OpenXmlWorkbookSnapshot FromBytesCore(
            string revisionId,
            byte[] exactSourceBytes,
            OpenXmlFormulaMetadataLimits metadataLimits,
            long maxPhysicalCellElements,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrEmpty(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (exactSourceBytes == null) throw new ArgumentNullException(nameof(exactSourceBytes));
            if (maxPhysicalCellElements <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxPhysicalCellElements));
            if (exactSourceBytes.Length > MaxSourcePackageBytes)
                throw new InvalidDataException(
                    "The OOXML package exceeds the current safe in-memory adapter limit. " +
                    "Use the future streaming/isolated path for this artifact.");

            var ownedBytes = new byte[exactSourceBytes.Length];
            for (int offset = 0; offset < exactSourceBytes.Length; offset += 1024 * 1024)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int length = Math.Min(1024 * 1024, exactSourceBytes.Length - offset);
                Buffer.BlockCopy(exactSourceBytes, offset, ownedBytes, offset, length);
            }
            PreflightPackage(ownedBytes, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            string byteContentId = StableId.ComputeByteContentId(ownedBytes);

            // A non-writable stream plus isEditable:false is intentional defense in depth. The SDK can neither
            // repair nor save the source package, and the caller's byte array is never handed to it directly.
            using (var packageStream = new MemoryStream(ownedBytes, writable: false))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(
                       packageStream,
                       false,
                       new OpenSettings
                       {
                           AutoSave = false,
                           MaxCharactersInPart = MaxCharactersInPart,
                       }))
            {
                return ReadDocument(
                    revisionId,
                    byteContentId,
                    document,
                    metadataLimits,
                    maxPhysicalCellElements,
                    cancellationToken);
            }
        }

        private static void PreflightPackage(byte[] bytes, CancellationToken cancellationToken)
        {
            ZipEnvelope envelope = ReadZipEnvelope(bytes);
            if (envelope.EntryCount > MaxPackageEntries ||
                envelope.CentralDirectorySize > MaxCentralDirectoryBytes)
                throw new InvalidDataException(
                    "The OOXML package central directory exceeds the safe adapter limit.");

            using (var stream = new MemoryStream(bytes, writable: false))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false))
            {
                if (archive.Entries.Count != envelope.EntryCount)
                    throw new InvalidDataException("The OOXML package entry count is inconsistent.");

                long totalLength = 0;
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (entry.Length < 0 ||
                        entry.Length > MaxDeclaredUncompressedBytes - totalLength)
                        throw new InvalidDataException(
                            "The OOXML package declares an excessive uncompressed size.");
                    totalLength += entry.Length;

                    string name = entry.FullName.Replace('\\', '/');
                    bool xmlPart = name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                                   name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
                    if (xmlPart && entry.Length > MaxCharactersInPart)
                        throw new InvalidDataException(
                            "An OOXML XML part exceeds the safe adapter limit: " + name);
                    if (xmlPart && entry.Length > 1024 * 1024 && entry.CompressedLength > 0 &&
                        entry.Length / entry.CompressedLength > 10000)
                        throw new InvalidDataException(
                            "An OOXML XML part has an unsafe compression ratio: " + name);
                }
            }
        }

        private static ZipEnvelope ReadZipEnvelope(byte[] bytes)
        {
            const uint eocdSignature = 0x06054B50U;
            int minimum = Math.Max(0, bytes.Length - 65557);
            for (int offset = bytes.Length - 22; offset >= minimum; offset--)
            {
                if (ReadUInt32(bytes, offset) != eocdSignature) continue;
                ushort commentLength = ReadUInt16(bytes, offset + 20);
                if ((long)offset + 22 + commentLength != bytes.Length) continue;
                if (ReadUInt16(bytes, offset + 4) != 0 || ReadUInt16(bytes, offset + 6) != 0)
                    throw new InvalidDataException("Multi-disk OOXML ZIP packages are unsupported.");

                ushort entriesOnDisk = ReadUInt16(bytes, offset + 8);
                ushort entryCount = ReadUInt16(bytes, offset + 10);
                uint centralDirectorySize = ReadUInt32(bytes, offset + 12);
                uint centralDirectoryOffset = ReadUInt32(bytes, offset + 16);
                if (entriesOnDisk == ushort.MaxValue || entryCount == ushort.MaxValue ||
                    centralDirectorySize == uint.MaxValue || centralDirectoryOffset == uint.MaxValue)
                    throw new InvalidDataException(
                        "ZIP64 OOXML packages exceed the current in-memory adapter envelope.");
                if (entriesOnDisk != entryCount ||
                    (long)centralDirectoryOffset + centralDirectorySize > offset)
                    throw new InvalidDataException("The OOXML ZIP central directory is inconsistent.");
                return new ZipEnvelope(entryCount, centralDirectorySize);
            }

            throw new InvalidDataException("The source bytes are not a complete OOXML ZIP package.");
        }

        private static ushort ReadUInt16(byte[] bytes, int offset)
        {
            if (offset < 0 || offset > bytes.Length - 2)
                throw new InvalidDataException("The OOXML ZIP header is truncated.");
            return (ushort)(bytes[offset] | bytes[offset + 1] << 8);
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            if (offset < 0 || offset > bytes.Length - 4)
                return 0;
            return (uint)(bytes[offset] |
                          bytes[offset + 1] << 8 |
                          bytes[offset + 2] << 16 |
                          bytes[offset + 3] << 24);
        }

        private static OpenXmlWorkbookSnapshot ReadDocument(
            string revisionId,
            string byteContentId,
            SpreadsheetDocument document,
            OpenXmlFormulaMetadataLimits metadataLimits,
            long maxPhysicalCellElements,
            CancellationToken cancellationToken)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            if (workbookPart == null || workbookPart.Workbook == null)
                throw new InvalidDataException("The OOXML package does not contain a readable workbook part.");

            ExcelDateSystem dateSystem = ReadDateSystem(workbookPart.Workbook);
            string[] sharedStrings = ReadSharedStrings(workbookPart, cancellationToken);
            IList<DateStyleKind> dateStyles = ReadDateStyles(workbookPart, cancellationToken);
            var formulaMetadata = new OpenXmlFormulaMetadataCatalog(workbookPart, metadataLimits);
            var workbookCoverage = new List<OpenXmlCoverageMessage>();
            var sheets = new List<OpenXmlWorksheetSnapshot>();
            var cellReadBudget = new WorkbookCellReadBudget(maxPhysicalCellElements);

            InventoryUnmodeledWorkbookMetadata(workbookPart, workbookCoverage);

            bool hasVbaProject = workbookPart.VbaProjectPart != null;
            if (hasVbaProject)
            {
                workbookCoverage.Add(Coverage(
                    OpenXmlCoverageCodes.VbaProjectNotInspected,
                    "The package contains a VBA project. Its exact bytes are covered by ByteContentId, but the " +
                    "formula/value adapter did not inspect or execute it.",
                    true,
                    affectsExactCellComparison: false));
            }

            InventoryUnsupportedPackageContent(document, workbookCoverage);

            if (workbookPart.ExternalRelationships.Any())
            {
                workbookCoverage.Add(Coverage(
                    OpenXmlCoverageCodes.ExternalRelationshipsIgnored,
                    "One or more external workbook relationships were deliberately not followed.",
                    true,
                    affectsExactCellComparison: false));
            }

            Sheets workbookSheets = workbookPart.Workbook.Sheets;
            // Budgets are consumed once per logical <sheet> read occurrence, not once per physical part. Multiple
            // declarations may legally point at the same WorksheetPart and each snapshot retains its own strings.
            int worksheetScopeCount = workbookSheets == null
                ? 1
                : Math.Max(1, workbookSheets.Elements<Sheet>().Count());
            if (workbookSheets == null)
            {
                workbookCoverage.Add(Coverage(
                    OpenXmlCoverageCodes.MissingSheetsCollection,
                    "The workbook has no sheets collection.",
                    true));
            }
            else
            {
                int occurrence = 0;
                foreach (Sheet sheet in workbookSheets.Elements<Sheet>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    occurrence++;
                    sheets.Add(ReadSheet(
                        revisionId,
                        occurrence,
                        workbookPart,
                        sheet,
                        sharedStrings,
                        dateStyles,
                        formulaMetadata,
                        metadataLimits,
                        worksheetScopeCount,
                        dateSystem,
                        cellReadBudget,
                        cancellationToken));
                }
            }

            return new OpenXmlWorkbookSnapshot(
                revisionId,
                byteContentId,
                dateSystem,
                hasVbaProject,
                sheets,
                workbookCoverage);
        }

        private static void InventoryUnmodeledWorkbookMetadata(
            WorkbookPart workbookPart,
            IList<OpenXmlCoverageMessage> coverage)
        {
            Workbook workbook = workbookPart.Workbook;
            string partUri = workbookPart.Uri == null ? "/xl/workbook.xml" : workbookPart.Uri.ToString();

            foreach (OpenXmlAttribute attribute in workbook.GetAttributes())
            {
                AddUnmodeledWorkbookMetadata(
                    coverage,
                    partUri,
                    "Workbook attribute '" + QualifiedAttributeName(attribute) + "' is not interpreted.");
            }

            foreach (OpenXmlElement child in workbook.ChildElements)
            {
                var properties = child as WorkbookProperties;
                if (properties != null)
                {
                    foreach (OpenXmlAttribute attribute in properties.GetAttributes())
                    {
                        if (string.Equals(attribute.LocalName, "date1904", StringComparison.Ordinal) &&
                            string.IsNullOrEmpty(attribute.NamespaceUri))
                            continue;

                        AddUnmodeledWorkbookMetadata(
                            coverage,
                            partUri,
                            "Workbook property '" + QualifiedAttributeName(attribute) + "' is not interpreted.");
                    }

                    if (properties.ChildElements.Count > 0)
                    {
                        AddUnmodeledWorkbookMetadata(
                            coverage,
                            partUri,
                            "Workbook properties contain unmodeled child content.");
                    }
                    continue;
                }

                var sheets = child as Sheets;
                if (sheets != null)
                {
                    InventorySheetDeclarationMetadata(sheets, coverage, partUri);
                    continue;
                }

                AddUnmodeledWorkbookMetadata(
                    coverage,
                    partUri,
                    "Workbook element '" + child.LocalName + "' is not interpreted by the formula/value adapter.");
            }
        }

        private static void InventorySheetDeclarationMetadata(
            Sheets sheets,
            IList<OpenXmlCoverageMessage> coverage,
            string partUri)
        {
            foreach (OpenXmlAttribute attribute in sheets.GetAttributes())
            {
                AddUnmodeledWorkbookMetadata(
                    coverage,
                    partUri,
                    "Sheets attribute '" + QualifiedAttributeName(attribute) + "' is not interpreted.");
            }

            foreach (OpenXmlElement child in sheets.ChildElements)
            {
                var sheet = child as Sheet;
                if (sheet == null)
                {
                    AddUnmodeledWorkbookMetadata(
                        coverage,
                        partUri,
                        "Sheets child element '" + child.LocalName + "' is not interpreted.");
                    continue;
                }

                foreach (OpenXmlAttribute attribute in sheet.GetAttributes())
                {
                    string localName = attribute.LocalName ?? "";
                    if (localName == "name" || localName == "sheetId" ||
                        localName == "state" || localName == "id")
                        continue;

                    AddUnmodeledWorkbookMetadata(
                        coverage,
                        partUri,
                        "Sheet declaration attribute '" + QualifiedAttributeName(attribute) + "' is not interpreted.");
                }

                if (sheet.ChildElements.Count > 0)
                {
                    AddUnmodeledWorkbookMetadata(
                        coverage,
                        partUri,
                        "A sheet declaration contains unmodeled child content.");
                }
            }
        }

        private static void AddUnmodeledWorkbookMetadata(
            IList<OpenXmlCoverageMessage> coverage,
            string partUri,
            string message)
        {
            coverage.Add(Coverage(
                OpenXmlCoverageCodes.UnmodeledWorkbookMetadata,
                message + " Stored cell payloads remain available, but broad artifact equality is unsafe.",
                true,
                partUri: partUri,
                affectsExactCellComparison: false));
        }

        private static string QualifiedAttributeName(OpenXmlAttribute attribute) =>
            string.IsNullOrEmpty(attribute.Prefix)
                ? attribute.LocalName ?? ""
                : attribute.Prefix + ":" + (attribute.LocalName ?? "");

        private static void InventoryUnsupportedPackageContent(
            SpreadsheetDocument document,
            IList<OpenXmlCoverageMessage> coverage)
        {
            var supportedPartTypes = new HashSet<string>(StringComparer.Ordinal)
            {
                "WorkbookPart",
                "WorkbookStylesPart",
                "SharedStringTablePart",
                "WorksheetPart",
            };
            if (document.ExternalRelationships.Any())
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.PackageExternalRelationshipIgnored,
                    "One or more package-level external relationships were deliberately not followed.",
                    true,
                    affectsExactCellComparison: false));
            }

            var pending = new Queue<OpenXmlPart>();
            var visited = new HashSet<OpenXmlPart>();
            foreach (IdPartPair root in document.Parts)
                pending.Enqueue(root.OpenXmlPart);

            while (pending.Count > 0)
            {
                OpenXmlPart part = pending.Dequeue();
                if (!visited.Add(part)) continue;
                string partType = part.GetType().Name;
                if (!supportedPartTypes.Contains(partType) && partType != "VbaProjectPart")
                {
                    coverage.Add(Coverage(
                        OpenXmlCoverageCodes.UnsupportedPackagePart,
                        "Package part '" + partType +
                        "' is retained by the byte content ID but is not interpreted by the formula/value adapter.",
                        true,
                        partUri: part.Uri?.ToString(),
                        affectsExactCellComparison: false));
                }

                if (!(part is WorkbookPart) && !(part is WorksheetPart) &&
                    (part.ExternalRelationships.Any() || part.HyperlinkRelationships.Any()))
                {
                    coverage.Add(Coverage(
                        OpenXmlCoverageCodes.PackageExternalRelationshipIgnored,
                        "An external relationship on package part '" + partType +
                        "' was deliberately not followed.",
                        true,
                        partUri: part.Uri?.ToString(),
                        affectsExactCellComparison: false));
                }

                foreach (IdPartPair child in part.Parts)
                    pending.Enqueue(child.OpenXmlPart);
            }
        }

        private static OpenXmlWorksheetSnapshot ReadSheet(
            string revisionId,
            int occurrence,
            WorkbookPart workbookPart,
            Sheet sheet,
            string[] sharedStrings,
            IList<DateStyleKind> dateStyles,
            OpenXmlFormulaMetadataCatalog formulaMetadata,
            OpenXmlFormulaMetadataLimits formulaLimits,
            int worksheetScopeCount,
            ExcelDateSystem dateSystem,
            WorkbookCellReadBudget cellReadBudget,
            CancellationToken cancellationToken)
        {
            string name = sheet.Name == null ? "" : sheet.Name.Value;
            uint? sheetId = sheet.SheetId == null ? (uint?)null : sheet.SheetId.Value;
            string relationshipId = sheet.Id == null ? "" : sheet.Id.Value;
            OpenXmlSheetVisibility visibility = ReadVisibility(sheet, out bool visibilityKnown);
            var coverage = new List<OpenXmlCoverageMessage>();

            if (string.IsNullOrEmpty(name))
                coverage.Add(Coverage(OpenXmlCoverageCodes.MissingSheetName, "The sheet has no name.", true, name));
            if (!sheetId.HasValue)
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.MissingSheetId,
                    "The sheet has no sheetId.",
                    true,
                    name,
                    affectsExactCellComparison: false));
            if (!visibilityKnown)
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.UnknownSheetVisibility,
                    "The sheet uses an unknown visibility state.",
                    true,
                    name,
                    affectsExactCellComparison: false));

            string snapshotRevisionId = revisionId + "/sheet:" +
                                        (string.IsNullOrEmpty(relationshipId)
                                            ? occurrence.ToString(CultureInfo.InvariantCulture)
                                            : relationshipId);

            if (string.IsNullOrEmpty(relationshipId))
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.MissingSheetRelationship,
                    "The sheet has no package relationship ID.",
                    true,
                    name));
                return EmptySheet(
                    name, sheetId, relationshipId, "", null, visibility, snapshotRevisionId, coverage);
            }

            // External package relationships are identified by ID and rejected before GetPartById is called.
            if (workbookPart.ExternalRelationships.Any(
                    relationship => string.Equals(relationship.Id, relationshipId, StringComparison.Ordinal)))
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.ExternalSheetRelationshipIgnored,
                    "The sheet relationship is external and was deliberately not followed.",
                    true,
                    name));
                return EmptySheet(
                    name, sheetId, relationshipId, "", null, visibility, snapshotRevisionId, coverage);
            }

            OpenXmlPart part;
            try
            {
                part = workbookPart.GetPartById(relationshipId);
            }
            catch (Exception exception) when (
                exception is ArgumentOutOfRangeException ||
                exception is KeyNotFoundException ||
                exception is InvalidOperationException)
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.UnresolvedSheetRelationship,
                    "The internal sheet relationship could not be resolved: " + exception.Message,
                    true,
                    name));
                return EmptySheet(
                    name, sheetId, relationshipId, "", null, visibility, snapshotRevisionId, coverage);
            }

            string partUri = part.Uri == null ? "" : part.Uri.ToString();
            var worksheetPart = part as WorksheetPart;
            if (worksheetPart == null || worksheetPart.Worksheet == null)
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.UnsupportedPhysicalSheet,
                    "The sheet is not a worksheet and has no formula/value grid in this adapter.",
                    true,
                    name,
                    null,
                    partUri));
                return new OpenXmlWorksheetSnapshot(
                    name,
                    sheetId,
                    relationshipId,
                    partUri,
                    null,
                    visibility,
                    OpenXmlPhysicalSheetKind.Unsupported,
                    new RevisionSnapshot(snapshotRevisionId, null),
                    null,
                    null,
                    coverage);
            }

            string codeName = worksheetPart.Worksheet.SheetProperties == null ||
                              worksheetPart.Worksheet.SheetProperties.CodeName == null
                ? null
                : worksheetPart.Worksheet.SheetProperties.CodeName.Value;

            if (worksheetPart.ExternalRelationships.Any() || worksheetPart.HyperlinkRelationships.Any())
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.ExternalRelationshipsIgnored,
                    "One or more external worksheet relationships were deliberately not followed.",
                    true,
                    name,
                    null,
                    partUri,
                    affectsExactCellComparison: false));
            }

            var cells = new List<CanonicalCellState>();
            var formulaObservations = new List<OpenXmlFormulaObservation>();
            var formulaEvidenceCache = new Dictionary<string, string>(StringComparer.Ordinal);
            var cellCoverageLimiter = new CellCoverageLimiter(coverage, name, partUri);
            var observedReferences = new HashSet<string>(StringComparer.Ordinal);
            var readableCells = new List<KeyValuePair<string, Cell>>();
            List<SheetData> sheetDataParts = worksheetPart.Worksheet.Elements<SheetData>().ToList();
            if (sheetDataParts.Count == 0)
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.MissingSheetData,
                    "The worksheet has no sheetData element, so an empty-cell conclusion would be unsafe.",
                    true,
                    name,
                    null,
                    partUri));
            }
            else
            {
                if (sheetDataParts.Count > 1)
                {
                    coverage.Add(Coverage(
                        OpenXmlCoverageCodes.MultipleSheetData,
                        "The worksheet has more than one sheetData element; all visible cells were retained but " +
                        "exact equality is blocked.",
                        true,
                        name,
                        null,
                        partUri));
                }

                foreach (Cell cell in sheetDataParts.SelectMany(data => data.Descendants<Cell>()))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    cellReadBudget.Observe();
                    string rawReference = GetAttributeValue(cell, "r");
                    string cellReference;
                    if (!TryNormalizeCellReference(rawReference, out cellReference))
                    {
                        cellCoverageLimiter.Add(
                            OpenXmlCoverageCodes.InvalidCellReference,
                            "A cell has a missing, malformed, or out-of-grid reference: '" +
                            Abbreviate(rawReference) + "'.",
                            rawReference);
                        continue;
                    }

                    if (!observedReferences.Add(cellReference))
                    {
                        cellCoverageLimiter.Add(
                            OpenXmlCoverageCodes.DuplicateCellReference,
                            "The worksheet contains more than one physical cell at " + cellReference + ".",
                            cellReference);
                        continue;
                    }

                    readableCells.Add(new KeyValuePair<string, Cell>(cellReference, cell));
                }
            }

            OpenXmlSharedFormulaCatalog sharedFormulas =
                OpenXmlSharedFormulaCatalog.Create(
                    readableCells,
                    formulaLimits.MaxSharedFormulaDerivedCharacters / worksheetScopeCount);
            formulaMetadata.PrepareScope(
                readableCells,
                formulaLimits.MaxDerivedRecords / worksheetScopeCount,
                formulaLimits.MaxDerivedEvidenceCharacters / worksheetScopeCount / 2L);
            foreach (KeyValuePair<string, Cell> pair in readableCells)
            {
                cells.Add(ReadCell(
                    pair.Value,
                    pair.Key,
                    name,
                    partUri,
                    sharedStrings,
                    dateStyles,
                    formulaMetadata,
                    sharedFormulas,
                    dateSystem,
                    formulaEvidenceCache,
                    cellCoverageLimiter,
                    coverage,
                    out OpenXmlFormulaObservation formulaObservation));
                if (formulaObservation != null) formulaObservations.Add(formulaObservation);
            }
            cellCoverageLimiter.FlushSummaries();

            IList<OpenXmlTableFormulaObservation> tableFormulaObservations =
                ReadTableFormulaObservations(
                    worksheetPart,
                    name,
                    formulaLimits.MaxTableFormulaRecords / worksheetScopeCount,
                    formulaLimits.MaxTableFormulaEvidenceCharacters / worksheetScopeCount,
                    coverage,
                    cancellationToken);

            return new OpenXmlWorksheetSnapshot(
                name,
                sheetId,
                relationshipId,
                partUri,
                codeName,
                visibility,
                OpenXmlPhysicalSheetKind.Worksheet,
                new RevisionSnapshot(snapshotRevisionId, cells),
                formulaObservations,
                tableFormulaObservations,
                coverage);
        }

        private static IList<OpenXmlTableFormulaObservation> ReadTableFormulaObservations(
            WorksheetPart worksheetPart,
            string sheetName,
            long recordLimit,
            long evidenceCharacterLimit,
            IList<OpenXmlCoverageMessage> coverage,
            CancellationToken cancellationToken)
        {
            var observations = new List<OpenXmlTableFormulaObservation>();
            long retainedRecords = 0;
            long retainedCharacters = 0;
            bool budgetExhausted = false;
            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts
                         .OrderBy(
                             part => part.Uri == null ? "" : part.Uri.ToString(),
                             StringComparer.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string partUri = tablePart.Uri == null ? "" : tablePart.Uri.ToString();
                Table table = tablePart.Table;
                if (table == null)
                {
                    coverage.Add(Coverage(
                        OpenXmlCoverageCodes.TableFormulaPayloadMalformed,
                        "A table-definition part has no readable table root, so formula coverage is incomplete.",
                        true,
                        sheetName,
                        partUri: partUri));
                    continue;
                }

                int beforeCount = observations.Count;
                string tableName = table.DisplayName == null
                    ? (table.Name == null ? "" : table.Name.Value ?? "")
                    : table.DisplayName.Value ?? "";
                string tableRange = table.Reference == null ? "" : table.Reference.Value ?? "";
                if (!TryCanonicalizeElementAttributes(
                        table,
                        "table",
                        evidenceCharacterLimit - retainedCharacters,
                        out string tableMetadata))
                {
                    budgetExhausted = true;
                    break;
                }
                bool tableMetadataCharged = false;
                TableColumns columns = table.TableColumns;
                if (columns != null)
                {
                    foreach (TableColumn column in columns.Elements<TableColumn>())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (!TryCanonicalizeElementAttributes(
                                column,
                                "column",
                                evidenceCharacterLimit - retainedCharacters,
                                out string columnMetadata))
                        {
                            budgetExhausted = true;
                            break;
                        }
                        bool columnMetadataCharged = false;
                        if (!TryAddTableFormulaObservation(
                            observations,
                            column,
                            column.CalculatedColumnFormula,
                            OpenXmlTableFormulaKind.CalculatedColumn,
                            tableName,
                            tableRange,
                            tableMetadata,
                            columnMetadata,
                            ref tableMetadataCharged,
                            ref columnMetadataCharged,
                            ref retainedRecords,
                            ref retainedCharacters,
                            recordLimit,
                            evidenceCharacterLimit,
                            partUri,
                            sheetName,
                            coverage))
                        {
                            budgetExhausted = true;
                            break;
                        }
                        if (!TryAddTableFormulaObservation(
                            observations,
                            column,
                            column.TotalsRowFormula,
                            OpenXmlTableFormulaKind.TotalsRow,
                            tableName,
                            tableRange,
                            tableMetadata,
                            columnMetadata,
                            ref tableMetadataCharged,
                            ref columnMetadataCharged,
                            ref retainedRecords,
                            ref retainedCharacters,
                            recordLimit,
                            evidenceCharacterLimit,
                            partUri,
                            sheetName,
                            coverage))
                        {
                            budgetExhausted = true;
                            break;
                        }
                    }
                }

                if (budgetExhausted) break;

                if (observations.Count > beforeCount)
                {
                    coverage.Add(Coverage(
                        OpenXmlCoverageCodes.TableFormulaNotProjected,
                        "Exact calculated-column and totals-row formulas are inventoried from the table " +
                        "definition, but they are not manufactured as formulas on worksheet cells.",
                        true,
                        sheetName,
                        partUri: partUri));
                }
            }

            if (budgetExhausted)
            {
                observations.Clear();
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.TableFormulaPayloadMalformed,
                    "The workbook-partitioned table-formula record/character budget was exhausted. Table " +
                    "formula observations for this worksheet were discarded as a whole so part or sheet order " +
                    "cannot select partial evidence.",
                    true,
                    sheetName));
            }

            return observations;
        }

        private static bool TryAddTableFormulaObservation(
            IList<OpenXmlTableFormulaObservation> observations,
            TableColumn column,
            OpenXmlLeafTextElement formula,
            OpenXmlTableFormulaKind kind,
            string tableName,
            string tableRange,
            string tableMetadata,
            string columnMetadata,
            ref bool tableMetadataCharged,
            ref bool columnMetadataCharged,
            ref long retainedRecords,
            ref long retainedCharacters,
            long recordLimit,
            long evidenceCharacterLimit,
            string partUri,
            string sheetName,
            IList<OpenXmlCoverageMessage> coverage)
        {
            if (formula == null) return true;

            string exactText = formula.Text ?? "";
            if (!TryCanonicalizeElementAttributes(
                    formula,
                    "formula",
                    evidenceCharacterLimit - retainedCharacters,
                    out string formulaMetadata))
                return false;
            long addedCharacters = exactText.Length + (formulaMetadata == null ? 0 : formulaMetadata.Length);
            if (!tableMetadataCharged) addedCharacters += tableMetadata == null ? 0 : tableMetadata.Length;
            if (!columnMetadataCharged) addedCharacters += columnMetadata == null ? 0 : columnMetadata.Length;
            if (retainedRecords >= recordLimit ||
                addedCharacters > evidenceCharacterLimit - retainedCharacters)
                return false;

            if (string.IsNullOrEmpty(exactText))
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.TableFormulaPayloadMalformed,
                    "A table formula has no stored formula text.",
                    true,
                    sheetName,
                    partUri: partUri));
            }
            else if (OpenXmlFormulaReferenceClassifier.HasExternalWorkbookReference(exactText))
            {
                coverage.Add(Coverage(
                    OpenXmlCoverageCodes.ExternalFormulaDependencyNotResolved,
                    "The exact external-link table formula is retained, but no external workbook dependency " +
                    "was opened or resolved.",
                    true,
                    sheetName,
                    partUri: partUri,
                    affectsExactCellComparison: false));
            }

            observations.Add(new OpenXmlTableFormulaObservation(
                tableName,
                tableRange,
                column.Id == null ? (uint?)null : column.Id.Value,
                column.Name == null ? "" : column.Name.Value ?? "",
                kind,
                exactText,
                tableMetadata,
                columnMetadata,
                formulaMetadata ?? columnMetadata ?? tableMetadata,
                partUri));
            retainedRecords++;
            retainedCharacters += addedCharacters;
            tableMetadataCharged = true;
            columnMetadataCharged = true;
            return true;
        }

        private static bool TryCanonicalizeElementAttributes(
            OpenXmlElement element,
            string role,
            long remainingCharacters,
            out string metadata)
        {
            metadata = null;
            List<OpenXmlAttribute> attributes = element.GetAttributes()
                .OrderBy(item => item.NamespaceUri ?? "", StringComparer.Ordinal)
                .ThenBy(item => item.LocalName ?? "", StringComparer.Ordinal)
                .ToList();
            long required = 0;
            foreach (OpenXmlAttribute attribute in attributes)
            {
                required += FormulaTokenLength(role) + FormulaTokenLength(attribute.NamespaceUri) +
                            FormulaTokenLength(attribute.LocalName) + FormulaTokenLength(attribute.Value);
                if (required > remainingCharacters || required > int.MaxValue) return false;
            }
            if (required == 0) return true;

            var result = new StringBuilder((int)required);
            foreach (OpenXmlAttribute attribute in attributes)
            {
                AppendFormulaToken(result, role);
                AppendFormulaToken(result, attribute.NamespaceUri);
                AppendFormulaToken(result, attribute.LocalName);
                AppendFormulaToken(result, attribute.Value);
            }
            metadata = result.ToString();
            return true;
        }

        private static CanonicalCellState ReadCell(
            Cell cell,
            string cellReference,
            string sheetName,
            string partUri,
            string[] sharedStrings,
            IList<DateStyleKind> dateStyles,
            OpenXmlFormulaMetadataCatalog formulaMetadata,
            OpenXmlSharedFormulaCatalog sharedFormulas,
            ExcelDateSystem dateSystem,
            IDictionary<string, string> formulaEvidenceCache,
            CellCoverageLimiter cellCoverageLimiter,
            IList<OpenXmlCoverageMessage> coverage,
            out OpenXmlFormulaObservation formulaObservation)
        {
            formulaObservation = null;
            int formulaCount = cell.Elements<CellFormula>().Take(2).Count();
            int valueCount = cell.Elements<CellValue>().Take(2).Count();
            int inlineStringCount = cell.Elements<InlineString>().Take(2).Count();
            if (formulaCount > 1 || valueCount > 1 || inlineStringCount > 1 ||
                (inlineStringCount > 0 && (valueCount > 0 || formulaCount > 0)))
            {
                return CanonicalCellState.Constant(
                    cellReference,
                    CoverageValue(
                        OpenXmlCoverageCodes.MalformedCellPayload,
                        "The cell has duplicate or incompatible formula/value payload elements.",
                        cellReference,
                        cell,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage));
            }

            CellFormula formula = cell.CellFormula;
            if (formula == null)
            {
                string rawCellMetadataIndex = GetAttributeValue(cell, "cm");
                string rawValueMetadataIndex = GetAttributeValue(cell, "vm");
                bool hasCellExtensions = cell.ExtensionList != null;
                if (!string.IsNullOrEmpty(rawCellMetadataIndex) ||
                    !string.IsNullOrEmpty(rawValueMetadataIndex) ||
                    hasCellExtensions)
                {
                    cellCoverageLimiter.Add(
                        OpenXmlCoverageCodes.ValueMetadataNotInterpreted,
                        "The constant cell carries value/cell metadata that this adapter does not interpret " +
                        "(cm='" + Abbreviate(rawCellMetadataIndex) + "', vm='" +
                        Abbreviate(rawValueMetadataIndex) + "', extLst=" +
                        (hasCellExtensions ? "present" : "absent") +
                        "). The normal typed value is retained, but exact equality is indeterminate.",
                        cellReference);
                }

                CanonicalValue constant = ReadConstantValue(
                    cell,
                    cellReference,
                    sheetName,
                    partUri,
                    sharedStrings,
                    dateStyles,
                    dateSystem,
                    cellCoverageLimiter,
                    coverage);
                return constant.Kind == CanonicalValueKind.Blank
                    ? CanonicalCellState.Blank(cellReference)
                    : CanonicalCellState.Constant(cellReference, constant);
            }

            string formulaType = GetAttributeValue(formula, "t");
            if (string.IsNullOrEmpty(formulaType)) formulaType = "normal";
            string storedFormulaText = formula.Text ?? "";
            string effectiveFormulaText = storedFormulaText;
            string formulaCoverageCode = null;
            string formulaCoverageMessage = null;
            string formulaRange = GetAttributeValue(formula, "ref");
            uint? sharedIndex = TryReadUInt32Attribute(formula, "si");
            bool wasReconstructed = false;
            bool effectiveFormulaProven = true;
            OpenXmlFormulaForm formulaForm;
            OpenXmlSharedFormulaResolution sharedResolution = null;
            OpenXmlFormulaMetadataResolution metadataResolution =
                formulaMetadata.Resolve(cell, cellReference);

            if (metadataResolution.FailureReason != null)
            {
                cellCoverageLimiter.Add(
                    OpenXmlCoverageCodes.FormulaMetadataUnresolved,
                    metadataResolution.FailureReason,
                    cellReference);
            }

            switch (formulaType)
            {
                case "normal":
                    formulaForm = metadataResolution.IsDynamicArray
                        ? OpenXmlFormulaForm.DynamicArray
                        : OpenXmlFormulaForm.Normal;
                    break;
                case "shared":
                    formulaForm = !string.IsNullOrEmpty(storedFormulaText) &&
                                  !string.IsNullOrEmpty(formulaRange)
                        ? OpenXmlFormulaForm.SharedMaster
                        : OpenXmlFormulaForm.SharedFollower;
                    if (!sharedFormulas.TryGet(cellReference, out sharedResolution) ||
                        !sharedResolution.IsResolved)
                    {
                        formulaCoverageCode = OpenXmlCoverageCodes.SharedFormulaNotExpanded;
                        formulaCoverageMessage = sharedResolution == null
                            ? "The shared formula has no deterministic group resolution."
                            : sharedResolution.FailureReason;
                        effectiveFormulaText = storedFormulaText;
                        effectiveFormulaProven = false;
                    }
                    else
                    {
                        effectiveFormulaText = sharedResolution.EffectiveText;
                        formulaRange = sharedResolution.Range;
                        sharedIndex = sharedResolution.SharedIndex;
                        wasReconstructed = sharedResolution.WasReconstructed;
                        formulaForm = sharedResolution.IsMaster
                            ? OpenXmlFormulaForm.SharedMaster
                            : OpenXmlFormulaForm.SharedFollower;
                    }
                    break;
                case "array":
                    formulaForm = metadataResolution.IsDynamicArray
                        ? OpenXmlFormulaForm.DynamicArray
                        : OpenXmlFormulaForm.Array;
                    formulaCoverageCode = metadataResolution.IsDynamicArray
                        ? OpenXmlCoverageCodes.DynamicArrayNotExpanded
                        : OpenXmlCoverageCodes.ArrayFormulaNotExpanded;
                    formulaCoverageMessage = metadataResolution.IsDynamicArray
                        ? "The formula has an XLDAPR fDynamic=true marker. Exact formula/range/metadata and the " +
                          "stored cache are retained, but spill membership and calculated results are not inferred."
                        : "The array master formula and exact ref metadata are retained, but the array range is " +
                          "not expanded into manufactured formula cells.";
                    break;
                case "dataTable":
                    formulaForm = OpenXmlFormulaForm.DataTable;
                    formulaCoverageCode = OpenXmlCoverageCodes.DataTableFormulaUnsupported;
                    formulaCoverageMessage =
                        "The data-table formula and exact attributes are retained, but TABLE calculation " +
                        "semantics are not evaluated or projected.";
                    break;
                default:
                    formulaForm = OpenXmlFormulaForm.Unknown;
                    formulaCoverageCode = OpenXmlCoverageCodes.UnknownFormulaType;
                    formulaCoverageMessage =
                        "The formula type is unknown. Stored text and exact attributes are retained only.";
                    break;
            }

            if (metadataResolution.IsDynamicArray && formulaType != "array")
            {
                formulaForm = OpenXmlFormulaForm.DynamicArray;
                cellCoverageLimiter.Add(
                    OpenXmlCoverageCodes.DynamicArrayNotExpanded,
                    "The formula has an XLDAPR fDynamic=true marker outside the expected array form. Exact " +
                    "metadata is retained, but spill semantics are not inferred.",
                    cellReference);
            }

            if (formulaCoverageCode != null)
            {
                string message = formulaCoverageMessage ??
                                 ("Formula type '" + formulaType +
                                  "' is retained but cannot be reconstructed safely.");
                cellCoverageLimiter.Add(formulaCoverageCode, message, cellReference);
            }

            string observationEffectiveFormulaText = effectiveFormulaProven
                ? effectiveFormulaText
                : null;
            if (string.IsNullOrEmpty(effectiveFormulaText))
            {
                string code = formulaCoverageCode ?? OpenXmlCoverageCodes.EmptyFormulaText;
                if (formulaCoverageCode == null)
                {
                    cellCoverageLimiter.Add(
                        code,
                        "The formula element has no stored formula text.",
                        cellReference);
                }

                // CanonicalCellState requires non-empty formula text. Formula attributes/group evidence are
                // compared separately, so the placeholder is intentionally code-only and locator-free.
                effectiveFormulaText = InternEvidence(
                    formulaEvidenceCache,
                    FormulaCoverageToken(code));
            }

            string externalProbe = string.IsNullOrEmpty(observationEffectiveFormulaText)
                ? storedFormulaText
                : observationEffectiveFormulaText;
            if (OpenXmlFormulaReferenceClassifier.HasExternalWorkbookReference(externalProbe))
            {
                cellCoverageLimiter.Add(
                    OpenXmlCoverageCodes.ExternalFormulaDependencyNotResolved,
                    "The exact external-link formula text is retained, but no external workbook relationship " +
                    "or dependency value was opened or resolved.",
                    cellReference,
                    affectsExactCellComparison: false);
            }

            CanonicalValue cachedResult = null;
            if (cell.CellValue != null || cell.InlineString != null)
            {
                cachedResult = ReadTypedValue(
                    cell,
                    cellReference,
                    sheetName,
                    partUri,
                    sharedStrings,
                    dateStyles,
                    dateSystem,
                    cellCoverageLimiter,
                    coverage);
            }

            string exactMetadata = CanonicalizeFormulaMetadata(
                formula,
                sharedResolution,
                metadataResolution);
            exactMetadata = InternEvidence(formulaEvidenceCache, exactMetadata);
            formulaObservation = new OpenXmlFormulaObservation(
                cellReference,
                formulaForm,
                storedFormulaText,
                string.IsNullOrEmpty(observationEffectiveFormulaText)
                    ? null
                    : observationEffectiveFormulaText,
                string.IsNullOrEmpty(formulaRange) ? null : formulaRange,
                sharedIndex,
                wasReconstructed,
                exactMetadata,
                cachedResult);

            return CanonicalCellState.Formula(
                cellReference,
                effectiveFormulaText,
                cachedResult,
                exactFormulaMetadata: exactMetadata);
        }

        private static string CanonicalizeFormulaMetadata(
            CellFormula formula,
            OpenXmlSharedFormulaResolution sharedResolution,
            OpenXmlFormulaMetadataResolution metadataResolution)
        {
            List<OpenXmlAttribute> attributes = formula.GetAttributes()
                .OrderBy(attribute => attribute.NamespaceUri ?? "", StringComparer.Ordinal)
                .ThenBy(attribute => attribute.LocalName ?? "", StringComparer.Ordinal)
                .ToList();
            var result = new StringBuilder();
            foreach (OpenXmlAttribute attribute in attributes)
            {
                AppendFormulaToken(result, attribute.NamespaceUri);
                AppendFormulaToken(result, attribute.LocalName);
                AppendFormulaToken(result, attribute.Value);
            }

            if (sharedResolution != null)
            {
                AppendFormulaToken(result, "shared-master-cell");
                AppendFormulaToken(result, sharedResolution.MasterCellReference);
                AppendFormulaToken(result, "shared-range");
                AppendFormulaToken(result, sharedResolution.Range);
                AppendFormulaToken(result, "shared-resolution");
                AppendFormulaToken(
                    result,
                    sharedResolution.IsResolved
                        ? (sharedResolution.WasReconstructed ? "follower-reconstructed" : "master-stored")
                        : "unresolved");
            }

            if (metadataResolution.ExactEvidence != null)
            {
                AppendFormulaToken(result, "cell-metadata");
                AppendFormulaToken(result, metadataResolution.ExactEvidence);
            }

            return result.Length == 0 ? null : result.ToString();
        }

        private static void AppendFormulaToken(StringBuilder target, string value)
        {
            value = value ?? "";
            target.Append(value.Length.ToString(CultureInfo.InvariantCulture));
            target.Append(':');
            target.Append(value);
        }

        private static long FormulaTokenLength(string value)
        {
            int length = value == null ? 0 : value.Length;
            return length + 1L + length.ToString(CultureInfo.InvariantCulture).Length;
        }

        private static string InternEvidence(IDictionary<string, string> cache, string evidence)
        {
            if (evidence == null) return null;
            if (cache.TryGetValue(evidence, out string retained)) return retained;
            cache.Add(evidence, evidence);
            return evidence;
        }

        private static CanonicalValue ReadConstantValue(
            Cell cell,
            string cellReference,
            string sheetName,
            string partUri,
            string[] sharedStrings,
            IList<DateStyleKind> dateStyles,
            ExcelDateSystem dateSystem,
            CellCoverageLimiter cellCoverageLimiter,
            IList<OpenXmlCoverageMessage> coverage)
        {
            string dataType = GetAttributeValue(cell, "t");
            if (cell.CellValue == null && cell.InlineString == null && string.IsNullOrEmpty(dataType))
                return CanonicalValue.Blank;

            return ReadTypedValue(
                cell,
                cellReference,
                sheetName,
                partUri,
                sharedStrings,
                dateStyles,
                dateSystem,
                cellCoverageLimiter,
                coverage);
        }

        private static CanonicalValue ReadTypedValue(
            Cell cell,
            string cellReference,
            string sheetName,
            string partUri,
            string[] sharedStrings,
            IList<DateStyleKind> dateStyles,
            ExcelDateSystem dateSystem,
            CellCoverageLimiter cellCoverageLimiter,
            IList<OpenXmlCoverageMessage> coverage)
        {
            string dataType = GetAttributeValue(cell, "t");
            string raw = cell.CellValue == null ? "" : cell.CellValue.Text ?? "";

            switch (dataType)
            {
                case "s":
                    int sharedIndex;
                    if (int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out sharedIndex) &&
                        sharedIndex >= 0 && sharedIndex < sharedStrings.Length)
                    {
                        return CanonicalValue.Text(sharedStrings[sharedIndex]);
                    }

                    return CoverageValue(
                        OpenXmlCoverageCodes.InvalidSharedStringIndex,
                        "Shared-string index '" + Abbreviate(raw) + "' is missing, malformed, or out of range.",
                        cellReference,
                        cell,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage);

                case "inlineStr":
                    return CanonicalValue.Text(ReadInlineString(cell.InlineString));

                case "str":
                    return CanonicalValue.Text(raw);

                case "b":
                    if (raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase))
                        return CanonicalValue.Boolean(true);
                    if (raw == "0" || string.Equals(raw, "false", StringComparison.OrdinalIgnoreCase))
                        return CanonicalValue.Boolean(false);
                    return CoverageValue(
                        OpenXmlCoverageCodes.InvalidBooleanValue,
                        "Boolean cell value '" + Abbreviate(raw) + "' is invalid.",
                        cellReference,
                        cell,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage);

                case "e":
                    if (!string.IsNullOrEmpty(raw)) return CanonicalValue.Error(raw);
                    return CoverageValue(
                        OpenXmlCoverageCodes.InvalidErrorValue,
                        "An error cell has no stored error token.",
                        cellReference,
                        cell,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage);

                case "d":
                    if (IsValidIsoDateValue(raw))
                        return CanonicalValue.IsoDateTime(raw);
                    return CoverageValue(
                        OpenXmlCoverageCodes.InvalidIsoDateValue,
                        "ISO date cell value '" + Abbreviate(raw) + "' could not be validated conservatively.",
                        cellReference,
                        cell,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage);

                case "":
                case "n":
                    double number;
                    if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                        double.IsNaN(number) || double.IsInfinity(number))
                    {
                        return CoverageValue(
                            OpenXmlCoverageCodes.InvalidNumericValue,
                            "Numeric cell value '" + Abbreviate(raw) + "' is invalid.",
                            cellReference,
                            cell,
                            sheetName,
                            partUri,
                            cellCoverageLimiter,
                            coverage);
                    }

                    // Number formats are presentation metadata. Retain an invalid-style coverage diagnostic for
                    // broad package auditing, but never let a date-looking format change the exact stored numeric
                    // value or calculated-value channel.
                    ResolveDateStyle(
                        cell,
                        dateStyles,
                        cellReference,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage);
                    return CanonicalValue.Number(number);

                default:
                    return CoverageValue(
                        OpenXmlCoverageCodes.UnsupportedCellDataType,
                        "Cell data type '" + Abbreviate(dataType) + "' is unsupported.",
                        cellReference,
                        cell,
                        sheetName,
                        partUri,
                        cellCoverageLimiter,
                        coverage);
            }
        }

        private static DateStyleKind ResolveDateStyle(
            Cell cell,
            IList<DateStyleKind> dateStyles,
            string cellReference,
            string sheetName,
            string partUri,
            CellCoverageLimiter cellCoverageLimiter,
            IList<OpenXmlCoverageMessage> coverage)
        {
            string rawStyleIndex = GetAttributeValue(cell, "s");
            if (string.IsNullOrEmpty(rawStyleIndex)) return DateStyleKind.Number;

            uint styleIndex;
            if (!uint.TryParse(rawStyleIndex, NumberStyles.None, CultureInfo.InvariantCulture, out styleIndex) ||
                styleIndex >= dateStyles.Count)
            {
                cellCoverageLimiter.Add(
                    OpenXmlCoverageCodes.InvalidCellStyleIndex,
                    "Cell style index '" + Abbreviate(rawStyleIndex) +
                    "' is missing from the workbook style table; the numeric value is retained as a number.",
                    cellReference,
                    affectsExactCellComparison: false);
                return DateStyleKind.Number;
            }

            return dateStyles[(int)styleIndex];
        }

        private static string[] ReadSharedStrings(
            WorkbookPart workbookPart,
            CancellationToken cancellationToken)
        {
            SharedStringTable table = workbookPart.SharedStringTablePart == null
                ? null
                : workbookPart.SharedStringTablePart.SharedStringTable;
            if (table == null) return Array.Empty<string>();

            var result = new List<string>();
            foreach (SharedStringItem item in table.Elements<SharedStringItem>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                result.Add(ReadStringItem(item));
            }
            return result.ToArray();
        }

        private static string ReadStringItem(SharedStringItem item)
        {
            var result = new StringBuilder();
            foreach (OpenXmlElement child in item.ChildElements)
            {
                var text = child as SpreadsheetText;
                if (text != null)
                {
                    result.Append(text.Text);
                    continue;
                }

                var run = child as Run;
                if (run != null)
                {
                    SpreadsheetText runText = run.GetFirstChild<SpreadsheetText>();
                    if (runText != null) result.Append(runText.Text);
                }
            }

            return result.ToString();
        }

        private static string ReadInlineString(InlineString inlineString)
        {
            if (inlineString == null) return "";

            var result = new StringBuilder();
            foreach (OpenXmlElement child in inlineString.ChildElements)
            {
                var text = child as SpreadsheetText;
                if (text != null)
                {
                    result.Append(text.Text);
                    continue;
                }

                var run = child as Run;
                if (run != null)
                {
                    SpreadsheetText runText = run.GetFirstChild<SpreadsheetText>();
                    if (runText != null) result.Append(runText.Text);
                }
            }

            return result.ToString();
        }

        private static IList<DateStyleKind> ReadDateStyles(
            WorkbookPart workbookPart,
            CancellationToken cancellationToken)
        {
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart == null
                ? null
                : workbookPart.WorkbookStylesPart.Stylesheet;
            CellFormats cellFormats = stylesheet == null ? null : stylesheet.CellFormats;
            if (cellFormats == null) return Array.Empty<DateStyleKind>();

            var customFormats = new Dictionary<uint, string>();
            if (stylesheet.NumberingFormats != null)
            {
                foreach (NumberingFormat format in stylesheet.NumberingFormats.Elements<NumberingFormat>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (format.NumberFormatId != null)
                        customFormats[format.NumberFormatId.Value] =
                            format.FormatCode == null ? "" : format.FormatCode.Value;
                }
            }

            var result = new List<DateStyleKind>();
            foreach (CellFormat format in cellFormats.Elements<CellFormat>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                uint numberFormatId = format.NumberFormatId == null ? 0U : format.NumberFormatId.Value;
                string customCode;
                bool isDate = IsBuiltInDateFormat(numberFormatId) ||
                              (customFormats.TryGetValue(numberFormatId, out customCode) &&
                               IsConservativeCustomDateFormat(customCode));
                result.Add(isDate ? DateStyleKind.Date : DateStyleKind.Number);
            }

            return result;
        }

        private static bool IsBuiltInDateFormat(uint id)
        {
            return (id >= 14 && id <= 22) ||
                   (id >= 27 && id <= 36) ||
                   (id >= 45 && id <= 47) ||
                   (id >= 50 && id <= 58);
        }

        private static bool IsConservativeCustomDateFormat(string formatCode)
        {
            if (string.IsNullOrEmpty(formatCode)) return false;

            var visibleTokens = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < formatCode.Length; i++)
            {
                char c = formatCode[i];
                if (c == '"')
                {
                    quoted = !quoted;
                    continue;
                }
                if (quoted) continue;
                if (c == '\\' || c == '_' || c == '*')
                {
                    i++;
                    continue;
                }
                if (c == '[')
                {
                    int close = formatCode.IndexOf(']', i + 1);
                    if (close < 0) return false;
                    string bracket = formatCode.Substring(i + 1, close - i - 1).ToLowerInvariant();
                    if (bracket == "h" || bracket == "hh" ||
                        bracket == "m" || bracket == "mm" ||
                        bracket == "s" || bracket == "ss")
                    {
                        visibleTokens.Append(bracket);
                    }
                    i = close;
                    continue;
                }
                visibleTokens.Append(char.ToLowerInvariant(c));
            }

            string tokens = visibleTokens.ToString();
            return tokens.IndexOf('y') >= 0 ||
                   tokens.IndexOf('d') >= 0 ||
                   tokens.IndexOf('h') >= 0 ||
                   tokens.IndexOf('m') >= 0 ||
                   tokens.IndexOf('s') >= 0;
        }

        private static ExcelDateSystem ReadDateSystem(Workbook workbook)
        {
            bool date1904 = workbook.WorkbookProperties != null &&
                            workbook.WorkbookProperties.Date1904 != null &&
                            workbook.WorkbookProperties.Date1904.Value;
            return date1904 ? ExcelDateSystem.Date1904 : ExcelDateSystem.Date1900;
        }

        private static OpenXmlSheetVisibility ReadVisibility(Sheet sheet, out bool known)
        {
            string state = GetAttributeValue(sheet, "state");
            known = true;
            switch (state)
            {
                case "":
                case "visible":
                    return OpenXmlSheetVisibility.Visible;
                case "hidden":
                    return OpenXmlSheetVisibility.Hidden;
                case "veryHidden":
                    return OpenXmlSheetVisibility.VeryHidden;
                default:
                    known = false;
                    return OpenXmlSheetVisibility.Unknown;
            }
        }

        private static bool TryNormalizeCellReference(string value, out string normalized)
        {
            normalized = null;
            if (string.IsNullOrEmpty(value)) return false;

            int index = 0;
            int column = 0;
            while (index < value.Length && IsAsciiLetter(value[index]))
            {
                char upper = char.ToUpperInvariant(value[index]);
                column = checked(column * 26 + (upper - 'A' + 1));
                index++;
                if (index > 3) return false;
            }

            if (index == 0 || column <= 0 || column > 16384 || index == value.Length) return false;

            int row = 0;
            int rowStart = index;
            while (index < value.Length && value[index] >= '0' && value[index] <= '9')
            {
                int digit = value[index] - '0';
                if (row > (1048576 - digit) / 10) return false;
                row = row * 10 + digit;
                index++;
                if (row > 1048576) return false;
            }

            if (index != value.Length || row <= 0 || value[rowStart] == '0') return false;
            normalized = value.Substring(0, rowStart).ToUpperInvariant() +
                         row.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static bool IsAsciiLetter(char value) =>
            (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');

        private static string GetAttributeValue(OpenXmlElement element, string localName)
        {
            foreach (OpenXmlAttribute attribute in element.GetAttributes())
            {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal) &&
                    string.IsNullOrEmpty(attribute.NamespaceUri))
                {
                    return attribute.Value ?? "";
                }
            }

            return "";
        }

        private static uint? TryReadUInt32Attribute(OpenXmlElement element, string localName)
        {
            string raw = GetAttributeValue(element, localName);
            return uint.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out uint value)
                ? (uint?)value
                : null;
        }

        private static bool IsValidIsoDateValue(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return false;
            if (!HasIsoDateShape(raw)) return false;

            int timeSeparator = raw.IndexOf('T');
            if (raw.EndsWith("Z", StringComparison.OrdinalIgnoreCase) ||
                (timeSeparator >= 0 &&
                 (raw.IndexOf('+', timeSeparator) >= 0 || raw.IndexOf('-', timeSeparator + 1) >= 0)))
            {
                // Excel serial dates have no timezone. Converting an offset value would silently change meaning.
                return false;
            }

            DateTime value;
            return DateTime.TryParse(
                    raw,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces,
                    out value);
        }

        private static bool HasIsoDateShape(string raw)
        {
            if (raw.Length < 10 || raw[4] != '-' || raw[7] != '-') return false;
            for (int i = 0; i < 10; i++)
            {
                if (i == 4 || i == 7) continue;
                if (raw[i] < '0' || raw[i] > '9') return false;
            }

            return raw.Length == 10 || raw[10] == 'T';
        }

        private static CanonicalValue CoverageValue(
            string code,
            string message,
            string cellReference,
            Cell cell,
            string sheetName,
            string partUri,
            CellCoverageLimiter cellCoverageLimiter,
            IList<OpenXmlCoverageMessage> coverage)
        {
            cellCoverageLimiter.Add(code, message, cellReference);
            return CanonicalValue.Text(ValueCoverageToken(code, ValuePayloadContentId(cell)));
        }

        private static string ValuePayloadContentId(Cell cell)
        {
            var evidence = new StringBuilder();
            AppendFormulaToken(evidence, "cell-value-type");
            AppendFormulaToken(evidence, CanonicalValuePayloadType(cell));
            AppendFormulaToken(evidence, "cell-value-metadata-index");
            AppendFormulaToken(evidence, CanonicalUnsignedAttribute(cell, "vm"));
            if (cell != null)
            {
                foreach (CellValue value in cell.Elements<CellValue>())
                {
                    AppendFormulaToken(evidence, "v");
                    AppendFormulaToken(evidence, value.Text ?? "");
                }
                foreach (InlineString inline in cell.Elements<InlineString>())
                {
                    AppendFormulaToken(evidence, "is");
                    AppendFormulaToken(evidence, ReadInlineString(inline));
                }
            }
            return StableId.ComputeByteContentId(Encoding.UTF8.GetBytes(evidence.ToString()));
        }

        private static string CanonicalValuePayloadType(Cell cell)
        {
            if (cell == null) return "";
            string raw = GetAttributeValue(cell, "t");
            return string.IsNullOrEmpty(raw) || string.Equals(raw, "n", StringComparison.Ordinal)
                ? "numeric"
                : raw;
        }

        private static string CanonicalUnsignedAttribute(OpenXmlElement element, string localName)
        {
            if (element == null) return "";
            string raw = GetAttributeValue(element, localName);
            return uint.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out uint value)
                ? value.ToString(CultureInfo.InvariantCulture)
                : raw;
        }

        private static string ValueCoverageToken(string code, string valuePayloadContentId) =>
            "#OPENXML-COVERAGE[" + code + ";" + valuePayloadContentId + "]";

        private static string FormulaCoverageToken(string code) =>
            "#OPENXML-COVERAGE[" + code + "]";

        private static OpenXmlWorksheetSnapshot EmptySheet(
            string name,
            uint? sheetId,
            string relationshipId,
            string partUri,
            string codeName,
            OpenXmlSheetVisibility visibility,
            string snapshotRevisionId,
            IList<OpenXmlCoverageMessage> coverage) =>
            new OpenXmlWorksheetSnapshot(
                name,
                sheetId,
                relationshipId,
                partUri,
                codeName,
                visibility,
                OpenXmlPhysicalSheetKind.Unsupported,
                new RevisionSnapshot(snapshotRevisionId, null),
                null,
                null,
                coverage);

        private static OpenXmlCoverageMessage Coverage(
            string code,
            string message,
            bool affectsExactComparison,
            string sheetName = null,
            string cellReference = null,
            string partUri = null,
            bool? affectsExactCellComparison = null) =>
            new OpenXmlCoverageMessage(
                code,
                message,
                affectsExactComparison,
                affectsExactCellComparison ?? affectsExactComparison,
                sheetName,
                cellReference,
                partUri);

        private static string Abbreviate(string value)
        {
            if (value == null) return "";
            return value.Length <= 80 ? value : value.Substring(0, 77) + "...";
        }

        private sealed class WorkbookCellReadBudget
        {
            private readonly long _maximum;
            private long _observed;

            internal WorkbookCellReadBudget(long maximum)
            {
                _maximum = maximum;
            }

            internal void Observe()
            {
                _observed++;
                if (_observed <= _maximum) return;
                throw new InvalidDataException(
                    "The workbook contains more than " +
                    _maximum.ToString("N0", CultureInfo.InvariantCulture) +
                    " physical cell elements and exceeds the safe in-memory parser limit. " +
                    "Malformed, duplicate, blank, and formatting-only cells count at this boundary.");
            }
        }

        private sealed class CellCoverageLimiter
        {
            private const int MaxRepresentativesPerCode = 8;

            private readonly IList<OpenXmlCoverageMessage> _coverage;
            private readonly string _sheetName;
            private readonly string _partUri;
            private readonly IDictionary<string, CoverageBucket> _buckets;

            internal CellCoverageLimiter(
                IList<OpenXmlCoverageMessage> coverage,
                string sheetName,
                string partUri)
            {
                _coverage = coverage;
                _sheetName = sheetName;
                _partUri = partUri;
                _buckets = new Dictionary<string, CoverageBucket>(StringComparer.Ordinal);
            }

            internal void Add(
                string code,
                string message,
                string cellReference,
                bool affectsExactCellComparison = true)
            {
                string key = code + (affectsExactCellComparison ? "\u0001" : "\u0000");
                if (!_buckets.TryGetValue(key, out CoverageBucket bucket))
                {
                    bucket = new CoverageBucket(code, affectsExactCellComparison);
                    _buckets.Add(key, bucket);
                }
                bucket.Observed++;
                if (bucket.Emitted >= MaxRepresentativesPerCode) return;

                bucket.Emitted++;
                _coverage.Add(Coverage(
                    code,
                    message,
                    true,
                    _sheetName,
                    cellReference,
                    _partUri,
                    bucket.AffectsExactCellComparison));
            }

            internal void FlushSummaries()
            {
                foreach (KeyValuePair<string, CoverageBucket> pair in _buckets.OrderBy(
                             item => item.Key,
                             StringComparer.Ordinal))
                {
                    long suppressed = pair.Value.Observed - pair.Value.Emitted;
                    if (suppressed <= 0) continue;
                    _coverage.Add(Coverage(
                        pair.Value.Code,
                        suppressed.ToString(CultureInfo.InvariantCulture) +
                        " additional cell-level findings with this code were grouped after " +
                        pair.Value.Emitted.ToString(CultureInfo.InvariantCulture) +
                        " representative diagnostics.",
                        true,
                        _sheetName,
                        partUri: _partUri,
                        affectsExactCellComparison: pair.Value.AffectsExactCellComparison));
                }
            }

            private sealed class CoverageBucket
            {
                internal CoverageBucket(string code, bool affectsExactCellComparison)
                {
                    Code = code;
                    AffectsExactCellComparison = affectsExactCellComparison;
                }

                internal string Code { get; }
                internal bool AffectsExactCellComparison { get; }
                internal long Observed;
                internal int Emitted;
            }
        }

        private enum DateStyleKind
        {
            Number,
            Date,
        }

        private struct ZipEnvelope
        {
            internal ZipEnvelope(int entryCount, long centralDirectorySize)
            {
                EntryCount = entryCount;
                CentralDirectorySize = centralDirectorySize;
            }

            internal int EntryCount { get; }
            internal long CentralDirectorySize { get; }
        }
    }
}
