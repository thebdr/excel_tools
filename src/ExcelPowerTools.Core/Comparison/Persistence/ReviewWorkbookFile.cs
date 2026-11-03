using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>
    /// Creates and reads the macro-free comparison review workbook. Public callers remain independent of the
    /// Open XML SDK; package types are an implementation detail confined to this persistence boundary.
    /// </summary>
    public static partial class ReviewWorkbookFile
    {
        public const string CustomXmlNamespace = "urn:excel-power-tools:comparison:review:v1";
        public const string DecisionCustomXmlNamespace =
            "urn:excel-power-tools:comparison:review:decisions:v1";
        public const string Marker = "ExcelPowerTools.Comparison.Review";
        public const string DecisionMarker = "ExcelPowerTools.Comparison.Review.Decisions";
        public const string SchemaVersion = "1";
        public const string DecisionBoundManifestSchemaVersion = "2";
        public const string StreamingProjectionManifestSchemaVersion = "3";
        public const string DecisionBoundStreamingProjectionManifestSchemaVersion = "4";
        public const string DecisionEventStateSchemaVersion = "2";
        public const string StreamingProjectionDigestVersion = "review-projection-v2";

        private const string SpreadsheetNamespace =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";
        // These are deliberately conservative cross-architecture spike limits. The final scale gate may move
        // larger projects to the archive-first store, but reopening an untrusted workbook must never make a
        // 32-bit Excel process attempt an unbounded allocation.
        public const int MaxEmbeddedStateBytes = 32 * 1024 * 1024;
        public const int MaxStateChunkSizeBytes = 4 * 1024 * 1024;
        public const int MaxStateChunkCount = 2048;
        public const int MaxExcelCellTextCharacters = 32767;
        public const int MaxProjectedChangeRows = 1048575;

        private const int MaxCustomXmlPartCount = MaxStateChunkCount + 64;
        private const int MaxChunkXmlOverheadCharacters = 4096;
        private const long MaxOwnedCustomXmlCharacters =
            ((long)MaxStateChunkSizeBytes + 2) / 3 * 4 + MaxChunkXmlOverheadCharacters;
        private const long MaxManifestCharacters = 2 * 1024 * 1024;
        public const long MaxDecisionXmlCharacters = 16L * 1024 * 1024;
        public const long MaxDecisionXmlBytes = 16L * 1024 * 1024;
        public const int MaxDecisionRevisionCount = 100;
        public const int MaxDecisionSubjectCount = 25000;
        public const int MaxDecisionObservationCount = 100000;
        public const int MaxDecisionGroupCount = 25000;
        public const int MaxDecisionMembershipCount = 250000;
        public const int MaxDecisionEventCount = 10000;
        public const int MaxDecisionAffectedSubjectCountPerEvent = MaxDecisionSubjectCount;
        public const int MaxDecisionSupportingRevisionCountPerEvent = MaxDecisionRevisionCount;
        public const int MaxDecisionSupportingContentHashCountPerEvent = MaxDecisionRevisionCount;
        public const int MaxDecisionReviewerSourceCountPerEvent = 64;
        public const int MaxDecisionLocationCountPerEvent = 256;
        public const int MaxDecisionXmlElementCount = 250000;
        public const int MaxDecisionXmlDepth = 8;
        public const int MaxDecisionAttributesPerElement = 32;
        public const int MaxDecisionPayloadBytes = 4 * 1024 * 1024;
        private const int MaxDecisionIdentityCharacters = 4096;
        private const int MaxDecisionLabelCharacters = MaxExcelCellTextCharacters;
        private const long MaxSpreadsheetPartCharacters = 64L * 1024 * 1024;
        public const long MaxStreamingWorksheetCharacters = 512L * 1024 * 1024;
        public const long MaxStreamingSharedStringCharacters = 256L * 1024 * 1024;
        public const int MaxStreamingSharedStringCacheEntries = 4096;
        public const int MaxStreamingSharedStringCacheCharacters = 1024 * 1024;
        public const long MaxReviewPackageBytes = 128L * 1024 * 1024;
        private const int MaxReviewPackageEntries = 10000;
        private const long MaxCentralDirectoryBytes = 16L * 1024 * 1024;
        private const long MaxEntryUncompressedBytes = 64L * 1024 * 1024;
        private const long MaxDeclaredUncompressedBytes = 768L * 1024 * 1024;
        private const long CompressionRatioCheckThresholdBytes = 1024L * 1024;
        private const double MaxCompressionRatio = 1200d;
        private const double MinVisibleRowHeightPoints = 8d;
        private const double MinVisibleColumnWidth = 1d;
        private const double DefaultVisibleColumnWidth = 8.43d;
        private const double MaxVisibleColumnWidth = 255d;
        private const int ContentWidthPaddingCharacters = 2;
        private const double MinVisibleFontSizePoints = 8d;
        private const string DetailTruncationMarker =
            "\u2026 [truncated; full detail is preserved in embedded comparison state]";
        private const string LegacyProjectionDigestVersion = "review-projection-v1";

        private static readonly XNamespace StateNamespace = CustomXmlNamespace;
        private static readonly XNamespace DecisionNamespace = DecisionCustomXmlNamespace;
        private static readonly ISet<string> AllowedReviewRunRgb =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "FF404040",
                "FFC62828",
                "FF2E7D32",
            };
        private static readonly ISet<string> AllowedVisibleFontNames =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Aptos",
                "Aptos Narrow",
                "Arial",
                "Calibri",
                "Cambria",
                "Courier New",
                "Segoe UI",
                "Times New Roman",
            };
        private static readonly IReadOnlyList<string> ExpectedSheetNames =
            new ReadOnlyCollection<string>(new[]
            {
                "Overview",
                "Revisions",
                "Changes",
                "Details",
                "Snapshots",
                "Diagnostics",
                "Merge Log",
                "Legend",
            });

        public static IReadOnlyList<string> SheetNames => ExpectedSheetNames;

        /// <summary>
        /// Creates a new workbook without overwriting. A temporary package is created in the destination
        /// directory, closed, reopened and verified before an atomic same-directory move publishes it.
        /// </summary>
        public static void CreateNew(
            string destinationPath,
            ReviewWorkbookState state,
            ReviewWorkbookWriteOptions options = null)
        {
            if (string.IsNullOrWhiteSpace(destinationPath))
                throw new ArgumentException("A destination path is required.", nameof(destinationPath));
            if (state == null) throw new ArgumentNullException(nameof(state));
            options = options ?? new ReviewWorkbookWriteOptions();
            if (options.StateChunkSizeBytes <= 0 ||
                options.StateChunkSizeBytes > MaxStateChunkSizeBytes)
                throw new ArgumentOutOfRangeException(
                    nameof(options.StateChunkSizeBytes),
                    "The state chunk size must be between 1 and " +
                    MaxStateChunkSizeBytes.ToString(CultureInfo.InvariantCulture) + " bytes.");

            string destination = Path.GetFullPath(destinationPath);
            if (!string.Equals(Path.GetExtension(destination), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("A review workbook destination must end in .xlsx.", nameof(destinationPath));
            EnsureSupportedPublicationPath(destination);

            string directory = Path.GetDirectoryName(destination);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException("The review workbook destination directory does not exist.");
            if (File.Exists(destination) || Directory.Exists(destination))
                throw new IOException("The review workbook destination already exists: " + destination);

            if (state.StateJson.Length > MaxEmbeddedStateBytes)
                throw new InvalidDataException("The embedded comparison state exceeds the safe review-workbook limit.");
            if (state.ChangeRows.Count > MaxProjectedChangeRows)
                throw new InvalidDataException("The visible change projection exceeds Excel's worksheet row limit.");
            byte[] payload = new UTF8Encoding(false, true).GetBytes(state.StateJson);
            if (payload.Length > MaxEmbeddedStateBytes)
                throw new InvalidDataException("The UTF-8 comparison state exceeds the safe review-workbook limit.");
            ValidateDecisionGenerationForWrite(state);
            List<StateChunk> chunks = SplitIntoChunks(payload, options.StateChunkSizeBytes);
            string totalSha256 = ComputeSha256(payload);
            string manifestItemId = NewItemId();
            string decisionItemId = state.DecisionGeneration == null ? "" : NewItemId();
            string temporary = Path.Combine(
                directory,
                "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                using (var output = new FileStream(
                    temporary,
                    FileMode.CreateNew,
                    FileAccess.ReadWrite,
                    FileShare.None,
                    64 * 1024,
                    FileOptions.SequentialScan))
                {
                    BuildPackage(
                        output,
                        state,
                        chunks,
                        manifestItemId,
                        decisionItemId,
                        totalSha256);
                }

                VerifyClosedPackage(temporary, state, chunks.Count, totalSha256);
                using (var durableStage = new FileStream(
                    temporary,
                    FileMode.Open,
                    FileAccess.ReadWrite,
                    FileShare.None,
                    4096,
                    FileOptions.WriteThrough))
                {
                    durableStage.Flush(true);
                }
                File.Move(temporary, destination);
            }
            catch
            {
                TryDelete(temporary);
                throw;
            }
        }

        /// <summary>Reads the embedded project state and verifies all marker, schema and chunk invariants.</summary>
        public static ReviewWorkbookReadResult ReadAndValidate(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
                throw new ArgumentException("A workbook path is required.", nameof(workbookPath));

            string path = Path.GetFullPath(workbookPath);
            using (var packageStream = new FileStream(
                       path,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.Read,
                       64 * 1024,
                       FileOptions.SequentialScan))
            {
                return ReadAndValidate(packageStream);
            }
        }

        /// <summary>
        /// Reads and validates a review package from an already-open stream. The caller retains ownership of
        /// the stream and its original position is restored. This overload lets security-sensitive callers
        /// keep the underlying file handle protected against writes, deletion and path replacement for the
        /// entire validation/open boundary.
        /// </summary>
        public static ReviewWorkbookReadResult ReadAndValidate(Stream packageStream)
        {
            if (packageStream == null) throw new ArgumentNullException(nameof(packageStream));
            if (!packageStream.CanRead || !packageStream.CanSeek)
                throw new InvalidDataException("The review workbook package must be a readable, seekable stream.");

            long originalPosition = packageStream.Position;
            try
            {
                packageStream.Position = 0;
                PreflightReviewPackage(packageStream);
                packageStream.Position = 0;

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(
                           packageStream,
                           false,
                           CreateBoundedOpenSettings()))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    if (workbookPart == null)
                        throw new InvalidDataException("The package has no workbook part.");

                    ValidatePassiveReviewPackage(document, workbookPart);

                    ManifestPartData manifest = null;
                    DecisionPartData decision = null;
                    var chunksByItemId = new Dictionary<Guid, ChunkPartData>();
                    var recognizedItemIds = new HashSet<Guid>();

                    List<CustomXmlPart> customParts = workbookPart.CustomXmlParts
                        .Take(MaxCustomXmlPartCount + 1)
                        .ToList();
                    if (customParts.Count > MaxCustomXmlPartCount)
                        throw new InvalidDataException("The workbook contains too many Custom XML parts to inspect safely.");

                    foreach (CustomXmlPart part in customParts)
                    {
                        CustomXmlRoot root = ReadCustomXmlRoot(part);
                        bool isReviewNamespace = string.Equals(
                            root.NamespaceName,
                            CustomXmlNamespace,
                            StringComparison.Ordinal);
                        bool isDecisionNamespace = string.Equals(
                            root.NamespaceName,
                            DecisionCustomXmlNamespace,
                            StringComparison.Ordinal);
                        if (!isReviewNamespace && !isDecisionNamespace)
                            continue;

                        Guid persistentItemId = ReadAndValidatePersistentItemId(
                            part,
                            isDecisionNamespace ? DecisionCustomXmlNamespace : CustomXmlNamespace);
                        if (!recognizedItemIds.Add(persistentItemId))
                            throw new InvalidDataException("Duplicate comparison Custom XML item ID: " + persistentItemId);

                        if (isDecisionNamespace)
                        {
                            if (root.LocalName != "decisionState")
                                throw new InvalidDataException(
                                    "Unknown element in the comparison decision Custom XML namespace: " +
                                    root.LocalName);
                            if (decision != null)
                                throw new InvalidDataException(
                                    "The workbook contains more than one comparison decision state.");
                            decision = ReadDecisionStatePart(part, persistentItemId);
                            continue;
                        }

                        if (root.LocalName == "reviewManifest")
                        {
                            if (manifest != null)
                                throw new InvalidDataException("The workbook contains more than one comparison manifest.");
                            XDocument xml = LoadCustomXml(part, MaxManifestCharacters);
                            XElement element = xml.Root;
                            if (element == null)
                                throw new InvalidDataException("The comparison manifest is empty.");
                            string manifestSchemaVersion = ValidateManifestEnvelope(
                                element,
                                persistentItemId);
                            manifest = ParseManifest(
                                element,
                                persistentItemId,
                                manifestSchemaVersion);
                        }
                        else if (root.LocalName == "stateChunk")
                        {
                            ChunkPartData chunk = ReadStateChunkMetadata(part, persistentItemId);
                            if (chunksByItemId.ContainsKey(chunk.ItemId))
                                throw new InvalidDataException("Duplicate comparison state chunk ID: " + chunk.ItemId);
                            chunksByItemId.Add(chunk.ItemId, chunk);
                        }
                        else
                        {
                            throw new InvalidDataException(
                                "Unknown element in the comparison Custom XML namespace: " + root.LocalName);
                        }
                    }

                    if (manifest == null)
                        throw new InvalidDataException("No Excel Power Tools comparison manifest was found.");
                    if (manifest.Chunks.Count == 0)
                        throw new InvalidDataException("The comparison manifest contains no state chunks.");
                    if (chunksByItemId.Count != manifest.Chunks.Count)
                        throw new InvalidDataException("The manifest and package state chunk counts do not match.");

                    var projectionMetrics = new ReviewWorkbookProjectionValidationMetrics(
                        manifest.ProjectionDigestVersion);
                    string projectionSha256 = ComputeProjectionSha256(
                        workbookPart,
                        manifest.ProjectionDigestVersion,
                        projectionMetrics);
                    if (!FixedTimeEquals(projectionSha256, manifest.ProjectionSha256))
                        throw new InvalidDataException(
                            "The visible review workbook projection differs from its embedded comparison state.");

                    ValidateDecisionBinding(manifest, decision);

                    var orderedItemIds = new List<string>(manifest.Chunks.Count);
                    long totalLength = 0;
                    for (int sequence = 0; sequence < manifest.Chunks.Count; sequence++)
                    {
                        ManifestChunkData expected = manifest.Chunks[sequence];
                        if (expected.Sequence != sequence)
                            throw new InvalidDataException("State chunk sequences must be unique and contiguous from zero.");
                        if (!chunksByItemId.TryGetValue(expected.ItemId, out ChunkPartData actual))
                            throw new InvalidDataException("A manifest state chunk is missing: " + expected.ItemId);
                        if (actual.Sequence != expected.Sequence || actual.Length != expected.Length ||
                            !FixedTimeEquals(actual.Sha256, expected.Sha256))
                            throw new InvalidDataException("State chunk metadata differs from the manifest: " + expected.ItemId);
                        if (!string.Equals(actual.ProjectId, manifest.ProjectId, StringComparison.Ordinal))
                            throw new InvalidDataException("A state chunk belongs to a different comparison project.");

                        orderedItemIds.Add(FormatItemId(actual.ItemId));
                        totalLength = checked(totalLength + actual.Length);
                    }

                    if (totalLength != manifest.TotalLength)
                        throw new InvalidDataException("The reconstructed state length differs from the manifest.");
                    if (totalLength > MaxEmbeddedStateBytes)
                        throw new InvalidDataException("The embedded state exceeds the safe review-workbook limit.");

                    var stateBytes = new byte[(int)totalLength];
                    int offset = 0;
                    for (int sequence = 0; sequence < manifest.Chunks.Count; sequence++)
                    {
                        ManifestChunkData expected = manifest.Chunks[sequence];
                        ChunkPartData actual = chunksByItemId[expected.ItemId];
                        DecodeStateChunk(actual, expected, stateBytes, offset);
                        offset += expected.Length;
                    }

                    string reconstructedSha256 = ComputeSha256(stateBytes);
                    if (!FixedTimeEquals(reconstructedSha256, manifest.StateSha256))
                        throw new InvalidDataException("The reconstructed state SHA-256 differs from the manifest.");

                    string stateJson;
                    try
                    {
                        stateJson = new UTF8Encoding(false, true).GetString(stateBytes);
                    }
                    catch (DecoderFallbackException exception)
                    {
                        throw new InvalidDataException("The embedded state is not valid UTF-8.", exception);
                    }

                    return new ReviewWorkbookReadResult(
                        manifest.ProjectId,
                        stateJson,
                        FormatItemId(manifest.ItemId),
                        manifest.StateSha256,
                        manifest.ProjectionSha256,
                        new ReadOnlyCollection<string>(orderedItemIds),
                        decision?.Generation,
                        decision == null ? "" : FormatItemId(decision.ItemId),
                        decision?.Sha256 ?? "",
                        projectionMetrics);
                }
            }
            finally
            {
                if (packageStream.CanSeek)
                {
                    try { packageStream.Position = originalPosition; }
                    catch { }
                }
            }
        }

        private static void PreflightReviewPackage(Stream packageStream)
        {
            if (packageStream == null) throw new ArgumentNullException(nameof(packageStream));
            if (!packageStream.CanRead || !packageStream.CanSeek)
                throw new InvalidDataException("The review workbook package must be a readable, seekable stream.");
            if (packageStream.Length > MaxReviewPackageBytes)
                throw new InvalidDataException("The review workbook package exceeds the safe source-size limit.");

            ZipEnvelope envelope = ReadZipEnvelope(packageStream);
            if (envelope.EntryCount > MaxReviewPackageEntries ||
                envelope.CentralDirectorySize > MaxCentralDirectoryBytes)
                throw new InvalidDataException(
                    "The review workbook ZIP central directory exceeds the safe inspection limit.");

            packageStream.Position = 0;
            using (var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true))
            {
                if (archive.Entries.Count != envelope.EntryCount)
                    throw new InvalidDataException("The review workbook ZIP entry count is inconsistent.");

                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string name = (entry.FullName ?? "").Replace('\\', '/');
                    if (name.Length == 0 || !names.Add(name))
                        throw new InvalidDataException(
                            "The review workbook ZIP contains an empty, duplicate, or case-colliding entry name.");
                    if (entry.Length < 0 || entry.Length > MaxStreamingWorksheetCharacters)
                        throw new InvalidDataException(
                            "A review workbook ZIP entry exceeds the safe expansion limit: " + name);
                }

                IReadOnlyDictionary<string, long> relationshipBoundLargeParts =
                    GetRelationshipBoundLargeSpreadsheetEntries(archive);
                long totalLength = 0;
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    string name = (entry.FullName ?? "").Replace('\\', '/');
                    long entryLimit = relationshipBoundLargeParts.TryGetValue(
                        name,
                        out long largePartLimit)
                        ? largePartLimit
                        : MaxEntryUncompressedBytes;
                    if (entry.Length > entryLimit)
                        throw new InvalidDataException(
                            "A review workbook ZIP entry exceeds its content-type and relationship-bound " +
                            "expansion limit: " + name);
                    if (entry.Length > MaxDeclaredUncompressedBytes - totalLength)
                        throw new InvalidDataException(
                            "The review workbook ZIP declares an excessive aggregate expansion size.");
                    totalLength += entry.Length;
                }

                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.Length == 0) continue;
                    if (entry.CompressedLength <= 0)
                        throw new InvalidDataException(
                            "A non-empty review workbook ZIP entry has no compressed payload: " + entry.FullName);
                    if (entry.Length >= CompressionRatioCheckThresholdBytes &&
                        (double)entry.Length / entry.CompressedLength > MaxCompressionRatio)
                        throw new InvalidDataException(
                            "A review workbook ZIP entry has an unsafe compression ratio: " + entry.FullName);
                }
            }
        }

        private static ZipEnvelope ReadZipEnvelope(Stream packageStream)
        {
            const uint eocdSignature = 0x06054B50U;
            if (packageStream.Length < 22)
                throw new InvalidDataException("The review workbook is not a complete ZIP package.");

            int tailLength = (int)Math.Min(packageStream.Length, 65557L);
            var tail = new byte[tailLength];
            packageStream.Position = packageStream.Length - tailLength;
            int totalRead = 0;
            while (totalRead < tail.Length)
            {
                int read = packageStream.Read(tail, totalRead, tail.Length - totalRead);
                if (read == 0)
                    throw new InvalidDataException("The review workbook ZIP trailer is truncated.");
                totalRead += read;
            }

            for (int offset = tail.Length - 22; offset >= 0; offset--)
            {
                if (ReadUInt32(tail, offset) != eocdSignature) continue;
                ushort commentLength = ReadUInt16(tail, offset + 20);
                if ((long)offset + 22 + commentLength != tail.Length) continue;
                if (ReadUInt16(tail, offset + 4) != 0 || ReadUInt16(tail, offset + 6) != 0)
                    throw new InvalidDataException("Multi-disk review workbook ZIP packages are unsupported.");

                ushort entriesOnDisk = ReadUInt16(tail, offset + 8);
                ushort entryCount = ReadUInt16(tail, offset + 10);
                uint centralDirectorySize = ReadUInt32(tail, offset + 12);
                uint centralDirectoryOffset = ReadUInt32(tail, offset + 16);
                if (entriesOnDisk == ushort.MaxValue || entryCount == ushort.MaxValue ||
                    centralDirectorySize == uint.MaxValue || centralDirectoryOffset == uint.MaxValue)
                    throw new InvalidDataException(
                        "ZIP64 review workbook packages exceed the current safe envelope.");

                long eocdOffset = packageStream.Length - tail.Length + offset;
                if (entriesOnDisk != entryCount ||
                    (long)centralDirectoryOffset + centralDirectorySize > eocdOffset)
                    throw new InvalidDataException(
                        "The review workbook ZIP central directory is inconsistent.");
                return new ZipEnvelope(entryCount, centralDirectorySize);
            }

            throw new InvalidDataException("The review workbook is not a complete ZIP package.");
        }

        private static ushort ReadUInt16(byte[] bytes, int offset)
        {
            if (offset < 0 || offset > bytes.Length - 2)
                throw new InvalidDataException("The review workbook ZIP header is truncated.");
            return (ushort)(bytes[offset] | bytes[offset + 1] << 8);
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            if (offset < 0 || offset > bytes.Length - 4)
                throw new InvalidDataException("The review workbook ZIP header is truncated.");
            return (uint)(bytes[offset] |
                          bytes[offset + 1] << 8 |
                          bytes[offset + 2] << 16 |
                          bytes[offset + 3] << 24);
        }

        private static void ValidateDecisionGenerationForWrite(ReviewWorkbookState state)
        {
            ReviewDecisionGeneration generation = state.DecisionGeneration;
            if (generation == null) return;
            ReviewDecisionCatalog catalog = generation.Catalog;
            ReviewEvidenceGeneration evidence = catalog.EvidenceGeneration;
            if (evidence == null)
                throw new InvalidDataException(
                    "A durable decision state requires the exact retained evidence generation.");
            if (!string.Equals(catalog.ProjectId, state.ProjectId, StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The decision state belongs to a different comparison project.");
            if (string.IsNullOrWhiteSpace(state.DecisionEvidenceGenerationId) ||
                string.IsNullOrWhiteSpace(state.DecisionMappingGenerationId) ||
                !string.Equals(
                    state.DecisionEvidenceGenerationId,
                    catalog.EvidenceGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    state.DecisionMappingGenerationId,
                    catalog.MappingGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The decision catalog does not match the authoritative evidence and mapping generations " +
                    "supplied with the project state.");
            ValidateDecisionAttributeForWrite(catalog.ProjectId, MaxDecisionIdentityCharacters, false);
            ValidateDecisionAttributeForWrite(catalog.CatalogId, MaxDecisionIdentityCharacters, false);
            ValidateDecisionAttributeForWrite(catalog.BaseGenerationId, MaxDecisionIdentityCharacters, false);
            ValidateDecisionAttributeForWrite(catalog.MappingGenerationId, MaxDecisionIdentityCharacters, false);
            ValidateDecisionAttributeForWrite(catalog.MasterRevisionId, MaxDecisionIdentityCharacters, false);
            ValidateDecisionAttributeForWrite(generation.GenerationId, MaxDecisionIdentityCharacters, false);
            BuildDecisionReplayRecordsForWrite(generation);
            if (catalog.Subjects.Any(subject => !subject.HasExplicitMetadata))
                throw new InvalidDataException(
                    "A durable decision catalog requires explicit subject aspect and policy metadata.");
            foreach (ReviewEvidenceRevisionIdentity revision in evidence.Revisions)
                ValidateDecisionAttributeForWrite(
                    revision.RevisionId,
                    MaxDecisionIdentityCharacters,
                    false);
            foreach (ReviewDecisionSubject subject in catalog.Subjects)
            {
                ValidateDecisionAttributeForWrite(subject.SubjectId, MaxDecisionIdentityCharacters, false);
                ValidateDecisionAttributeForWrite(
                    subject.LogicalSubjectKey,
                    MaxDecisionIdentityCharacters,
                    false);
            }
            foreach (Observation observation in catalog.Observations)
            {
                ValidateDecisionAttributeForWrite(
                    observation.ObservationId,
                    MaxDecisionIdentityCharacters,
                    false);
                ValidateDecisionAttributeForWrite(
                    observation.SubjectId,
                    MaxDecisionIdentityCharacters,
                    false);
                ValidateDecisionAttributeForWrite(
                    observation.RevisionId,
                    MaxDecisionIdentityCharacters,
                    false);
                ValidateDecisionAttributeForWrite(
                    observation.ContentHash,
                    MaxDecisionIdentityCharacters,
                    false);
                ValidateDecisionAttributeForWrite(
                    observation.Provenance.LogicalSheetId,
                    MaxDecisionIdentityCharacters,
                    true);
                ValidateDecisionAttributeForWrite(
                    observation.Provenance.SourceLabel,
                    MaxDecisionLabelCharacters,
                    true);
                ValidatePortableDecisionSourceLabel(observation.Provenance.SourceLabel);
                ValidateDecisionAttributeForWrite(
                    observation.Provenance.SheetName,
                    MaxDecisionLabelCharacters,
                    true);
                ValidateDecisionAttributeForWrite(
                    observation.Provenance.Address,
                    MaxDecisionLabelCharacters,
                    true);
            }
            foreach (ReviewDecisionGroup group in catalog.Groups)
                ValidateDecisionAttributeForWrite(group.GroupId, MaxDecisionIdentityCharacters, false);

            var evidenceRevisionIds = new HashSet<string>(
                evidence.Revisions.Select(item => item.RevisionId),
                StringComparer.Ordinal);
            long expectedObservationCount = checked(
                (long)catalog.Subjects.Count * evidence.Revisions.Count);
            var observationPairs = new HashSet<string>(StringComparer.Ordinal);
            if (expectedObservationCount != catalog.Observations.Count ||
                catalog.Observations.Any(item =>
                    !evidenceRevisionIds.Contains(item.RevisionId) ||
                    !observationPairs.Add(item.SubjectId + "\u001F" + item.RevisionId)))
                throw new InvalidDataException(
                    "A durable decision catalog requires one observation for every subject and evidence revision.");
            long membershipTotal = catalog.Groups.Sum(group => (long)group.SubjectIds.Count);
            if (membershipTotal > MaxDecisionMembershipCount)
                throw new InvalidDataException(
                    "The decision group-membership count exceeds the safe limit.");
            int membershipCount = (int)membershipTotal;
            ValidateDecisionCounts(
                evidence.Revisions.Count,
                catalog.Subjects.Count,
                catalog.Observations.Count,
                catalog.Groups.Count,
                membershipCount);
        }

        private static void ValidateDecisionAttributeForWrite(
            string value,
            int maximumCharacters,
            bool allowEmpty)
        {
            value = value ?? "";
            if ((!allowEmpty && string.IsNullOrWhiteSpace(value)) || value.Length > maximumCharacters)
                throw new InvalidDataException(
                    "A decision-state XML attribute is blank or exceeds its safe limit.");
            try
            {
                XmlConvert.VerifyXmlChars(value);
            }
            catch (XmlException exception)
            {
                throw new InvalidDataException(
                    "A decision-state XML attribute contains an unsupported XML character.",
                    exception);
            }
        }

        private static void ValidatePortableDecisionSourceLabel(string sourceLabel)
        {
            if (string.IsNullOrEmpty(sourceLabel)) return;
            if (Path.IsPathRooted(sourceLabel) ||
                sourceLabel.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
                sourceLabel.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
                throw new InvalidDataException(
                    "A portable decision source label cannot contain an absolute or directory path.");
        }

        private static void ValidateDecisionCounts(
            int revisionCount,
            int subjectCount,
            int observationCount,
            int groupCount,
            int membershipCount)
        {
            if (revisionCount <= 0 || revisionCount > MaxDecisionRevisionCount)
                throw new InvalidDataException("The decision evidence revision count exceeds the safe limit.");
            if (subjectCount < 0 || subjectCount > MaxDecisionSubjectCount)
                throw new InvalidDataException("The decision subject count exceeds the safe limit.");
            if (observationCount < 0 || observationCount > MaxDecisionObservationCount)
                throw new InvalidDataException("The decision observation count exceeds the safe limit.");
            if (groupCount < 0 || groupCount > MaxDecisionGroupCount)
                throw new InvalidDataException("The decision group count exceeds the safe limit.");
            if (membershipCount < 0 || membershipCount > MaxDecisionMembershipCount)
                throw new InvalidDataException("The decision group-membership count exceeds the safe limit.");
        }

        private static IReadOnlyList<ReviewDecisionReplayRecord> BuildDecisionReplayRecordsForWrite(
            ReviewDecisionGeneration generation)
        {
            IReadOnlyList<ReviewDecisionReplayRecord> records;
            try
            {
                records = ReviewDecisionReplayEngine.Capture(generation);
                ReviewDecisionGeneration replayed = ReviewDecisionReplayEngine.Replay(
                    generation.Catalog,
                    records,
                    generation.GenerationId,
                    generation.PriorGenerationId);
                ValidateDecisionMaterializedStateMatchesReplay(generation, replayed);
            }
            catch (ReviewDecisionReplayException exception)
            {
                throw new InvalidDataException(
                    "The decision generation cannot be persisted because its event history does not replay: " +
                    exception.Code,
                    exception);
            }

            if (records.Count > MaxDecisionEventCount)
                throw new InvalidDataException("The decision event count exceeds the safe limit.");
            if ((records.Count == 0) != string.IsNullOrEmpty(generation.PriorGenerationId))
                throw new InvalidDataException(
                    "The decision generation root/non-root envelope is inconsistent.");
            foreach (ReviewDecisionReplayRecord record in records)
                ValidateDecisionReplayRecordForWrite(record);
            return records;
        }

        private static void ValidateDecisionMaterializedStateMatchesReplay(
            ReviewDecisionGeneration supplied,
            ReviewDecisionGeneration replayed)
        {
            if (supplied.SubjectStates.Count != replayed.SubjectStates.Count ||
                supplied.ActiveDecisions.Count != replayed.ActiveDecisions.Count)
                throw new InvalidDataException(
                    "The materialized decision state does not match its deterministic event replay.");
            foreach (KeyValuePair<string, ReviewDecisionState> pair in supplied.SubjectStates)
            {
                if (!replayed.SubjectStates.TryGetValue(
                        pair.Key,
                        out ReviewDecisionState replayedState) ||
                    pair.Value.Status != replayedState.Status ||
                    !string.Equals(
                        pair.Value.SelectedObservationId,
                        replayedState.SelectedObservationId,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        pair.Value.EffectiveEventId,
                        replayedState.EffectiveEventId,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The materialized decision state does not match its deterministic event replay.");
            }
            foreach (KeyValuePair<string, DecisionEvent> pair in supplied.ActiveDecisions)
            {
                if (!replayed.ActiveDecisions.TryGetValue(
                        pair.Key,
                        out DecisionEvent replayedEvent) ||
                    !string.Equals(
                        pair.Value.EventId,
                        replayedEvent.EventId,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The active decision index does not match its deterministic event replay.");
            }
        }

        private static void ValidateDecisionReplayRecordForWrite(ReviewDecisionReplayRecord record)
        {
            if (record == null)
                throw new InvalidDataException("A persisted decision replay record is required.");
            if (record.ContractVersion != ReviewDecisionEventContractVersions.V1)
                throw new InvalidDataException("The decision event contract version is unsupported.");
            ValidateDecisionAttributeForWrite(
                record.PriorGenerationId,
                MaxDecisionIdentityCharacters,
                false);
            ValidateDecisionAttributeForWrite(
                record.NewGenerationId,
                MaxDecisionIdentityCharacters,
                false);

            DecisionEvent decisionEvent = record.DecisionEvent;
            if (decisionEvent == null)
                throw new InvalidDataException("A persisted decision replay record has no event.");
            ValidateDecisionAttributeForWrite(
                decisionEvent.EventId,
                MaxDecisionIdentityCharacters,
                false);
            ValidateDecisionAttributeForWrite(
                decisionEvent.RequestId,
                MaxDecisionIdentityCharacters,
                false);
            ValidateDecisionAttributeForWrite(
                decisionEvent.ProjectId,
                MaxDecisionIdentityCharacters,
                false);
            ValidateDecisionAttributeForWrite(
                decisionEvent.EvidenceGenerationId,
                MaxDecisionIdentityCharacters,
                false);
            ValidateDecisionAttributeForWrite(
                decisionEvent.MappingGenerationId,
                MaxDecisionIdentityCharacters,
                false);
            ValidateDecisionAttributeForWrite(
                decisionEvent.RunId,
                MaxDecisionIdentityCharacters,
                true);
            ValidateDecisionAttributeForWrite(
                decisionEvent.SubjectId,
                MaxDecisionIdentityCharacters,
                true);
            ValidateDecisionAttributeForWrite(
                decisionEvent.GroupId,
                MaxDecisionIdentityCharacters,
                true);
            ValidateDecisionAttributeForWrite(
                decisionEvent.SelectedObservationId,
                MaxDecisionIdentityCharacters,
                true);
            ValidateDecisionAttributeForWrite(
                decisionEvent.SupersededEventId,
                MaxDecisionIdentityCharacters,
                true);

            ValidateDecisionEventCollectionCount(
                decisionEvent.AffectedSubjectIds.Count,
                MaxDecisionAffectedSubjectCountPerEvent,
                "affected-subject");
            ValidateDecisionEventCollectionCount(
                decisionEvent.SupportingRevisionIds.Count,
                MaxDecisionSupportingRevisionCountPerEvent,
                "supporting-revision");
            ValidateDecisionEventCollectionCount(
                decisionEvent.SupportingContentHashes.Count,
                MaxDecisionSupportingContentHashCountPerEvent,
                "supporting-content-hash");
            ValidateDecisionEventCollectionCount(
                decisionEvent.ReviewerSnapshot.IdentitySourceLabels.Count,
                MaxDecisionReviewerSourceCountPerEvent,
                "reviewer-source");
            ValidateDecisionEventCollectionCount(
                decisionEvent.Locations.Count,
                MaxDecisionLocationCountPerEvent,
                "location");

            foreach (string subjectId in decisionEvent.AffectedSubjectIds)
                ValidateDecisionAttributeForWrite(
                    subjectId,
                    MaxDecisionIdentityCharacters,
                    false);
            foreach (string revisionId in decisionEvent.SupportingRevisionIds)
                ValidateDecisionAttributeForWrite(
                    revisionId,
                    MaxDecisionIdentityCharacters,
                    false);
            foreach (string contentHash in decisionEvent.SupportingContentHashes)
            {
                if (contentHash == null || contentHash.Length != 64 || contentHash.Any(ch => !IsHex(ch)))
                    throw new InvalidDataException(
                        "A decision event supporting content hash is not SHA-256.");
            }
            foreach (DecisionLocationContext location in decisionEvent.Locations)
            {
                ValidateDecisionAttributeForWrite(
                    location.Role,
                    MaxDecisionIdentityCharacters,
                    false);
                ValidateDecisionAttributeForWrite(
                    location.RevisionId,
                    MaxDecisionIdentityCharacters,
                    false);
                DecisionPayloadBytes(location.LogicalSheetId);
                DecisionPayloadBytes(location.SheetName);
                DecisionPayloadBytes(location.Address);
                DecisionPayloadBytes(location.SourceLabel);
                ValidatePortableDecisionSourceLabel(location.SourceLabel);
            }

            DecisionPayloadBytes(decisionEvent.ReviewerSnapshot.ConfirmedDisplayName);
            DecisionPayloadBytes(decisionEvent.ReviewerSnapshot.OfficeDisplayName);
            foreach (string source in decisionEvent.ReviewerSnapshot.IdentitySourceLabels)
                DecisionPayloadBytes(source);
            DecisionPayloadBytes(decisionEvent.AddInBuild);
            DecisionPayloadBytes(decisionEvent.Note);
        }

        private static void ValidateDecisionEventCollectionCount(
            int count,
            int maximum,
            string label)
        {
            if (count < 0 || count > maximum)
                throw new InvalidDataException(
                    "The decision event " + label + " count exceeds the safe limit.");
        }

        private static void BuildPackage(
            Stream output,
            ReviewWorkbookState state,
            IReadOnlyList<StateChunk> chunks,
            string manifestItemId,
            string decisionItemId,
            string totalSha256)
        {
            var worksheetBudget = new ProjectionWriteBudget(MaxDeclaredUncompressedBytes);
            using (var boundedOutput = new BoundedPackageWriteStream(
                       output,
                       MaxReviewPackageBytes,
                       leaveOpen: true))
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(
                boundedOutput,
                SpreadsheetDocumentType.Workbook,
                true))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                for (int index = 0; index < ExpectedSheetNames.Count; index++)
                {
                    string sheetName = ExpectedSheetNames[index];
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    if (sheetName == "Overview")
                        WriteOverviewSheet(
                            worksheetPart,
                            state.ProjectId,
                            totalSha256,
                            state.ComparisonStatus ?? "",
                            state.RevisionRows.Count,
                            state.ChangeRows.Count,
                            state.DiagnosticRows.Count,
                            worksheetBudget);
                    else if (sheetName == "Revisions")
                        WriteRevisionsSheet(worksheetPart, state.RevisionRows, worksheetBudget);
                    else if (sheetName == "Changes")
                        WriteChangesSheet(worksheetPart, state.ChangeRows, worksheetBudget);
                    else if (sheetName == "Details")
                        WriteDetailsSheet(
                            worksheetPart,
                            state.DetailBefore ?? "",
                            state.DetailAfter ?? "",
                            state.DetailUnavailableReason ?? "",
                            worksheetBudget);
                    else if (sheetName == "Diagnostics")
                        WriteDiagnosticsSheet(worksheetPart, state.DiagnosticRows, worksheetBudget);
                    else if (sheetName == "Legend")
                        WriteLegendSheet(worksheetPart, worksheetBudget);
                    else
                        WriteTitleSheet(worksheetPart, sheetName, worksheetBudget);

                    sheets.Append(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (uint)(index + 1),
                        Name = sheetName,
                        State = SheetStateValues.Visible,
                    });
                }

                workbookPart.Workbook.Save();
                string projectionSha256 = ComputeProjectionSha256(
                    workbookPart,
                    StreamingProjectionDigestVersion);

                // Deliberately create chunks before the manifest. Readers must discover parts by marker and
                // persistent item ID, not by relationship order or package part name.
                foreach (StateChunk chunk in chunks)
                    WriteStateChunkPart(workbookPart, state.ProjectId, chunk);
                string decisionSha256 = "";
                if (state.DecisionGeneration != null)
                {
                    decisionSha256 = WriteDecisionStatePart(
                        workbookPart,
                        state.DecisionGeneration,
                        decisionItemId,
                        manifestItemId,
                        totalSha256,
                        projectionSha256);
                }
                WriteManifestPart(
                    workbookPart,
                    state.ProjectId,
                    manifestItemId,
                    chunks,
                    totalSha256,
                    projectionSha256,
                    decisionItemId,
                    decisionSha256,
                    state.DecisionEvidenceGenerationId,
                    state.DecisionMappingGenerationId);
            }
        }

        private static void WriteOverviewSheet(
            WorksheetPart part,
            string projectId,
            string stateSha256,
            string comparisonStatus,
            int revisionCount,
            int changeCount,
            int diagnosticCount,
            ProjectionWriteBudget budget)
        {
            var widths = new ContentWidthAccumulator();
            widths.Add("Spreadsheet comparison review");
            widths.Add("Project", projectId);
            widths.Add("Status", comparisonStatus);
            widths.Add("State SHA-256", stateSha256);
            widths.Add("Navigation");
            widths.Add("Open Changes", changeCount.ToString(CultureInfo.InvariantCulture));
            widths.Add("Open Revisions", revisionCount.ToString(CultureInfo.InvariantCulture));
            widths.Add("Open Diagnostics", diagnosticCount.ToString(CultureInfo.InvariantCulture));
            widths.Add("Open Legend");
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);
                WriteRow(writer, 1, "Spreadsheet comparison review");
                WriteRow(writer, 2, "Project", projectId);
                WriteRow(writer, 3, "Status", comparisonStatus);
                WriteRow(writer, 4, "State SHA-256", stateSha256);
                WriteRow(writer, 6, "Navigation");
                WriteRow(writer, 7, "Open Changes", changeCount.ToString(CultureInfo.InvariantCulture));
                WriteRow(writer, 8, "Open Revisions", revisionCount.ToString(CultureInfo.InvariantCulture));
                WriteRow(writer, 9, "Open Diagnostics", diagnosticCount.ToString(CultureInfo.InvariantCulture));
                WriteRow(writer, 10, "Open Legend");
                writer.WriteEndElement();

                writer.WriteStartElement("hyperlinks", SpreadsheetNamespace);
                WriteInternalHyperlink(writer, "A7", "'Changes'!A1", "Open the Changes sheet");
                WriteInternalHyperlink(writer, "A8", "'Revisions'!A1", "Open the Revisions sheet");
                WriteInternalHyperlink(writer, "A9", "'Diagnostics'!A1", "Open the Diagnostics sheet");
                WriteInternalHyperlink(writer, "A10", "'Legend'!A1", "Open the Legend sheet");
                writer.WriteEndElement();

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static void WriteInternalHyperlink(
            XmlWriter writer,
            string reference,
            string location,
            string tooltip)
        {
            writer.WriteStartElement("hyperlink", SpreadsheetNamespace);
            writer.WriteAttributeString("ref", reference);
            writer.WriteAttributeString("location", location);
            writer.WriteAttributeString("tooltip", tooltip);
            writer.WriteEndElement();
        }

        private static void WriteTitleSheet(
            WorksheetPart part,
            string title,
            ProjectionWriteBudget budget)
        {
            var widths = new ContentWidthAccumulator();
            widths.Add(title);
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);
                WriteRow(writer, 1, title);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static void WriteChangesSheet(
            WorksheetPart part,
            IList<ReviewChangeProjectionRow> rows,
            ProjectionWriteBudget budget)
        {
            var headers = new[]
            {
                "Change ID",
                "Sheet",
                "Address",
                "Aspect",
                "Before",
                "After",
                "Group ID",
                "Flags",
                "Revisions",
                "Distinct states",
                "Alignment confidence",
            };
            var widths = new ContentWidthAccumulator();
            widths.Add(headers);
            foreach (ReviewChangeProjectionRow row in rows)
            {
                if (row == null)
                    throw new InvalidDataException("A visible change projection row cannot be null.");
                widths.Add(
                    ProjectCellText(row.ChangeId),
                    ProjectCellText(row.SheetName),
                    ProjectCellText(row.Address),
                    ProjectCellText(row.Aspect),
                    ProjectCellText(row.BeforeText),
                    ProjectCellText(row.AfterText),
                    ProjectCellText(row.GroupId),
                    ProjectCellText(row.Flags),
                    row.RevisionCount.ToString(CultureInfo.InvariantCulture),
                    row.DistinctStateCount.ToString(CultureInfo.InvariantCulture),
                    ProjectCellText(row.AlignmentConfidence));
            }
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);
                WriteRow(writer, 1, headers);

                for (int index = 0; index < rows.Count; index++)
                {
                    ReviewChangeProjectionRow row = rows[index];
                    if (row == null)
                        throw new InvalidDataException("A visible change projection row cannot be null.");
                    uint rowIndex = checked((uint)index + 2U);
                    WriteRow(
                        writer,
                        rowIndex,
                        ProjectCellText(row.ChangeId),
                        ProjectCellText(row.SheetName),
                        ProjectCellText(row.Address),
                        ProjectCellText(row.Aspect),
                        ProjectCellText(row.BeforeText),
                        ProjectCellText(row.AfterText),
                        ProjectCellText(row.GroupId),
                        ProjectCellText(row.Flags),
                        row.RevisionCount.ToString(CultureInfo.InvariantCulture),
                        row.DistinctStateCount.ToString(CultureInfo.InvariantCulture),
                        ProjectCellText(row.AlignmentConfidence));
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static void WriteRevisionsSheet(
            WorksheetPart part,
            IList<ReviewRevisionProjectionRow> rows,
            ProjectionWriteBudget budget)
        {
            var headers = new[]
            {
                "Revision ID",
                "Label",
                "Order",
                "Parent revision ID",
                "Role",
                "Source file",
                "Original source path",
                "Byte content ID",
                "Modified UTC",
                "Inference evidence",
                "Confidence",
            };
            var widths = new ContentWidthAccumulator();
            widths.Add(headers);
            foreach (ReviewRevisionProjectionRow row in rows)
            {
                if (row == null)
                    throw new InvalidDataException("A visible revision projection row cannot be null.");
                widths.Add(
                    ProjectCellText(row.RevisionId),
                    ProjectCellText(row.Label),
                    row.Order.ToString(CultureInfo.InvariantCulture),
                    ProjectCellText(row.ParentRevisionId),
                    ProjectCellText(row.Role),
                    ProjectCellText(row.SourceBaseName),
                    ProjectCellText(row.SourcePath),
                    ProjectCellText(row.ByteContentId),
                    ProjectCellText(row.ModifiedUtc),
                    ProjectCellText(row.Inference),
                    ProjectCellText(row.Confidence));
            }
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);
                WriteRow(writer, 1, headers);

                for (int index = 0; index < rows.Count; index++)
                {
                    ReviewRevisionProjectionRow row = rows[index];
                    if (row == null)
                        throw new InvalidDataException("A visible revision projection row cannot be null.");
                    WriteRow(
                        writer,
                        checked((uint)index + 2U),
                        ProjectCellText(row.RevisionId),
                        ProjectCellText(row.Label),
                        row.Order.ToString(CultureInfo.InvariantCulture),
                        ProjectCellText(row.ParentRevisionId),
                        ProjectCellText(row.Role),
                        ProjectCellText(row.SourceBaseName),
                        ProjectCellText(row.SourcePath),
                        ProjectCellText(row.ByteContentId),
                        ProjectCellText(row.ModifiedUtc),
                        ProjectCellText(row.Inference),
                        ProjectCellText(row.Confidence));
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static void WriteDiagnosticsSheet(
            WorksheetPart part,
            IList<ReviewDiagnosticProjectionRow> rows,
            ProjectionWriteBudget budget)
        {
            var headers = new[]
            {
                "Finding UID",
                "Code",
                "Severity",
                "Revision ID",
                "Summary",
                "Full detail",
                "Suggested action",
            };
            var widths = new ContentWidthAccumulator();
            widths.Add(headers);
            foreach (ReviewDiagnosticProjectionRow row in rows)
            {
                if (row == null)
                    throw new InvalidDataException("A visible diagnostic projection row cannot be null.");
                widths.Add(
                    ProjectCellText(row.FindingUid),
                    ProjectCellText(row.Code),
                    ProjectCellText(row.Severity),
                    ProjectCellText(row.RevisionId),
                    ProjectCellText(row.Summary),
                    ProjectCellText(row.Detail),
                    ProjectCellText(row.SuggestedAction));
            }
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);
                WriteRow(writer, 1, headers);

                for (int index = 0; index < rows.Count; index++)
                {
                    ReviewDiagnosticProjectionRow row = rows[index];
                    if (row == null)
                        throw new InvalidDataException("A visible diagnostic projection row cannot be null.");
                    WriteRow(
                        writer,
                        checked((uint)index + 2U),
                        ProjectCellText(row.FindingUid),
                        ProjectCellText(row.Code),
                        ProjectCellText(row.Severity),
                        ProjectCellText(row.RevisionId),
                        ProjectCellText(row.Summary),
                        ProjectCellText(row.Detail),
                        ProjectCellText(row.SuggestedAction));
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static void WriteLegendSheet(
            WorksheetPart part,
            ProjectionWriteBudget budget)
        {
            var widths = new ContentWidthAccumulator();
            widths.Add("Comparison legend");
            widths.Add("Channel", "Meaning");
            widths.Add("Value", "Exact stored typed-value change");
            widths.Add("Calculated value", "Exact stored result last calculated by the workbook");
            widths.Add("Formula", "Exact stored formula text or formula metadata change");
            widths.Add("Structure", "Sheet, row, or column alignment context for the three content channels");
            widths.Add("Equivalent", "Quiet classification only; exact evidence remains visible");
            widths.Add("Reverted", "Master returned to an earlier exact state after an intermediate change");
            widths.Add("Conflict", "Sibling revisions changed the same logical subject differently");
            widths.Add("Branch-only", "Exact change exists outside the confirmed master lineage");
            widths.Add("Scope", "Only values, calculated values, and formulas are compared as cell content");
            widths.Add("Important", "Color is never the only signal; labels and identifiers are authoritative.");
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);
                WriteRow(writer, 1, "Comparison legend");
                WriteRow(writer, 3, "Channel", "Meaning");
                WriteRow(writer, 4, "Value", "Exact stored typed-value change");
                WriteRow(writer, 5, "Calculated value", "Exact stored result last calculated by the workbook");
                WriteRow(writer, 6, "Formula", "Exact stored formula text or formula metadata change");
                WriteRow(writer, 7, "Structure", "Sheet, row, or column alignment context for the three content channels");
                WriteRow(writer, 8, "Equivalent", "Quiet classification only; exact evidence remains visible");
                WriteRow(writer, 9, "Reverted", "Master returned to an earlier exact state after an intermediate change");
                WriteRow(writer, 10, "Conflict", "Sibling revisions changed the same logical subject differently");
                WriteRow(writer, 11, "Branch-only", "Exact change exists outside the confirmed master lineage");
                WriteRow(writer, 13, "Scope", "Only values, calculated values, and formulas are compared as cell content");
                WriteRow(writer, 14, "Important", "Color is never the only signal; labels and identifiers are authoritative.");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static void WriteDetailsSheet(
            WorksheetPart part,
            string before,
            string after,
            string unavailableReason,
            ProjectionWriteBudget budget)
        {
            string projectedBefore = string.IsNullOrEmpty(before) && string.IsNullOrEmpty(after)
                ? ""
                : ProjectCellText(before);
            string projectedAfter = string.IsNullOrEmpty(before) && string.IsNullOrEmpty(after)
                ? ""
                : ProjectCellText(after);
            var widths = new ContentWidthAccumulator();
            if (projectedBefore.Length == 0 && projectedAfter.Length == 0)
                widths.Add(ProjectCellText(unavailableReason));
            else
            {
                widths.Add("Character-level difference");
                widths.Add("Before", projectedBefore);
                widths.Add("After", projectedAfter);
            }
            using (XmlWriter writer = CreateWorksheetWriter(part, budget))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", SpreadsheetNamespace);
                WriteContentSizedColumns(writer, widths);
                writer.WriteStartElement("sheetData", SpreadsheetNamespace);

                if (string.IsNullOrEmpty(before) && string.IsNullOrEmpty(after))
                {
                    WriteRow(writer, 1, ProjectCellText(unavailableReason));
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                    return;
                }

                // The complete evidence remains in StateJson. Only this visible Excel-cell projection is
                // sanitized and, when necessary, visibly truncated to Excel's hard per-cell text limit.
                CharacterDiffResult diff = CharacterDiff.Compare(projectedBefore, projectedAfter);
                WriteRow(writer, 1, "Character-level difference");

                writer.WriteStartElement("row", SpreadsheetNamespace);
                writer.WriteAttributeString("r", "3");
                WriteInlineStringCell(writer, "A3", "Before");
                WriteRichDiffCell(writer, "B3", diff.Runs, false);
                writer.WriteEndElement();

                writer.WriteStartElement("row", SpreadsheetNamespace);
                writer.WriteAttributeString("r", "4");
                WriteInlineStringCell(writer, "A4", "After");
                WriteRichDiffCell(writer, "B4", diff.Runs, true);
                writer.WriteEndElement();

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static XmlWriter CreateWorksheetWriter(
            WorksheetPart part,
            ProjectionWriteBudget budget)
        {
            Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            var bounded = new BoundedWorksheetWriteStream(
                stream,
                MaxStreamingWorksheetCharacters,
                budget);
            return XmlWriter.Create(bounded, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                CloseOutput = true,
                CheckCharacters = true,
            });
        }

        private static void WriteContentSizedColumns(
            XmlWriter writer,
            ContentWidthAccumulator widths)
        {
            if (widths == null || widths.Count == 0) return;
            writer.WriteStartElement("cols", SpreadsheetNamespace);
            for (int index = 0; index < widths.Count; index++)
            {
                uint column = checked((uint)index + 1U);
                writer.WriteStartElement("col", SpreadsheetNamespace);
                writer.WriteAttributeString("min", column.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("max", column.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString(
                    "width",
                    widths.WidthAt(index).ToString("0.##", CultureInfo.InvariantCulture));
                writer.WriteAttributeString("bestFit", "1");
                writer.WriteAttributeString("customWidth", "1");
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private sealed class ContentWidthAccumulator
        {
            private readonly List<int> _maximumCharacters = new List<int>();

            internal int Count => _maximumCharacters.Count;

            internal void Add(params string[] values)
            {
                if (values == null) return;
                while (_maximumCharacters.Count < values.Length) _maximumCharacters.Add(0);
                for (int index = 0; index < values.Length; index++)
                    _maximumCharacters[index] = Math.Max(
                        _maximumCharacters[index],
                        VisibleLineLength(values[index]));
            }

            internal double WidthAt(int index) => Math.Min(
                MaxVisibleColumnWidth,
                Math.Max(
                    DefaultVisibleColumnWidth,
                    _maximumCharacters[index] + ContentWidthPaddingCharacters));

            private static int VisibleLineLength(string value)
            {
                if (string.IsNullOrEmpty(value)) return 0;
                int current = 0;
                int maximum = 0;
                int cap = checked((int)MaxVisibleColumnWidth - ContentWidthPaddingCharacters);
                foreach (char character in value)
                {
                    if (character == '\r' || character == '\n')
                    {
                        maximum = Math.Max(maximum, current);
                        current = 0;
                    }
                    else
                    {
                        current = Math.Min(cap, current + (character == '\t' ? 4 : 1));
                        maximum = Math.Max(maximum, current);
                    }
                    if (maximum >= cap) return cap;
                }
                return Math.Max(maximum, current);
            }
        }

        private static void WriteRow(XmlWriter writer, uint rowIndex, params string[] values)
        {
            writer.WriteStartElement("row", SpreadsheetNamespace);
            writer.WriteAttributeString("r", rowIndex.ToString(CultureInfo.InvariantCulture));
            for (int column = 0; column < values.Length; column++)
                WriteInlineStringCell(writer, ColumnName(column + 1) + rowIndex, values[column] ?? "");
            writer.WriteEndElement();
        }

        private static void WriteInlineStringCell(XmlWriter writer, string reference, string value)
        {
            writer.WriteStartElement("c", SpreadsheetNamespace);
            writer.WriteAttributeString("r", reference);
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is", SpreadsheetNamespace);
            writer.WriteStartElement("t", SpreadsheetNamespace);
            writer.WriteAttributeString("xml", "space", XmlNamespace, "preserve");
            writer.WriteString(SanitizeXmlText(value));
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteRichDiffCell(
            XmlWriter writer,
            string reference,
            IReadOnlyList<CharacterDiffRun> runs,
            bool after)
        {
            writer.WriteStartElement("c", SpreadsheetNamespace);
            writer.WriteAttributeString("r", reference);
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is", SpreadsheetNamespace);

            bool wroteRun = false;
            foreach (CharacterDiffRun run in runs)
            {
                if ((!after && run.Kind == CharacterDiffKind.Added) ||
                    (after && run.Kind == CharacterDiffKind.Removed))
                    continue;

                writer.WriteStartElement("r", SpreadsheetNamespace);
                writer.WriteStartElement("rPr", SpreadsheetNamespace);
                if (run.Kind == CharacterDiffKind.Added)
                {
                    writer.WriteStartElement("b", SpreadsheetNamespace);
                    writer.WriteEndElement();
                    WriteRunColor(writer, "FF2E7D32");
                }
                else if (run.Kind == CharacterDiffKind.Removed)
                {
                    writer.WriteStartElement("strike", SpreadsheetNamespace);
                    writer.WriteEndElement();
                    WriteRunColor(writer, "FFC62828");
                }
                else
                {
                    WriteRunColor(writer, "FF404040");
                }
                writer.WriteEndElement();
                writer.WriteStartElement("t", SpreadsheetNamespace);
                writer.WriteAttributeString("xml", "space", XmlNamespace, "preserve");
                string visibleText = run.Kind == CharacterDiffKind.Equal
                    ? run.Text
                    : CharacterDiff.MakeWhitespaceVisible(run.Text);
                writer.WriteString(SanitizeXmlText(visibleText));
                writer.WriteEndElement();
                writer.WriteEndElement();
                wroteRun = true;
            }

            if (!wroteRun)
            {
                writer.WriteStartElement("r", SpreadsheetNamespace);
                writer.WriteStartElement("t", SpreadsheetNamespace);
                writer.WriteString("");
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteRunColor(XmlWriter writer, string rgb)
        {
            writer.WriteStartElement("color", SpreadsheetNamespace);
            writer.WriteAttributeString("rgb", rgb);
            writer.WriteEndElement();
        }

        private static void WriteStateChunkPart(WorkbookPart workbookPart, string projectId, StateChunk chunk)
        {
            CustomXmlPart part = AddCustomXmlPart(workbookPart, chunk.ItemId);
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter writer = CreateCustomXmlWriter(stream))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("ept", "stateChunk", CustomXmlNamespace);
                WriteEnvelopeAttributes(writer, chunk.ItemId);
                writer.WriteAttributeString("projectId", projectId);
                writer.WriteAttributeString("sequence", chunk.Sequence.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("length", chunk.Length.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("sha256", chunk.Sha256);
                writer.WriteBase64(chunk.Payload, chunk.Offset, chunk.Length);
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static string WriteDecisionStatePart(
            WorkbookPart workbookPart,
            ReviewDecisionGeneration generation,
            string itemId,
            string manifestItemId,
            string stateSha256,
            string projectionSha256)
        {
            CustomXmlPart part = AddCustomXmlPart(
                workbookPart,
                itemId,
                DecisionCustomXmlNamespace);
            return WriteDecisionStatePart(
                part,
                generation,
                itemId,
                manifestItemId,
                stateSha256,
                projectionSha256);
        }

        /// <summary>
        /// Rewrites an existing decision Custom XML part in place. The caller owns package-level item-ID and
        /// manifest updates; this narrow helper is shared by the create-new and copy-on-write persistence paths.
        /// </summary>
        private static string WriteDecisionStatePart(
            CustomXmlPart part,
            ReviewDecisionGeneration generation,
            string itemId,
            string manifestItemId,
            string stateSha256,
            string projectionSha256)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            ReviewDecisionCatalog catalog = generation.Catalog;
            ReviewEvidenceGeneration evidence = catalog.EvidenceGeneration;
            IReadOnlyList<ReviewDecisionReplayRecord> replayRecords =
                BuildDecisionReplayRecordsForWrite(generation);
            string decisionSchemaVersion = replayRecords.Count == 0
                ? SchemaVersion
                : DecisionEventStateSchemaVersion;
            string digest;
            using (SHA256 sha256 = SHA256.Create())
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (var bounded = new BoundedWriteStream(stream, MaxDecisionXmlBytes))
            using (var hashing = new CryptoStream(bounded, sha256, CryptoStreamMode.Write))
            {
                using (XmlWriter writer = CreateCustomXmlWriter(hashing))
                {
                    writer.WriteStartDocument(true);
                    writer.WriteStartElement("eptd", "decisionState", DecisionCustomXmlNamespace);
                    writer.WriteAttributeString("marker", DecisionMarker);
                    writer.WriteAttributeString("schemaVersion", decisionSchemaVersion);
                    writer.WriteAttributeString("itemId", itemId);
                    writer.WriteAttributeString("projectId", catalog.ProjectId);
                    writer.WriteAttributeString("manifestItemId", manifestItemId);
                    writer.WriteAttributeString("stateSha256", stateSha256);
                    writer.WriteAttributeString("projectionSha256", projectionSha256);
                    writer.WriteAttributeString("catalogId", catalog.CatalogId);
                    writer.WriteAttributeString("evidenceGenerationId", evidence.GenerationId);
                    writer.WriteAttributeString("mappingGenerationId", catalog.MappingGenerationId);
                    writer.WriteAttributeString("masterRevisionId", catalog.MasterRevisionId);
                    writer.WriteAttributeString("generationId", generation.GenerationId);
                    writer.WriteAttributeString("priorGenerationId", generation.PriorGenerationId);
                    writer.WriteAttributeString("revisionCount", Invariant(evidence.Revisions.Count));
                    writer.WriteAttributeString("subjectCount", Invariant(catalog.Subjects.Count));
                    writer.WriteAttributeString("observationCount", Invariant(catalog.Observations.Count));
                    writer.WriteAttributeString("groupCount", Invariant(catalog.Groups.Count));
                    writer.WriteAttributeString(
                        "membershipCount",
                        Invariant(checked((int)catalog.Groups.Sum(
                            group => (long)group.SubjectIds.Count))));
                    writer.WriteAttributeString("eventCount", Invariant(replayRecords.Count));

                    writer.WriteStartElement("eptd", "evidence", DecisionCustomXmlNamespace);
                    writer.WriteAttributeString(
                        "effectiveMappingGenerationId",
                        evidence.EffectiveMappingGenerationId);
                    foreach (ReviewEvidenceRevisionIdentity revision in evidence.Revisions)
                    {
                        writer.WriteStartElement("eptd", "revision", DecisionCustomXmlNamespace);
                        writer.WriteAttributeString("revisionId", revision.RevisionId);
                        writer.WriteAttributeString("order", Invariant(revision.Order));
                        writer.WriteAttributeString("byteContentId", revision.ByteContentId);
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();

                    writer.WriteStartElement("eptd", "subjects", DecisionCustomXmlNamespace);
                    foreach (ReviewDecisionSubject subject in catalog.Subjects)
                    {
                        writer.WriteStartElement("eptd", "subject", DecisionCustomXmlNamespace);
                        writer.WriteAttributeString("subjectId", subject.SubjectId);
                        writer.WriteAttributeString("logicalSubjectKey", subject.LogicalSubjectKey);
                        writer.WriteAttributeString("aspect", Invariant((int)subject.Aspect));
                        writer.WriteAttributeString("policy", Invariant((int)subject.Policy));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();

                    writer.WriteStartElement("eptd", "observations", DecisionCustomXmlNamespace);
                    foreach (Observation observation in catalog.Observations)
                    {
                        ObservationProvenance provenance = observation.Provenance;
                        writer.WriteStartElement("eptd", "observation", DecisionCustomXmlNamespace);
                        writer.WriteAttributeString("observationId", observation.ObservationId);
                        writer.WriteAttributeString("subjectId", observation.SubjectId);
                        writer.WriteAttributeString("revisionId", observation.RevisionId);
                        writer.WriteAttributeString("exactType", Invariant((int)observation.ExactType));
                        writer.WriteAttributeString("contentHash", observation.ContentHash);
                        writer.WriteAttributeString(
                            "cachedResultTrust",
                            Invariant((int)observation.CachedResultTrust));
                        writer.WriteAttributeString("logicalSheetId", provenance.LogicalSheetId);
                        writer.WriteAttributeString("sourceLabel", provenance.SourceLabel);
                        writer.WriteAttributeString("sheetName", provenance.SheetName);
                        writer.WriteAttributeString("address", provenance.Address);
                        writer.WriteAttributeString("physicalRow", Invariant(provenance.PhysicalRow));
                        writer.WriteAttributeString("physicalColumn", Invariant(provenance.PhysicalColumn));

                        if (observation.ExactType == ObservationExactType.Formula)
                            WriteEncodedTextElement(writer, "formula", observation.Formula);
                        else
                            WriteScalarElement(writer, "value", observation.Value);
                        if (observation.CachedResult != null)
                            WriteScalarElement(writer, "cachedResult", observation.CachedResult);
                        if (observation.DisplayText.Length != 0)
                            WriteEncodedTextElement(writer, "displayText", observation.DisplayText);
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();

                    writer.WriteStartElement("eptd", "groups", DecisionCustomXmlNamespace);
                    foreach (ReviewDecisionGroup group in catalog.Groups)
                    {
                        writer.WriteStartElement("eptd", "group", DecisionCustomXmlNamespace);
                        writer.WriteAttributeString("groupId", group.GroupId);
                        foreach (string subjectId in group.SubjectIds)
                        {
                            writer.WriteStartElement("eptd", "member", DecisionCustomXmlNamespace);
                            writer.WriteAttributeString("subjectId", subjectId);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();

                    writer.WriteStartElement("eptd", "generation", DecisionCustomXmlNamespace);
                    writer.WriteAttributeString("generationId", generation.GenerationId);
                    writer.WriteAttributeString("priorGenerationId", generation.PriorGenerationId);
                    writer.WriteAttributeString("eventCount", Invariant(replayRecords.Count));
                    foreach (ReviewDecisionReplayRecord record in replayRecords)
                        WriteDecisionReplayRecord(writer, record);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                hashing.FlushFinalBlock();
                digest = BytesToHex(sha256.Hash);
            }
            return digest;
        }

        private static void WriteDecisionReplayRecord(
            XmlWriter writer,
            ReviewDecisionReplayRecord record)
        {
            DecisionEvent decisionEvent = record.DecisionEvent;
            writer.WriteStartElement("eptd", "replayRecord", DecisionCustomXmlNamespace);
            writer.WriteAttributeString("contractVersion", Invariant(record.ContractVersion));
            writer.WriteAttributeString("sequence", Invariant(record.Sequence));
            writer.WriteAttributeString("priorGenerationId", record.PriorGenerationId);
            writer.WriteAttributeString("newGenerationId", record.NewGenerationId);

            writer.WriteStartElement("eptd", "event", DecisionCustomXmlNamespace);
            writer.WriteAttributeString("eventId", decisionEvent.EventId);
            writer.WriteAttributeString("requestId", decisionEvent.RequestId);
            writer.WriteAttributeString("projectId", decisionEvent.ProjectId);
            writer.WriteAttributeString("evidenceGenerationId", decisionEvent.EvidenceGenerationId);
            writer.WriteAttributeString("mappingGenerationId", decisionEvent.MappingGenerationId);
            writer.WriteAttributeString("priorDecision", Invariant((int)decisionEvent.PriorDecision));
            writer.WriteAttributeString("newDecision", Invariant((int)decisionEvent.NewDecision));
            writer.WriteAttributeString(
                "createdUtc",
                decisionEvent.CreatedUtc.ToString("O", CultureInfo.InvariantCulture));
            writer.WriteAttributeString("selectionScope", Invariant((int)decisionEvent.SelectionScope));
            writer.WriteAttributeString("runId", decisionEvent.RunId);
            writer.WriteAttributeString("subjectId", decisionEvent.SubjectId);
            writer.WriteAttributeString("groupId", decisionEvent.GroupId);
            writer.WriteAttributeString(
                "selectedObservationId",
                decisionEvent.SelectedObservationId);
            writer.WriteAttributeString("supersededEventId", decisionEvent.SupersededEventId);
            writer.WriteAttributeString(
                "affectedSubjectCount",
                Invariant(decisionEvent.AffectedSubjectIds.Count));
            writer.WriteAttributeString(
                "supportingRevisionCount",
                Invariant(decisionEvent.SupportingRevisionIds.Count));
            writer.WriteAttributeString(
                "supportingContentHashCount",
                Invariant(decisionEvent.SupportingContentHashes.Count));
            writer.WriteAttributeString(
                "reviewerSourceCount",
                Invariant(decisionEvent.ReviewerSnapshot.IdentitySourceLabels.Count));
            writer.WriteAttributeString("locationCount", Invariant(decisionEvent.Locations.Count));

            writer.WriteStartElement("eptd", "affectedSubjects", DecisionCustomXmlNamespace);
            foreach (string subjectId in decisionEvent.AffectedSubjectIds)
            {
                writer.WriteStartElement("eptd", "subject", DecisionCustomXmlNamespace);
                writer.WriteAttributeString("subjectId", subjectId);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            writer.WriteStartElement("eptd", "supportingRevisions", DecisionCustomXmlNamespace);
            foreach (string revisionId in decisionEvent.SupportingRevisionIds)
            {
                writer.WriteStartElement("eptd", "revision", DecisionCustomXmlNamespace);
                writer.WriteAttributeString("revisionId", revisionId);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            writer.WriteStartElement("eptd", "supportingContentHashes", DecisionCustomXmlNamespace);
            foreach (string contentHash in decisionEvent.SupportingContentHashes)
            {
                writer.WriteStartElement("eptd", "content", DecisionCustomXmlNamespace);
                writer.WriteAttributeString("contentHash", contentHash);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            writer.WriteStartElement("eptd", "reviewer", DecisionCustomXmlNamespace);
            WriteEncodedTextElement(
                writer,
                "confirmedDisplayName",
                decisionEvent.ReviewerSnapshot.ConfirmedDisplayName);
            WriteEncodedTextElement(
                writer,
                "officeDisplayName",
                decisionEvent.ReviewerSnapshot.OfficeDisplayName);
            writer.WriteStartElement("eptd", "identitySources", DecisionCustomXmlNamespace);
            foreach (string source in decisionEvent.ReviewerSnapshot.IdentitySourceLabels)
                WriteEncodedTextElement(writer, "source", source);
            writer.WriteEndElement();
            writer.WriteEndElement();

            WriteEncodedTextElement(writer, "addInBuild", decisionEvent.AddInBuild);
            WriteEncodedTextElement(writer, "note", decisionEvent.Note);

            writer.WriteStartElement("eptd", "locations", DecisionCustomXmlNamespace);
            foreach (DecisionLocationContext location in decisionEvent.Locations)
            {
                writer.WriteStartElement("eptd", "location", DecisionCustomXmlNamespace);
                writer.WriteAttributeString("role", location.Role);
                writer.WriteAttributeString("revisionId", location.RevisionId);
                writer.WriteAttributeString("physicalRow", Invariant(location.PhysicalRow));
                writer.WriteAttributeString("physicalColumn", Invariant(location.PhysicalColumn));
                WriteEncodedTextElement(writer, "logicalSheetId", location.LogicalSheetId);
                WriteEncodedTextElement(writer, "sheetName", location.SheetName);
                WriteEncodedTextElement(writer, "address", location.Address);
                WriteEncodedTextElement(writer, "sourceLabel", location.SourceLabel);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteScalarElement(XmlWriter writer, string localName, CanonicalScalar scalar)
        {
            if (scalar == null) throw new InvalidDataException("A persisted canonical scalar is required.");
            byte[] payload = DecisionPayloadBytes(scalar.CanonicalText);
            writer.WriteStartElement("eptd", localName, DecisionCustomXmlNamespace);
            writer.WriteAttributeString("type", Invariant((int)scalar.Type));
            writer.WriteAttributeString("dateSystem", Invariant((int)scalar.DateSystem));
            writer.WriteAttributeString("encoding", "base64-utf8");
            writer.WriteAttributeString("byteLength", Invariant(payload.Length));
            writer.WriteBase64(payload, 0, payload.Length);
            writer.WriteEndElement();
        }

        private static void WriteEncodedTextElement(XmlWriter writer, string localName, string value)
        {
            byte[] payload = DecisionPayloadBytes(value ?? "");
            writer.WriteStartElement("eptd", localName, DecisionCustomXmlNamespace);
            writer.WriteAttributeString("encoding", "base64-utf8");
            writer.WriteAttributeString("byteLength", Invariant(payload.Length));
            writer.WriteBase64(payload, 0, payload.Length);
            writer.WriteEndElement();
        }

        private static byte[] DecisionPayloadBytes(string value)
        {
            byte[] payload;
            try
            {
                payload = new UTF8Encoding(false, true).GetBytes(value ?? "");
            }
            catch (EncoderFallbackException exception)
            {
                throw new InvalidDataException(
                    "A decision-state text payload contains invalid Unicode.",
                    exception);
            }
            if (payload.Length > MaxDecisionPayloadBytes)
                throw new InvalidDataException(
                    "A decision-state text payload exceeds the safe per-value limit.");
            return payload;
        }

        private static string Invariant(int value) =>
            value.ToString(CultureInfo.InvariantCulture);

        private static void WriteManifestPart(
            WorkbookPart workbookPart,
            string projectId,
            string itemId,
            IReadOnlyList<StateChunk> chunks,
            string totalSha256,
            string projectionSha256,
            string decisionItemId,
            string decisionSha256,
            string decisionEvidenceGenerationId,
            string decisionMappingGenerationId)
        {
            CustomXmlPart part = AddCustomXmlPart(workbookPart, itemId);
            long totalLength = chunks.Sum(chunk => (long)chunk.Length);
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (XmlWriter writer = CreateCustomXmlWriter(stream))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("ept", "reviewManifest", CustomXmlNamespace);
                WriteEnvelopeAttributes(
                    writer,
                    itemId,
                    string.IsNullOrEmpty(decisionItemId)
                        ? StreamingProjectionManifestSchemaVersion
                        : DecisionBoundStreamingProjectionManifestSchemaVersion);
                writer.WriteAttributeString("projectId", projectId);
                writer.WriteAttributeString("encoding", "utf-8");
                writer.WriteAttributeString("chunkCount", chunks.Count.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("totalLength", totalLength.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("sha256", totalSha256);
                writer.WriteAttributeString("projectionSha256", projectionSha256);
                writer.WriteAttributeString(
                    "projectionDigestVersion",
                    StreamingProjectionDigestVersion);
                if (!string.IsNullOrEmpty(decisionItemId))
                {
                    writer.WriteAttributeString("decisionItemId", decisionItemId);
                    writer.WriteAttributeString("decisionSha256", decisionSha256);
                    writer.WriteAttributeString(
                        "decisionEvidenceGenerationId",
                        decisionEvidenceGenerationId);
                    writer.WriteAttributeString(
                        "decisionMappingGenerationId",
                        decisionMappingGenerationId);
                }
                writer.WriteAttributeString("createdUtc", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));

                foreach (StateChunk chunk in chunks)
                {
                    writer.WriteStartElement("ept", "chunk", CustomXmlNamespace);
                    writer.WriteAttributeString("itemId", chunk.ItemId);
                    writer.WriteAttributeString("sequence", chunk.Sequence.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("length", chunk.Length.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("sha256", chunk.Sha256);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static CustomXmlPart AddCustomXmlPart(WorkbookPart workbookPart, string itemId)
        {
            return AddCustomXmlPart(workbookPart, itemId, CustomXmlNamespace);
        }

        private static CustomXmlPart AddCustomXmlPart(
            WorkbookPart workbookPart,
            string itemId,
            string schemaNamespace)
        {
            CustomXmlPart part = workbookPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            CustomXmlPropertiesPart properties = part.AddNewPart<CustomXmlPropertiesPart>();
            var dataStoreItem = new DataStoreItem { ItemId = itemId };
            var references = new SchemaReferences();
            references.Append(new SchemaReference { Uri = schemaNamespace });
            dataStoreItem.Append(references);
            properties.DataStoreItem = dataStoreItem;
            return part;
        }

        private static XmlWriter CreateCustomXmlWriter(Stream stream)
        {
            return XmlWriter.Create(stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                Indent = false,
                CloseOutput = false,
                CheckCharacters = true,
            });
        }

        private static void WriteEnvelopeAttributes(
            XmlWriter writer,
            string itemId,
            string schemaVersion = SchemaVersion)
        {
            writer.WriteAttributeString("marker", Marker);
            writer.WriteAttributeString("schemaVersion", schemaVersion);
            writer.WriteAttributeString("itemId", itemId);
        }

        private static CustomXmlRoot ReadCustomXmlRoot(CustomXmlPart part)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(
                       stream,
                       Math.Max(MaxOwnedCustomXmlCharacters, MaxDecisionXmlCharacters)))
            {
                if (reader.MoveToContent() != XmlNodeType.Element)
                    throw new InvalidDataException("A Custom XML part has no document element.");
                return new CustomXmlRoot(reader.LocalName, reader.NamespaceURI);
            }
        }

        private static XDocument LoadCustomXml(CustomXmlPart part, long maxCharacters)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(stream, maxCharacters))
                return XDocument.Load(reader, LoadOptions.None);
        }

        private static DecisionPartData ReadDecisionStatePart(
            CustomXmlPart part,
            Guid persistentItemId)
        {
            try
            {
                return ParseDecisionStatePart(part, persistentItemId);
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is XmlException || exception is FormatException)
            {
                throw new InvalidDataException(
                    "The comparison decision state is not valid bounded XML.",
                    exception);
            }
        }

        private static void ValidateDecisionBinding(
            ManifestPartData manifest,
            DecisionPartData decision)
        {
            if (!manifest.DecisionItemId.HasValue)
            {
                if (decision != null)
                    throw new InvalidDataException(
                        "The workbook contains an unbound comparison decision state.");
                return;
            }
            if (decision == null)
                throw new InvalidDataException(
                    "The comparison manifest requires a missing decision state.");
            if (decision.ItemId != manifest.DecisionItemId.Value ||
                !FixedTimeEquals(decision.Sha256, manifest.DecisionSha256) ||
                decision.ManifestItemId != manifest.ItemId ||
                !string.Equals(decision.ProjectId, manifest.ProjectId, StringComparison.Ordinal) ||
                !FixedTimeEquals(decision.StateSha256, manifest.StateSha256) ||
                !FixedTimeEquals(decision.ProjectionSha256, manifest.ProjectionSha256) ||
                !string.Equals(
                    decision.Generation.Catalog.EvidenceGenerationId,
                    manifest.DecisionEvidenceGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    decision.Generation.Catalog.MappingGenerationId,
                    manifest.DecisionMappingGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The comparison decision state does not match its review manifest bindings.");
        }

        private static XmlReader CreateBoundedXmlReader(Stream stream, long maxCharacters)
        {
            return XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maxCharacters,
                IgnoreComments = true,
            });
        }

        private static OpenSettings CreateBoundedOpenSettings()
        {
            return new OpenSettings
            {
                AutoSave = false,
                MaxCharactersInPart = MaxSpreadsheetPartCharacters,
            };
        }

        private static long MaxChunkXmlCharacters(int decodedLength)
        {
            return ((long)decodedLength + 2) / 3 * 4 + MaxChunkXmlOverheadCharacters;
        }

        private static Guid ReadAndValidatePersistentItemId(
            CustomXmlPart part,
            string expectedSchemaNamespace)
        {
            CustomXmlPropertiesPart properties = part.CustomXmlPropertiesPart;
            DataStoreItem item = properties?.DataStoreItem;
            string itemId = item?.ItemId?.Value;
            if (string.IsNullOrWhiteSpace(itemId) || !Guid.TryParse(itemId, out Guid parsed))
                throw new InvalidDataException("A comparison Custom XML part has no valid persistent item ID.");
            bool hasSchemaReference = item
                .Descendants<SchemaReference>()
                .Any(reference => string.Equals(
                    reference.Uri?.Value,
                    expectedSchemaNamespace,
                    StringComparison.Ordinal));
            if (!hasSchemaReference)
                throw new InvalidDataException(
                    "A comparison Custom XML part has no matching v1 schema reference.");
            return parsed;
        }

        private static void ValidateEnvelope(XElement root, Guid persistentItemId)
        {
            if (!string.Equals(RequireAttribute(root, "marker"), Marker, StringComparison.Ordinal))
                throw new InvalidDataException("The comparison Custom XML marker is invalid.");
            if (!string.Equals(RequireAttribute(root, "schemaVersion"), SchemaVersion, StringComparison.Ordinal))
                throw new InvalidDataException("The comparison Custom XML schema version is unsupported.");
            Guid contentItemId = ParseItemId(RequireAttribute(root, "itemId"), "itemId");
            if (contentItemId != persistentItemId)
                throw new InvalidDataException("The Custom XML content and properties item IDs differ.");
        }

        private static string ValidateManifestEnvelope(XElement root, Guid persistentItemId)
        {
            if (!string.Equals(RequireAttribute(root, "marker"), Marker, StringComparison.Ordinal))
                throw new InvalidDataException("The comparison Custom XML marker is invalid.");
            string schemaVersion = RequireAttribute(root, "schemaVersion");
            if (!string.Equals(schemaVersion, SchemaVersion, StringComparison.Ordinal) &&
                !string.Equals(
                    schemaVersion,
                    DecisionBoundManifestSchemaVersion,
                    StringComparison.Ordinal) &&
                !string.Equals(
                    schemaVersion,
                    StreamingProjectionManifestSchemaVersion,
                    StringComparison.Ordinal) &&
                !string.Equals(
                    schemaVersion,
                    DecisionBoundStreamingProjectionManifestSchemaVersion,
                    StringComparison.Ordinal))
                throw new InvalidDataException("The comparison manifest schema version is unsupported.");
            Guid contentItemId = ParseItemId(RequireAttribute(root, "itemId"), "itemId");
            if (contentItemId != persistentItemId)
                throw new InvalidDataException("The Custom XML content and properties item IDs differ.");
            return schemaVersion;
        }

        private static ManifestPartData ParseManifest(
            XElement root,
            Guid itemId,
            string manifestSchemaVersion)
        {
            ValidateAllowedAttributes(
                root,
                "marker", "schemaVersion", "itemId", "projectId", "encoding", "chunkCount",
                "totalLength", "sha256", "projectionSha256", "decisionItemId",
                "decisionSha256", "decisionEvidenceGenerationId",
                "decisionMappingGenerationId", "projectionDigestVersion", "createdUtc");
            EnsureOnlyExpectedChildren(root, StateNamespace + "chunk");

            if (!string.Equals(RequireAttribute(root, "encoding"), "utf-8", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The comparison state encoding is unsupported.");
            string projectId = RequireNonEmptyAttribute(root, "projectId");
            int chunkCount = ParseNonNegativeInt(root, "chunkCount");
            if (chunkCount == 0 || chunkCount > MaxStateChunkCount)
                throw new InvalidDataException("The comparison manifest chunk count exceeds the safe limit.");
            long totalLength = ParseNonNegativeLong(root, "totalLength");
            if (totalLength > MaxEmbeddedStateBytes)
                throw new InvalidDataException("The comparison manifest state length exceeds the safe limit.");
            string stateSha256 = ParseSha256(root, "sha256");
            string projectionSha256 = ParseSha256(root, "projectionSha256");
            bool streamingProjectionManifest = string.Equals(
                manifestSchemaVersion,
                StreamingProjectionManifestSchemaVersion,
                StringComparison.Ordinal) || string.Equals(
                manifestSchemaVersion,
                DecisionBoundStreamingProjectionManifestSchemaVersion,
                StringComparison.Ordinal);
            XAttribute projectionDigestVersionAttribute = root.Attribute("projectionDigestVersion");
            if (streamingProjectionManifest)
            {
                if (projectionDigestVersionAttribute == null ||
                    !string.Equals(
                        projectionDigestVersionAttribute.Value,
                        StreamingProjectionDigestVersion,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "A streaming comparison manifest requires its exact projection digest version.");
            }
            else if (projectionDigestVersionAttribute != null)
            {
                throw new InvalidDataException(
                    "A legacy comparison manifest cannot claim a streaming projection digest version.");
            }
            string projectionDigestVersion = streamingProjectionManifest
                ? StreamingProjectionDigestVersion
                : LegacyProjectionDigestVersion;
            XAttribute decisionItemAttribute = root.Attribute("decisionItemId");
            XAttribute decisionShaAttribute = root.Attribute("decisionSha256");
            XAttribute decisionEvidenceAttribute = root.Attribute("decisionEvidenceGenerationId");
            XAttribute decisionMappingAttribute = root.Attribute("decisionMappingGenerationId");
            int decisionBindingAttributeCount =
                (decisionItemAttribute == null ? 0 : 1) +
                (decisionShaAttribute == null ? 0 : 1) +
                (decisionEvidenceAttribute == null ? 0 : 1) +
                (decisionMappingAttribute == null ? 0 : 1);
            if (decisionBindingAttributeCount != 0 && decisionBindingAttributeCount != 4)
                throw new InvalidDataException(
                    "The comparison manifest decision binding is incomplete.");
            bool decisionBoundManifest = string.Equals(
                manifestSchemaVersion,
                DecisionBoundManifestSchemaVersion,
                StringComparison.Ordinal) || string.Equals(
                manifestSchemaVersion,
                DecisionBoundStreamingProjectionManifestSchemaVersion,
                StringComparison.Ordinal);
            if (decisionBoundManifest && decisionItemAttribute == null)
                throw new InvalidDataException(
                    "A decision-bound comparison manifest requires its decision item ID and SHA-256.");
            if (!decisionBoundManifest && decisionItemAttribute != null)
                throw new InvalidDataException(
                    "A legacy comparison manifest cannot claim a decision sidecar binding.");
            Guid? decisionItemId = decisionItemAttribute == null
                ? (Guid?)null
                : ParseItemId(decisionItemAttribute.Value, "decisionItemId");
            string decisionSha256 = decisionShaAttribute == null
                ? ""
                : ParseSha256(root, "decisionSha256");
            string decisionEvidenceGenerationId = decisionEvidenceAttribute == null
                ? ""
                : RequireNonEmptyAttribute(root, "decisionEvidenceGenerationId");
            string decisionMappingGenerationId = decisionMappingAttribute == null
                ? ""
                : RequireNonEmptyAttribute(root, "decisionMappingGenerationId");
            if (!DateTimeOffset.TryParse(
                    RequireAttribute(root, "createdUtc"),
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind,
                    out _))
                throw new InvalidDataException("The comparison manifest creation timestamp is invalid.");

            var bySequence = new Dictionary<int, ManifestChunkData>();
            var itemIds = new HashSet<Guid>();
            long declaredLength = 0;
            foreach (XElement element in root.Elements())
            {
                ValidateAllowedAttributes(element, "itemId", "sequence", "length", "sha256");
                EnsureEmptyElement(element);
                Guid chunkItemId = ParseItemId(RequireNonEmptyAttribute(element, "itemId"), "itemId");
                if (!itemIds.Add(chunkItemId))
                    throw new InvalidDataException("A manifest chunk item ID is duplicated.");
                int sequence = ParseNonNegativeInt(element, "sequence");
                if (bySequence.ContainsKey(sequence))
                    throw new InvalidDataException("A manifest state chunk sequence is duplicated.");
                int length = ParseNonNegativeInt(element, "length");
                if (length > MaxStateChunkSizeBytes)
                    throw new InvalidDataException("A manifest state chunk exceeds the safe chunk limit.");
                declaredLength = checked(declaredLength + length);
                bySequence.Add(sequence, new ManifestChunkData(
                    chunkItemId,
                    sequence,
                    length,
                    ParseSha256(element, "sha256")));
            }

            if (bySequence.Count != chunkCount || declaredLength != totalLength)
                throw new InvalidDataException("The comparison manifest counts or lengths are inconsistent.");
            var ordered = new List<ManifestChunkData>(chunkCount);
            for (int sequence = 0; sequence < chunkCount; sequence++)
            {
                if (!bySequence.TryGetValue(sequence, out ManifestChunkData chunk))
                    throw new InvalidDataException("State chunk sequences must be contiguous from zero.");
                ordered.Add(chunk);
            }
            return new ManifestPartData(
                itemId,
                projectId,
                totalLength,
                stateSha256,
                projectionSha256,
                projectionDigestVersion,
                decisionItemId,
                decisionSha256,
                decisionEvidenceGenerationId,
                decisionMappingGenerationId,
                ordered);
        }

        private static ChunkPartData ReadStateChunkMetadata(CustomXmlPart part, Guid persistentItemId)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(stream, MaxOwnedCustomXmlCharacters))
            {
                if (reader.MoveToContent() != XmlNodeType.Element ||
                    reader.LocalName != "stateChunk" || reader.NamespaceURI != CustomXmlNamespace)
                    throw new InvalidDataException("A comparison state chunk has an invalid root element.");

                ValidateAllowedAttributes(
                    reader,
                    "marker", "schemaVersion", "itemId", "projectId", "sequence", "length", "sha256");
                ValidateEnvelope(reader, persistentItemId);
                string projectId = RequireNonEmptyAttribute(reader, "projectId");
                int sequence = ParseNonNegativeInt(reader, "sequence");
                int length = ParseNonNegativeInt(reader, "length");
                if (length > MaxStateChunkSizeBytes)
                    throw new InvalidDataException("A comparison state chunk exceeds the safe chunk limit.");
                string sha256 = ParseSha256(reader, "sha256");
                return new ChunkPartData(
                    persistentItemId,
                    projectId,
                    sequence,
                    length,
                    sha256,
                    part);
            }
        }

        private static void DecodeStateChunk(
            ChunkPartData actual,
            ManifestChunkData expected,
            byte[] destination,
            int destinationOffset)
        {
            using (Stream stream = actual.Part.GetStream(FileMode.Open, FileAccess.Read))
            using (XmlReader reader = CreateBoundedXmlReader(
                       stream,
                       MaxChunkXmlCharacters(expected.Length)))
            using (SHA256 sha256 = SHA256.Create())
            {
                try
                {
                    if (reader.MoveToContent() != XmlNodeType.Element ||
                        reader.LocalName != "stateChunk" || reader.NamespaceURI != CustomXmlNamespace)
                        throw new InvalidDataException("A comparison state chunk has an invalid root element.");

                    var buffer = new byte[8192];
                    int written = 0;
                    int read;
                    while ((read = reader.ReadElementContentAsBase64(buffer, 0, buffer.Length)) > 0)
                    {
                        if (written > expected.Length - read)
                            throw new InvalidDataException("A comparison state chunk contains more bytes than declared.");
                        Buffer.BlockCopy(buffer, 0, destination, destinationOffset + written, read);
                        sha256.TransformBlock(buffer, 0, read, buffer, 0);
                        written += read;
                    }
                    sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);

                    if (written != expected.Length)
                        throw new InvalidDataException("A comparison state chunk length is invalid: " + actual.ItemId);
                    if (!FixedTimeEquals(BytesToHex(sha256.Hash), expected.Sha256))
                        throw new InvalidDataException("A comparison state chunk SHA-256 is invalid: " + actual.ItemId);
                }
                catch (Exception exception) when (
                    exception is XmlException || exception is FormatException || exception is InvalidOperationException)
                {
                    throw new InvalidDataException("A comparison state chunk is not valid Base64 text.", exception);
                }
            }
        }

        private static void ValidatePassiveReviewPackage(
            SpreadsheetDocument document,
            WorkbookPart workbookPart)
        {
            if (document.DocumentType != SpreadsheetDocumentType.Workbook)
                throw new InvalidDataException("A comparison review must be a macro-free .xlsx workbook.");
            if (workbookPart.VbaProjectPart != null)
                throw new InvalidDataException("A comparison review workbook must not contain VBA.");
            if (workbookPart.ExternalWorkbookParts.Any())
                throw new InvalidDataException("A comparison review workbook must not contain external workbooks.");
            if (workbookPart.Workbook.ExternalReferences != null)
                throw new InvalidDataException("A comparison review workbook must not contain external references.");
            if (workbookPart.Workbook.DefinedNames != null &&
                workbookPart.Workbook.DefinedNames.Elements<DefinedName>().Any())
                throw new InvalidDataException("A comparison review workbook must not contain defined-name formulas.");
            ValidateStreamingWorkbookPresentation(workbookPart.Workbook);

            var blockedPartTypeNames = new HashSet<string>(StringComparer.Ordinal)
            {
                "ActiveXControlPart",
                "ActiveXControlBinaryPart",
                "ConnectionsPart",
                "ControlPropertiesPart",
                "CustomDataPart",
                "DrawingsPart",
                "EmbeddedObjectPart",
                "EmbeddedPackagePart",
                "ExternalWorkbookPart",
                "PivotTablePart",
                "QuickAccessToolbarCustomizationsPart",
                "QueryTablePart",
                "RibbonAndBackstageCustomizationsPart",
                "RibbonExtensibilityPart",
                "TableDefinitionPart",
                "ThreadedCommentsPart",
                "VbaDataPart",
                "VbaProjectPart",
                "VmlDrawingPart",
                "WebExTaskpanesPart",
                "WorksheetCommentsPart",
            };
            var pending = new Queue<OpenXmlPart>();
            var visited = new HashSet<OpenXmlPart>();
            if (document.ExternalRelationships.Any())
                throw new InvalidDataException(
                    "A comparison review workbook contains a package-level external relationship.");
            foreach (IdPartPair rootPart in document.Parts)
                pending.Enqueue(rootPart.OpenXmlPart);
            while (pending.Count > 0)
            {
                OpenXmlPart part = pending.Dequeue();
                if (!visited.Add(part)) continue;
                if (blockedPartTypeNames.Contains(part.GetType().Name))
                    throw new InvalidDataException(
                        "A comparison review workbook contains blocked or unsupported content: " +
                        part.GetType().Name);
                if (part.ExternalRelationships.Any() || part.HyperlinkRelationships.Any())
                    throw new InvalidDataException(
                        "A comparison review workbook contains an external relationship: " + part.Uri);

                foreach (IdPartPair child in part.Parts)
                    pending.Enqueue(child.OpenXmlPart);
            }

            ValidatePassiveReviewPresentationParts(workbookPart);
        }

        private static bool HasUnsupportedPresentation(Worksheet worksheet)
        {
            SheetFormatProperties format = worksheet.SheetFormatProperties;
            if (format != null &&
                (format.ZeroHeight?.Value == true ||
                 (format.DefaultRowHeight != null &&
                  format.DefaultRowHeight.Value < MinVisibleRowHeightPoints) ||
                 (format.DefaultColumnWidth != null &&
                  format.DefaultColumnWidth.Value < MinVisibleColumnWidth)))
                return true;

            if (worksheet.Descendants<Row>().Any(row =>
                    row.Hidden?.Value == true ||
                    row.Collapsed?.Value == true ||
                    (row.StyleIndex != null && row.StyleIndex.Value != 0) ||
                    (row.Height != null && row.Height.Value < MinVisibleRowHeightPoints)) ||
                worksheet.Descendants<Column>().Any(column =>
                    column.Hidden?.Value == true ||
                    column.Collapsed?.Value == true ||
                    (column.Style != null && column.Style.Value != 0) ||
                    (column.Width != null && column.Width.Value < MinVisibleColumnWidth)) ||
                worksheet.Descendants<MergeCell>().Any() ||
                worksheet.Descendants<ConditionalFormatting>().Any() ||
                worksheet.Descendants<Cell>().Any(cell =>
                    cell.StyleIndex != null && cell.StyleIndex.Value != 0) ||
                worksheet.Descendants<RunProperties>().Any(HasUnsupportedRunProperties))
                return true;

            return false;
        }

        private static bool HasUnsupportedRunProperties(RunProperties properties)
        {
            if (properties.GetAttributes().Count != 0) return true;

            int boldCount = 0;
            int strikeCount = 0;
            int colorCount = 0;
            int fontNameCount = 0;
            int fontSizeCount = 0;
            foreach (OpenXmlElement child in properties.ChildElements)
            {
                if (child is Bold)
                {
                    if (++boldCount > 1 || !HasOnlyUnqualifiedAttributes(child, "val")) return true;
                }
                else if (child is Strike)
                {
                    if (++strikeCount > 1 || !HasOnlyUnqualifiedAttributes(child, "val")) return true;
                }
                else if (child is Color)
                {
                    var color = (Color)child;
                    string rgb = color.Rgb?.Value ?? "";
                    if (++colorCount > 1 ||
                        !HasOnlyUnqualifiedAttributes(child, "rgb") ||
                        !AllowedReviewRunRgb.Contains(rgb)) return true;
                }
                else if (child is RunFont)
                {
                    var font = (RunFont)child;
                    if (++fontNameCount > 1 ||
                        !HasOnlyUnqualifiedAttributes(child, "val") ||
                        !AllowedVisibleFontNames.Contains(font.Val?.Value ?? "")) return true;
                }
                else if (child is FontSize)
                {
                    var size = (FontSize)child;
                    if (++fontSizeCount > 1 ||
                        !HasOnlyUnqualifiedAttributes(child, "val") ||
                        size.Val == null || size.Val.Value < MinVisibleFontSizePoints) return true;
                }
                else
                {
                    // Italic, underline, outline, shadow, vertical alignment and similar properties can hide
                    // evidence or change the meaning of the protected color vocabulary.
                    return true;
                }
            }

            return false;
        }

        private static void ValidatePassiveReviewPresentationParts(WorkbookPart workbookPart)
        {
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart;
            ThemePart themePart = workbookPart.ThemePart;
            if (stylesPart == null)
            {
                if (themePart != null)
                    throw new InvalidDataException(
                        "A comparison review workbook contains an unreferenced theme without a validated style table.");
                return;
            }

            Stylesheet stylesheet = stylesPart.Stylesheet;
            CellFormats formats = stylesheet?.CellFormats;
            CellFormat defaultFormat = formats?.Elements<CellFormat>().FirstOrDefault();
            if (defaultFormat == null ||
                ValueOrZero(defaultFormat.FontId) != 0U ||
                ValueOrZero(defaultFormat.FillId) != 0U ||
                ValueOrZero(defaultFormat.BorderId) != 0U ||
                ValueOrZero(defaultFormat.NumberFormatId) != 0U ||
                ValueOrZero(defaultFormat.FormatId) != 0U ||
                IsTrue(defaultFormat.ApplyNumberFormat) ||
                IsTrue(defaultFormat.ApplyFont) ||
                IsTrue(defaultFormat.ApplyFill) ||
                IsTrue(defaultFormat.ApplyBorder) ||
                IsTrue(defaultFormat.ApplyAlignment) ||
                IsTrue(defaultFormat.ApplyProtection) ||
                IsTrue(defaultFormat.PivotButton) ||
                IsTrue(defaultFormat.QuotePrefix) ||
                defaultFormat.Alignment != null ||
                defaultFormat.Protection != null)
            {
                throw new InvalidDataException(
                    "A comparison review workbook has an unsafe active default cell style.");
            }

            Font defaultFont = stylesheet.Fonts?.Elements<Font>().FirstOrDefault();
            Fill defaultFill = stylesheet.Fills?.Elements<Fill>().FirstOrDefault();
            Border defaultBorder = stylesheet.Borders?.Elements<Border>().FirstOrDefault();
            if (!IsSafeDefaultFont(defaultFont, themePart) ||
                !IsSafeDefaultFill(defaultFill) ||
                !IsSafeDefaultBorder(defaultBorder))
            {
                throw new InvalidDataException(
                    "A comparison review workbook has an unsafe active default font, fill, border, or theme.");
            }
        }

        private static bool IsSafeDefaultFont(Font font, ThemePart themePart)
        {
            if (font == null || font.GetAttributes().Count != 0) return false;
            FontSize size = font.GetFirstChild<FontSize>();
            FontName name = font.GetFirstChild<FontName>();
            Color color = font.GetFirstChild<Color>();
            if (size?.Val == null || size.Val.Value < MinVisibleFontSizePoints ||
                name?.Val == null || !AllowedVisibleFontNames.Contains(name.Val.Value) ||
                color == null)
                return false;

            foreach (OpenXmlElement child in font.ChildElements)
            {
                if (!(child is FontSize) && !(child is Color) && !(child is FontName) &&
                    !(child is FontFamilyNumbering) && !(child is FontScheme) &&
                    !(child is FontCharSet))
                    return false;
            }

            if (color.Rgb != null)
                return IsDarkVisibleRgb(color.Rgb.Value);
            if (color.Theme == null || color.Theme.Value != 1U || color.Tint != null ||
                color.Indexed != null || color.Auto?.Value == true)
                return false;
            return IsSafeDarkOneTheme(themePart);
        }

        private static bool IsSafeDefaultFill(Fill fill)
        {
            PatternFill pattern = fill?.PatternFill;
            return pattern != null &&
                   (pattern.PatternType == null || pattern.PatternType.Value == PatternValues.None) &&
                   pattern.ForegroundColor == null && pattern.BackgroundColor == null;
        }

        private static bool IsSafeDefaultBorder(Border border)
        {
            if (border == null || border.GetAttributes().Count != 0) return false;
            foreach (OpenXmlElement child in border.ChildElements)
            {
                var property = child as BorderPropertiesType;
                if (property == null || property.Style != null || property.Color != null)
                    return false;
            }
            return true;
        }

        private static bool IsSafeDarkOneTheme(ThemePart themePart)
        {
            if (themePart?.Theme == null) return false;
            XDocument theme;
            try
            {
                theme = XDocument.Parse(themePart.Theme.OuterXml, LoadOptions.None);
            }
            catch (Exception exception) when (exception is XmlException || exception is InvalidOperationException)
            {
                return false;
            }

            XElement darkOne = theme.Descendants().FirstOrDefault(element => element.Name.LocalName == "dk1");
            XElement color = darkOne?.Elements().SingleOrDefault();
            string rgb = color?.Attribute("lastClr")?.Value ?? color?.Attribute("val")?.Value;
            return IsDarkVisibleRgb(rgb);
        }

        private static bool IsDarkVisibleRgb(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            if (value.Length == 8) value = value.Substring(2);
            if (value.Length != 6 || !int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb))
                return false;
            int red = (rgb >> 16) & 0xFF;
            int green = (rgb >> 8) & 0xFF;
            int blue = rgb & 0xFF;
            double relativeLuminance = (0.2126d * red + 0.7152d * green + 0.0722d * blue) / 255d;
            return relativeLuminance <= 0.55d;
        }

        private static uint ValueOrZero(UInt32Value value) => value?.Value ?? 0U;

        private static bool IsTrue(BooleanValue value) => value?.Value == true;

        private static bool HasOnlyUnqualifiedAttributes(OpenXmlElement element, params string[] allowedNames)
        {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            foreach (OpenXmlAttribute attribute in element.GetAttributes())
            {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri) || !allowed.Contains(attribute.LocalName))
                    return false;
            }
            return true;
        }

        private static string ComputeLegacyProjectionSha256(WorkbookPart workbookPart)
        {
            Sheets sheets = workbookPart.Workbook?.Sheets;
            if (sheets == null)
                throw new InvalidDataException("The review workbook has no sheets collection.");

            List<Sheet> actualSheets = sheets.Elements<Sheet>().ToList();
            if (actualSheets.Count != ExpectedSheetNames.Count)
                throw new InvalidDataException("The review workbook sheet count is invalid.");

            List<SharedStringItem> sharedStrings = workbookPart.SharedStringTablePart?
                .SharedStringTable?
                .Elements<SharedStringItem>()
                .ToList() ?? new List<SharedStringItem>();

            using (SHA256 sha256 = SHA256.Create())
            {
                HashProjectionValue(sha256, LegacyProjectionDigestVersion);
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
                        throw new InvalidDataException("A review workbook sheet relationship is invalid.", exception);
                    }
                    if (worksheetPart?.Worksheet == null)
                        throw new InvalidDataException("A review workbook sheet is not a worksheet.");

                    HashProjectionValue(sha256, name);
                    List<Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                        .OrderBy(cell => cell.CellReference?.Value ?? "", StringComparer.Ordinal)
                        .ToList();
                    var references = new HashSet<string>(StringComparer.Ordinal);
                    HashProjectionValue(sha256, cells.Count.ToString(CultureInfo.InvariantCulture));
                    foreach (Cell cell in cells)
                    {
                        string reference = cell.CellReference?.Value;
                        if (string.IsNullOrEmpty(reference) || !references.Add(reference))
                            throw new InvalidDataException("A review workbook cell reference is missing or duplicated.");
                        if (cell.CellFormula != null)
                            throw new InvalidDataException(
                                "A comparison review workbook must not contain executable worksheet formulas.");
                        HashProjectionValue(sha256, reference);
                        HashProjectionValue(sha256, ReadProjectionCellValue(cell, sharedStrings));
                    }

                    List<Hyperlink> hyperlinks = worksheetPart.Worksheet.Descendants<Hyperlink>()
                        .OrderBy(link => link.Reference?.Value ?? "", StringComparer.Ordinal)
                        .ThenBy(link => link.Location?.Value ?? "", StringComparer.Ordinal)
                        .ToList();
                    HashProjectionValue(sha256, hyperlinks.Count.ToString(CultureInfo.InvariantCulture));
                    foreach (Hyperlink hyperlink in hyperlinks)
                    {
                        if (hyperlink.Id != null && !string.IsNullOrEmpty(hyperlink.Id.Value))
                            throw new InvalidDataException("Review workbook hyperlinks must be internal.");
                        string reference = hyperlink.Reference?.Value ?? "";
                        string display = hyperlink.Display?.Value ?? "";
                        Cell linkedCell = cells.SingleOrDefault(cell =>
                            string.Equals(cell.CellReference?.Value, reference, StringComparison.Ordinal));
                        string visibleCellText = ReadProjectionVisibleText(linkedCell, sharedStrings);
                        if (visibleCellText != null &&
                            string.Equals(display, visibleCellText, StringComparison.Ordinal))
                        {
                            // Desktop Excel materializes the cell text as a redundant display attribute
                            // when it saves an internal hyperlink. Its presence does not change what the
                            // user sees, while a different explicit display value remains integrity-bound.
                            display = "";
                        }
                        HashProjectionValue(sha256, reference);
                        HashProjectionValue(sha256, hyperlink.Location?.Value ?? "");
                        HashProjectionValue(sha256, display);
                        HashProjectionValue(sha256, hyperlink.Tooltip?.Value ?? "");
                    }
                }

                sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                return BytesToHex(sha256.Hash);
            }
        }

        private static string ReadProjectionCellValue(Cell cell, IReadOnlyList<SharedStringItem> sharedStrings)
        {
            string type = cell.GetAttribute("t", "").Value;
            string raw = cell.CellValue?.Text ?? "";
            switch (type)
            {
                case "s":
                    if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int index) ||
                        index < 0 || index >= sharedStrings.Count)
                        throw new InvalidDataException("A review workbook shared-string index is invalid.");
                    return "text:" + ReadProjectionText(sharedStrings[index]);
                case "inlineStr":
                    return "text:" + ReadProjectionText(cell.InlineString);
                case "str":
                    return "text:" + raw;
                default:
                    return "scalar:" + type + ":" + raw;
            }
        }

        private static string ReadProjectionVisibleText(
            Cell cell,
            IReadOnlyList<SharedStringItem> sharedStrings)
        {
            if (cell == null) return null;
            string type = cell.GetAttribute("t", "").Value;
            string raw = cell.CellValue?.Text ?? "";
            switch (type)
            {
                case "s":
                    if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int index) ||
                        index < 0 || index >= sharedStrings.Count)
                        throw new InvalidDataException("A review workbook shared-string index is invalid.");
                    return sharedStrings[index].InnerText ?? "";
                case "inlineStr":
                    return cell.InlineString?.InnerText ?? "";
                case "str":
                    return raw;
                default:
                    return null;
            }
        }

        private static string ReadProjectionText(OpenXmlElement container)
        {
            if (container == null) return "";
            List<Run> runs = container.Elements<Run>().ToList();
            if (runs.Count == 0) return container.InnerText ?? "";

            var styles = new List<string>();
            var texts = new List<StringBuilder>();
            foreach (OpenXmlElement child in container.ChildElements)
            {
                Run run = child as Run;
                string style;
                string text;
                if (run == null)
                {
                    style = child is DocumentFormat.OpenXml.Spreadsheet.Text
                        ? ""
                        : "element:" + child.LocalName;
                    text = child.InnerText ?? "";
                }
                else
                {
                    RunProperties properties = run.RunProperties;
                    Color color = properties?.GetFirstChild<Color>();
                    string bold = IsEnabled(properties?.GetFirstChild<Bold>()) ? "bold" : "";
                    string strike = IsEnabled(properties?.GetFirstChild<Strike>()) ? "strike" : "";
                    string rgb = (color?.Rgb?.Value ?? "").ToUpperInvariant();
                    string theme = color?.Theme == null
                        ? ""
                        : color.Theme.Value.ToString(CultureInfo.InvariantCulture);
                    string indexed = color?.Indexed == null
                        ? ""
                        : color.Indexed.Value.ToString(CultureInfo.InvariantCulture);
                    if (bold.Length == 0 && strike.Length == 0 && rgb.Length == 0 &&
                        theme.Length == 0 && indexed.Length == 0)
                    {
                        style = "";
                    }
                    else
                    {
                        var signature = new StringBuilder();
                        AppendProjectionToken(signature, bold);
                        AppendProjectionToken(signature, strike);
                        AppendProjectionToken(signature, rgb);
                        AppendProjectionToken(signature, theme);
                        AppendProjectionToken(signature, indexed);
                        style = signature.ToString();
                    }
                    text = run.InnerText ?? "";
                }

                if (styles.Count > 0 && string.Equals(styles[styles.Count - 1], style, StringComparison.Ordinal))
                    texts[texts.Count - 1].Append(text);
                else
                {
                    styles.Add(style);
                    texts.Add(new StringBuilder(text));
                }
            }

            if (styles.All(string.IsNullOrEmpty))
                return string.Concat(texts.Select(text => text.ToString()));

            var result = new StringBuilder("rich-v1");
            for (int index = 0; index < styles.Count; index++)
            {
                AppendProjectionToken(result, styles[index]);
                AppendProjectionToken(result, texts[index].ToString());
            }
            return result.ToString();
        }

        private static bool IsEnabled(OpenXmlElement element)
        {
            if (element == null) return false;
            OpenXmlAttribute attribute = element.GetAttributes()
                .FirstOrDefault(candidate =>
                    candidate.LocalName == "val" && string.IsNullOrEmpty(candidate.NamespaceUri));
            string value = attribute.Value ?? "";
            return string.IsNullOrEmpty(value) || value == "1" ||
                   string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static void AppendProjectionToken(StringBuilder target, string value)
        {
            value = value ?? "";
            target.Append(value.Length.ToString(CultureInfo.InvariantCulture));
            target.Append(':');
            target.Append(value);
        }

        private static void HashProjectionValue(HashAlgorithm hash, string value)
        {
            byte[] bytes = new UTF8Encoding(false, true).GetBytes(value ?? "");
            byte[] length =
            {
                (byte)(bytes.Length >> 24),
                (byte)(bytes.Length >> 16),
                (byte)(bytes.Length >> 8),
                (byte)bytes.Length,
            };
            hash.TransformBlock(length, 0, length.Length, length, 0);
            if (bytes.Length > 0)
                hash.TransformBlock(bytes, 0, bytes.Length, bytes, 0);
        }

        private static void VerifyClosedPackage(
            string temporaryPath,
            ReviewWorkbookState expected,
            int expectedChunkCount,
            string expectedSha256)
        {
            // ReadAndValidate is deliberately the only full-package verification pass. A whole-document
            // OpenXmlValidator walk loads every worksheet DOM and defeats the streaming scale boundary.
            // The writer emits fixed workbook grammar; the reader validates its relationships, worksheet
            // grammar, presentation policy, formulas and semantic projection from bounded raw streams.
            ReviewWorkbookReadResult read = ReadAndValidate(temporaryPath);
            if (!string.Equals(read.ProjectId, expected.ProjectId, StringComparison.Ordinal) ||
                !string.Equals(read.StateJson, expected.StateJson, StringComparison.Ordinal) ||
                read.ChunkCount != expectedChunkCount ||
                !FixedTimeEquals(read.StateSha256, expectedSha256) ||
                read.HasDecisionState != (expected.DecisionGeneration != null) ||
                (expected.DecisionGeneration != null &&
                 (!string.Equals(
                      read.DecisionGeneration.CatalogId,
                      expected.DecisionGeneration.CatalogId,
                      StringComparison.Ordinal) ||
                  !string.Equals(
                      read.DecisionGeneration.GenerationId,
                      expected.DecisionGeneration.GenerationId,
                      StringComparison.Ordinal))))
                throw new InvalidDataException("The closed review workbook did not reproduce its source state.");
        }

        private static List<StateChunk> SplitIntoChunks(byte[] payload, int chunkSize)
        {
            int count = Math.Max(1, (int)(((long)payload.Length + chunkSize - 1) / chunkSize));
            if (count > MaxStateChunkCount)
                throw new InvalidDataException("The comparison state requires too many review-workbook chunks.");
            var chunks = new List<StateChunk>(count);
            for (int sequence = 0; sequence < count; sequence++)
            {
                int offset = sequence * chunkSize;
                int length = Math.Min(chunkSize, payload.Length - offset);
                chunks.Add(new StateChunk(
                    sequence,
                    NewItemId(),
                    payload,
                    offset,
                    length,
                    ComputeSha256(payload, offset, length)));
            }
            return chunks;
        }

        private static string ComputeSha256(byte[] bytes)
        {
            byte[] digest;
            using (SHA256 sha256 = SHA256.Create()) digest = sha256.ComputeHash(bytes);
            return BytesToHex(digest);
        }

        private static string ComputeSha256(byte[] bytes, int offset, int length)
        {
            byte[] digest;
            using (SHA256 sha256 = SHA256.Create()) digest = sha256.ComputeHash(bytes, offset, length);
            return BytesToHex(digest);
        }

        private static string BytesToHex(byte[] digest)
        {
            var text = new StringBuilder(digest.Length * 2);
            foreach (byte value in digest) text.Append(value.ToString("x2", CultureInfo.InvariantCulture));
            return text.ToString();
        }

        private static bool FixedTimeEquals(string left, string right)
        {
            if (left == null || right == null || left.Length != right.Length) return false;
            int difference = 0;
            for (int index = 0; index < left.Length; index++)
                difference |= char.ToUpperInvariant(left[index]) ^ char.ToUpperInvariant(right[index]);
            return difference == 0;
        }

        private static string RequireAttribute(XElement element, string name)
        {
            XAttribute attribute = element.Attribute(name);
            if (attribute == null)
                throw new InvalidDataException("Required comparison state attribute is missing: " + name);
            return attribute.Value;
        }

        private static string RequireAttribute(XmlReader reader, string name)
        {
            string value = reader.GetAttribute(name);
            if (value == null)
                throw new InvalidDataException("Required comparison state attribute is missing: " + name);
            return value;
        }

        private static string RequireNonEmptyAttribute(XElement element, string name)
        {
            string value = RequireAttribute(element, name);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException("Required comparison state attribute is empty: " + name);
            return value;
        }

        private static string RequireNonEmptyAttribute(XmlReader reader, string name)
        {
            string value = RequireAttribute(reader, name);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException("Required comparison state attribute is empty: " + name);
            return value;
        }

        private static int ParseNonNegativeInt(XElement element, string name)
        {
            string value = RequireAttribute(element, name);
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) || parsed < 0)
                throw new InvalidDataException("Invalid non-negative integer attribute: " + name);
            return parsed;
        }

        private static int ParseNonNegativeInt(XmlReader reader, string name)
        {
            string value = RequireAttribute(reader, name);
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) || parsed < 0)
                throw new InvalidDataException("Invalid non-negative integer attribute: " + name);
            return parsed;
        }

        private static long ParseNonNegativeLong(XElement element, string name)
        {
            string value = RequireAttribute(element, name);
            if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) || parsed < 0)
                throw new InvalidDataException("Invalid non-negative integer attribute: " + name);
            return parsed;
        }

        private static string ParseSha256(XElement element, string name)
        {
            string value = RequireAttribute(element, name);
            if (value.Length != 64 || value.Any(ch => !IsHex(ch)))
                throw new InvalidDataException("Invalid SHA-256 attribute: " + name);
            return value;
        }

        private static string ParseSha256(XmlReader reader, string name)
        {
            string value = RequireAttribute(reader, name);
            if (value.Length != 64 || value.Any(ch => !IsHex(ch)))
                throw new InvalidDataException("Invalid SHA-256 attribute: " + name);
            return value;
        }

        private static Guid ParseItemId(string value, string name)
        {
            if (!Guid.TryParse(value, out Guid parsed))
                throw new InvalidDataException("Invalid comparison state GUID attribute: " + name);
            return parsed;
        }

        private static string FormatItemId(Guid itemId) => itemId.ToString("B").ToUpperInvariant();

        private static void ValidateEnvelope(XmlReader reader, Guid persistentItemId)
        {
            if (!string.Equals(RequireAttribute(reader, "marker"), Marker, StringComparison.Ordinal))
                throw new InvalidDataException("The comparison Custom XML marker is invalid.");
            if (!string.Equals(RequireAttribute(reader, "schemaVersion"), SchemaVersion, StringComparison.Ordinal))
                throw new InvalidDataException("The comparison Custom XML schema version is unsupported.");
            if (ParseItemId(RequireAttribute(reader, "itemId"), "itemId") != persistentItemId)
                throw new InvalidDataException("The Custom XML content and properties item IDs differ.");
        }

        private static void ValidateAllowedAttributes(XElement element, params string[] allowedNames)
        {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            foreach (XAttribute attribute in element.Attributes())
            {
                if (attribute.IsNamespaceDeclaration) continue;
                if (attribute.Name.NamespaceName.Length != 0 || !allowed.Contains(attribute.Name.LocalName))
                    throw new InvalidDataException(
                        "Unknown comparison state attribute: " + attribute.Name);
            }
        }

        private static void ValidateAllowedAttributes(XmlReader reader, params string[] allowedNames)
        {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            if (!reader.HasAttributes) return;
            reader.MoveToFirstAttribute();
            do
            {
                bool namespaceDeclaration = reader.Prefix == "xmlns" || reader.Name == "xmlns";
                if (!namespaceDeclaration &&
                    (reader.NamespaceURI.Length != 0 || !allowed.Contains(reader.LocalName)))
                    throw new InvalidDataException("Unknown comparison state attribute: " + reader.Name);
            } while (reader.MoveToNextAttribute());
            reader.MoveToElement();
        }

        private static void EnsureOnlyExpectedChildren(XElement element, XName expectedName)
        {
            foreach (XNode node in element.Nodes())
            {
                XElement child = node as XElement;
                if (child != null)
                {
                    if (child.Name != expectedName)
                        throw new InvalidDataException(
                            "Unknown comparison state element: " + child.Name);
                    continue;
                }
                XText text = node as XText;
                if (text == null || !string.IsNullOrWhiteSpace(text.Value))
                    throw new InvalidDataException("Unexpected content in the comparison manifest.");
            }
        }

        private static void EnsureEmptyElement(XElement element)
        {
            if (element.HasElements || element.Nodes().Any(node =>
                    !(node is XText text) || !string.IsNullOrWhiteSpace(text.Value)))
                throw new InvalidDataException("A comparison manifest chunk entry must be empty.");
        }

        private static bool IsHex(char value)
        {
            return (value >= '0' && value <= '9') ||
                   (value >= 'a' && value <= 'f') ||
                   (value >= 'A' && value <= 'F');
        }

        private static string ColumnName(int column)
        {
            var result = new StringBuilder();
            while (column > 0)
            {
                column--;
                result.Insert(0, (char)('A' + column % 26));
                column /= 26;
            }
            return result.ToString();
        }

        private static void EnsureSupportedPublicationPath(string destination)
        {
            Uri uri;
            if (Uri.TryCreate(destination, UriKind.Absolute, out uri) && uri.IsUnc)
                throw new NotSupportedException(
                    "Atomic review-workbook publication has not been established for network paths. " +
                    "Choose a local fixed drive, then copy the verified result if required.");

            string root = Path.GetPathRoot(destination);
            if (string.IsNullOrEmpty(root))
                throw new NotSupportedException("The review-workbook destination has no supported local volume.");
            try
            {
                DriveType driveType = new DriveInfo(root).DriveType;
                if (driveType != DriveType.Fixed)
                    throw new NotSupportedException(
                        "Atomic review-workbook publication is currently limited to local fixed drives.");
            }
            catch (ArgumentException exception)
            {
                throw new NotSupportedException(
                    "The review-workbook destination volume could not be validated.", exception);
            }
        }

        private static string SanitizeXmlText(string value)
        {
            if (string.IsNullOrEmpty(value)) return value ?? "";
            var result = new StringBuilder(value.Length);
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                if (char.IsHighSurrogate(current) && index + 1 < value.Length &&
                    char.IsLowSurrogate(value[index + 1]))
                {
                    result.Append(current).Append(value[++index]);
                }
                else if (!char.IsSurrogate(current) && XmlConvert.IsXmlChar(current))
                {
                    result.Append(current);
                }
                else
                {
                    // XML 1.0 cannot represent every .NET character. Keep the alteration visible in the
                    // projection; the exact original remains integrity-protected in the embedded state.
                    result.Append('\uFFFD');
                }
            }
            return result.ToString();
        }

        private static string ProjectCellText(string value)
        {
            string sanitized = SanitizeXmlText(value ?? "");
            if (sanitized.Length <= MaxExcelCellTextCharacters) return sanitized;

            int prefixLength = MaxExcelCellTextCharacters - DetailTruncationMarker.Length;
            if (prefixLength > 0 && prefixLength < sanitized.Length &&
                char.IsHighSurrogate(sanitized[prefixLength - 1]) &&
                char.IsLowSurrogate(sanitized[prefixLength]))
                prefixLength--;
            return sanitized.Substring(0, prefixLength) + DetailTruncationMarker;
        }

        private static string NewItemId()
        {
            return Guid.NewGuid().ToString("B").ToUpperInvariant();
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch
            {
                // Preserve the original generation/publication failure. The GUID temporary name cannot be
                // mistaken for a published review workbook and can be cleaned on the next maintenance pass.
            }
        }

        private sealed class StateChunk
        {
            public StateChunk(
                int sequence,
                string itemId,
                byte[] payload,
                int offset,
                int length,
                string sha256)
            {
                Sequence = sequence;
                ItemId = itemId;
                Payload = payload;
                Offset = offset;
                Length = length;
                Sha256 = sha256;
            }

            public int Sequence { get; }
            public string ItemId { get; }
            public byte[] Payload { get; }
            public int Offset { get; }
            public int Length { get; }
            public string Sha256 { get; }
        }

        private sealed class ZipEnvelope
        {
            public ZipEnvelope(int entryCount, long centralDirectorySize)
            {
                EntryCount = entryCount;
                CentralDirectorySize = centralDirectorySize;
            }

            public int EntryCount { get; }
            public long CentralDirectorySize { get; }
        }

        private sealed class ManifestChunkData
        {
            public ManifestChunkData(Guid itemId, int sequence, int length, string sha256)
            {
                ItemId = itemId;
                Sequence = sequence;
                Length = length;
                Sha256 = sha256;
            }

            public Guid ItemId { get; }
            public int Sequence { get; }
            public int Length { get; }
            public string Sha256 { get; }
        }

        private sealed class ManifestPartData
        {
            public ManifestPartData(
                Guid itemId,
                string projectId,
                long totalLength,
                string stateSha256,
                string projectionSha256,
                string projectionDigestVersion,
                Guid? decisionItemId,
                string decisionSha256,
                string decisionEvidenceGenerationId,
                string decisionMappingGenerationId,
                IReadOnlyList<ManifestChunkData> chunks)
            {
                ItemId = itemId;
                ProjectId = projectId;
                TotalLength = totalLength;
                StateSha256 = stateSha256;
                ProjectionSha256 = projectionSha256;
                ProjectionDigestVersion = projectionDigestVersion;
                DecisionItemId = decisionItemId;
                DecisionSha256 = decisionSha256 ?? "";
                DecisionEvidenceGenerationId = decisionEvidenceGenerationId ?? "";
                DecisionMappingGenerationId = decisionMappingGenerationId ?? "";
                Chunks = chunks;
            }

            public Guid ItemId { get; }
            public string ProjectId { get; }
            public long TotalLength { get; }
            public string StateSha256 { get; }
            public string ProjectionSha256 { get; }
            public string ProjectionDigestVersion { get; }
            public Guid? DecisionItemId { get; }
            public string DecisionSha256 { get; }
            public string DecisionEvidenceGenerationId { get; }
            public string DecisionMappingGenerationId { get; }
            public IReadOnlyList<ManifestChunkData> Chunks { get; }
        }

        private sealed class ChunkPartData
        {
            public ChunkPartData(
                Guid itemId,
                string projectId,
                int sequence,
                int length,
                string sha256,
                CustomXmlPart part)
            {
                ItemId = itemId;
                ProjectId = projectId;
                Sequence = sequence;
                Length = length;
                Sha256 = sha256;
                Part = part;
            }

            public Guid ItemId { get; }
            public string ProjectId { get; }
            public int Sequence { get; }
            public int Length { get; }
            public string Sha256 { get; }
            public CustomXmlPart Part { get; }
        }

        private sealed class DecisionPartData
        {
            public DecisionPartData(
                Guid itemId,
                Guid manifestItemId,
                string projectId,
                string stateSha256,
                string projectionSha256,
                string sha256,
                ReviewDecisionGeneration generation)
            {
                ItemId = itemId;
                ManifestItemId = manifestItemId;
                ProjectId = projectId;
                StateSha256 = stateSha256;
                ProjectionSha256 = projectionSha256;
                Sha256 = sha256;
                Generation = generation;
            }

            public Guid ItemId { get; }
            public Guid ManifestItemId { get; }
            public string ProjectId { get; }
            public string StateSha256 { get; }
            public string ProjectionSha256 { get; }
            public string Sha256 { get; }
            public ReviewDecisionGeneration Generation { get; }
        }

        private sealed class CustomXmlRoot
        {
            public CustomXmlRoot(string localName, string namespaceName)
            {
                LocalName = localName ?? "";
                NamespaceName = namespaceName ?? "";
            }

            public string LocalName { get; }
            public string NamespaceName { get; }
        }
    }
}
