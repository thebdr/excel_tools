using System;
using System.IO;

namespace ExcelPowerTools.Core.Comparison.Inventory
{
    /// <summary>The workbook and separated-value formats understood by the comparison pipeline.</summary>
    public enum ArtifactFormatKind
    {
        Unknown = 0,
        Xlsx = 1,
        Xlsm = 2,
        Xlsb = 3,
        Csv = 4,
        Tsv = 5,
        DelimitedText = 6
    }

    /// <summary>How content evidence relates to the filename extension supplied by the user.</summary>
    public enum ArtifactFormatMatchKind
    {
        Match = 0,
        Compatible = 1,
        DeclaredHintOnly = 2,
        Mismatch = 3,
        UnsupportedExtension = 4,
        UnrecognizedContent = 5
    }

    /// <summary>The strongest bounded content evidence found while inventorying an artifact.</summary>
    public enum ArtifactFormatEvidenceKind
    {
        None = 0,
        OpenXmlWorkbookXml = 1,
        OpenXmlMacroWorkbook = 2,
        OpenXmlBinaryWorkbook = 3,
        CommaDelimiter = 4,
        TabDelimiter = 5,
        OtherDelimiter = 6,
        TextWithoutStableDelimiter = 7,
        EmptyFile = 8,
        InvalidZipPackage = 9,
        ZipWithoutWorkbook = 10,
        CompoundOfficeFile = 11,
        BinaryContent = 12,
        InspectionLimitExceeded = 13
    }

    /// <summary>
    /// Deterministic assessment of the declared extension and bounded signature/content inspection.
    /// No parser is selected silently when reliable evidence contradicts the extension.
    /// </summary>
    public sealed class ArtifactFormatAssessment
    {
        internal ArtifactFormatAssessment(
            string declaredExtension,
            ArtifactFormatKind declaredFormat,
            ArtifactFormatKind detectedFormat,
            ArtifactFormatKind effectiveFormat,
            ArtifactFormatMatchKind matchKind,
            ArtifactFormatEvidenceKind evidenceKind,
            bool isSupportedForComparison,
            string message)
        {
            DeclaredExtension = declaredExtension ?? "";
            DeclaredFormat = declaredFormat;
            DetectedFormat = detectedFormat;
            EffectiveFormat = effectiveFormat;
            MatchKind = matchKind;
            EvidenceKind = evidenceKind;
            IsSupportedForComparison = isSupportedForComparison;
            Message = message ?? "";
        }

        public string DeclaredExtension { get; }
        public ArtifactFormatKind DeclaredFormat { get; }
        public ArtifactFormatKind DetectedFormat { get; }
        public ArtifactFormatKind EffectiveFormat { get; }
        public ArtifactFormatMatchKind MatchKind { get; }
        public ArtifactFormatEvidenceKind EvidenceKind { get; }
        public bool IsSupportedForComparison { get; }
        public string Message { get; }

        public bool HasMismatch => MatchKind == ArtifactFormatMatchKind.Mismatch;
    }

    /// <summary>
    /// Immutable inventory of one file occurrence. InventoryFingerprint deterministically identifies the
    /// observed path/metadata/content tuple; it is not the project-scoped RevisionId assigned on import.
    /// ContentSha256 identifies exact bytes and can be shared by multiple revision occurrences. Creation and
    /// last-write timestamps are explicitly weak source metadata, not ordering proof.
    /// </summary>
    public sealed class ArtifactInventoryRecord
    {
        internal ArtifactInventoryRecord(
            string inventoryFingerprint,
            string contentSha256,
            string fullNormalizedPath,
            long byteLength,
            DateTimeOffset weakCreationTimeUtc,
            DateTimeOffset weakLastWriteTimeUtc,
            ArtifactFormatAssessment format)
        {
            InventoryFingerprint = inventoryFingerprint ??
                                   throw new ArgumentNullException(nameof(inventoryFingerprint));
            ContentSha256 = contentSha256 ?? throw new ArgumentNullException(nameof(contentSha256));
            FullNormalizedPath = fullNormalizedPath ?? throw new ArgumentNullException(nameof(fullNormalizedPath));
            ByteLength = byteLength;
            WeakCreationTimeUtc = weakCreationTimeUtc;
            WeakLastWriteTimeUtc = weakLastWriteTimeUtc;
            Format = format ?? throw new ArgumentNullException(nameof(format));
        }

        public string InventoryFingerprint { get; }

        /// <summary>
        /// Compatibility alias. This value fingerprints an inventory observation; callers must create a
        /// separate project-scoped RevisionId even when the same file is imported more than once.
        /// </summary>
        public string OccurrenceId => InventoryFingerprint;
        public string ContentSha256 { get; }
        public string FullNormalizedPath { get; }
        public long ByteLength { get; }
        public DateTimeOffset WeakCreationTimeUtc { get; }
        public DateTimeOffset WeakLastWriteTimeUtc { get; }
        public ArtifactFormatAssessment Format { get; }
    }

    public enum ArtifactInventoryStage
    {
        Hashing = 0,
        DetectingFormat = 1,
        VerifyingStability = 2,
        Completed = 3
    }

    /// <summary>Bounded progress payload suitable for large-file UI and cancellation.</summary>
    public sealed class ArtifactInventoryProgress
    {
        internal ArtifactInventoryProgress(
            string fullNormalizedPath,
            ArtifactInventoryStage stage,
            long bytesRead,
            long totalBytes)
        {
            FullNormalizedPath = fullNormalizedPath ?? "";
            Stage = stage;
            BytesRead = bytesRead;
            TotalBytes = totalBytes;
        }

        public string FullNormalizedPath { get; }
        public ArtifactInventoryStage Stage { get; }
        public long BytesRead { get; }
        public long TotalBytes { get; }

        public double Fraction => TotalBytes == 0
            ? (Stage == ArtifactInventoryStage.Completed ? 1d : 0d)
            : Math.Min(1d, (double)BytesRead / TotalBytes);
    }

    /// <summary>Streaming limits and progress cadence. All buffers are bounded independently of file size.</summary>
    public sealed class ArtifactInventoryOptions
    {
        public static readonly ArtifactInventoryOptions Default = new ArtifactInventoryOptions();

        public ArtifactInventoryOptions(
            int bufferSize = 128 * 1024,
            long progressIntervalBytes = 1024 * 1024,
            int contentSampleBytes = 64 * 1024)
        {
            if (bufferSize < 1024 || bufferSize > 4 * 1024 * 1024)
                throw new ArgumentOutOfRangeException(nameof(bufferSize), "Buffer size must be between 1 KiB and 4 MiB.");
            if (progressIntervalBytes < 1)
                throw new ArgumentOutOfRangeException(nameof(progressIntervalBytes), "Progress interval must be positive.");
            if (contentSampleBytes < 512 || contentSampleBytes > 1024 * 1024)
                throw new ArgumentOutOfRangeException(nameof(contentSampleBytes), "Content sample must be between 512 bytes and 1 MiB.");

            BufferSize = bufferSize;
            ProgressIntervalBytes = progressIntervalBytes;
            ContentSampleBytes = contentSampleBytes;
        }

        public int BufferSize { get; }
        public long ProgressIntervalBytes { get; }
        public int ContentSampleBytes { get; }
    }

    public enum ArtifactInventoryErrorCode
    {
        InvalidPath = 0,
        MissingArtifact = 1,
        PathIsNotAFile = 2,
        UnreadableArtifact = 3,
        ArtifactChangedDuringRead = 4
    }

    /// <summary>Actionable inventory failure with a stable machine-readable category.</summary>
    public sealed class ArtifactInventoryException : IOException
    {
        internal ArtifactInventoryException(
            ArtifactInventoryErrorCode errorCode,
            string artifactPath,
            string message,
            Exception innerException = null)
            : base(message, innerException)
        {
            ErrorCode = errorCode;
            ArtifactPath = artifactPath ?? "";
        }

        public ArtifactInventoryErrorCode ErrorCode { get; }
        public string ArtifactPath { get; }
    }
}
