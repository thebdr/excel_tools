using System;

namespace ExcelPowerTools.Core.Comparison.Publication
{
    public enum ValidatedStagePublicationCapabilityCode
    {
        NotEvaluated = 0,
        Supported,
        UnsupportedOperatingSystem,
        InvalidPath,
        NetworkPath,
        DifferentDirectory,
        DirectoryUnavailable,
        StageUnavailable,
        ReparsePoint,
        NonFixedVolume,
        UnsupportedFileSystem,
        IdentityUnavailable,
        StageAliased,
        CaseSensitiveDirectory,
        StageAlternateDataStreams,
        PathAlias,
        StageTooLarge,
        PathTooDeep,
    }

    public enum ValidatedStagePublicationStatus
    {
        Published = 0,
        CapabilityRejected,
        InvalidExpectedIdentity,
        StageValidationFailed,
        DestinationAlreadyExists,
        DestinationRaceLost,
        PublicationFailed,
        PostPublicationVerificationFailed,
    }

    /// <summary>Immutable explanation of whether the bounded NTFS publication primitive can be used.</summary>
    public sealed class ValidatedStagePublicationCapability
    {
        internal ValidatedStagePublicationCapability(
            ValidatedStagePublicationCapabilityCode code,
            string message,
            string correctiveAction,
            int nativeErrorCode = 0)
        {
            Code = code;
            Message = message ?? "";
            CorrectiveAction = correctiveAction ?? "";
            NativeErrorCode = nativeErrorCode;
        }

        public ValidatedStagePublicationCapabilityCode Code { get; }
        public bool IsSupported => Code == ValidatedStagePublicationCapabilityCode.Supported;
        public string Message { get; }
        public string CorrectiveAction { get; }
        public int NativeErrorCode { get; }
    }

    /// <summary>Identity captured before and after publishing from the same retained stage handle.</summary>
    public sealed class ValidatedStagePublicationMetadata
    {
        internal ValidatedStagePublicationMetadata(
            string stagePath,
            string destinationPath,
            long byteLength,
            string sha256,
            string stageOccurrenceToken,
            string publishedOccurrenceToken,
            string directoryOccurrenceToken,
            string destinationVolumeGuidPath,
            bool singleUnnamedDataStreamObserved,
            DateTimeOffset publishedUtc)
        {
            StagePath = stagePath ?? "";
            DestinationPath = destinationPath ?? "";
            ByteLength = byteLength;
            Sha256 = sha256 ?? "";
            StageOccurrenceToken = stageOccurrenceToken ?? "";
            PublishedOccurrenceToken = publishedOccurrenceToken ?? "";
            DirectoryOccurrenceToken = directoryOccurrenceToken ?? "";
            DestinationVolumeGuidPath = destinationVolumeGuidPath ?? "";
            SingleUnnamedDataStreamObserved = singleUnnamedDataStreamObserved;
            PublishedUtc = publishedUtc;
        }

        public string StagePath { get; }
        public string DestinationPath { get; }
        public long ByteLength { get; }
        public string Sha256 { get; }
        public string StageOccurrenceToken { get; }
        public string PublishedOccurrenceToken { get; }
        public string DirectoryOccurrenceToken { get; }
        public string DestinationVolumeGuidPath { get; }
        public bool SingleUnnamedDataStreamObserved { get; }
        public DateTimeOffset PublishedUtc { get; }
    }

    /// <summary>Structured outcome. Routine capability and race failures do not throw.</summary>
    public sealed class ValidatedStagePublicationResult
    {
        internal ValidatedStagePublicationResult(
            ValidatedStagePublicationStatus status,
            ValidatedStagePublicationCapability capability,
            ValidatedStagePublicationMetadata metadata,
            string message,
            string correctiveAction,
            int nativeErrorCode,
            bool destinationMayContainPublishedStage)
        {
            Status = status;
            Capability = capability;
            Metadata = metadata;
            Message = message ?? "";
            CorrectiveAction = correctiveAction ?? "";
            NativeErrorCode = nativeErrorCode;
            DestinationMayContainPublishedStage = destinationMayContainPublishedStage;
        }

        public ValidatedStagePublicationStatus Status { get; }
        public bool IsPublished => Status == ValidatedStagePublicationStatus.Published;
        public ValidatedStagePublicationCapability Capability { get; }
        public ValidatedStagePublicationMetadata Metadata { get; }
        public string Message { get; }
        public string CorrectiveAction { get; }
        public int NativeErrorCode { get; }

        /// <summary>
        /// True only when the path-chain-locked handle rename succeeded but a mandatory post-rename proof failed.
        /// The primitive never deletes or replaces that destination while reporting the failure.
        /// </summary>
        public bool DestinationMayContainPublishedStage { get; }
    }
}
