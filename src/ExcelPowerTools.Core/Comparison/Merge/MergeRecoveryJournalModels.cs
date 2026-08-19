using System;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    /// <summary>Versioned safety limits for the portable merge-recovery ledger.</summary>
    public static class MergeRecoveryJournalContract
    {
        public const int Version = 1;
        public const int MaxRecords = 4;
        public const int MaxRecordBytes = 64 * 1024;
        public const int MaxJournalBytes = (MaxRecords * MaxRecordBytes) + 4096;
        public const int MaxIdentityCharacters = 256;
        public const int MaxLeafCharacters = 255;
        public const string XmlNamespace =
            "urn:excel-power-tools:comparison:merge-recovery-journal:v1";
    }

    /// <summary>
    /// Durable recovery states. Failure and recovery are outcomes/evidence, not extra ledger states.
    /// </summary>
    public enum MergeRecoveryJournalState
    {
        Prepared = 1,
        Validated = 2,
        OutputPublished = 3,
        AuditCommitted = 4
    }

    /// <summary>How the exact expected output occurrence was observed after validation.</summary>
    public enum MergePublicationEvidenceKind
    {
        DirectPublisher = 1,
        RestartReconciled = 2
    }

    /// <summary>
    /// Immutable, portable binding frozen before any output is written. It deliberately contains a randomized
    /// destination locator and a privacy-safe leaf only; no local path, volume identity, machine name, or file
    /// occurrence token belongs in this model.
    /// </summary>
    public sealed class MergeRecoveryPreparedBinding
    {
        public MergeRecoveryPreparedBinding(
            string projectId,
            string mergeId,
            string sourceMasterRevisionId,
            string sourceProjectStateRoot,
            long baseByteLength,
            string baseSha256,
            string evidenceGenerationId,
            string mappingGenerationId,
            string decisionGenerationId,
            string decisionSetHash,
            string mergePlanHash,
            string destinationLocatorId,
            string destinationLeaf)
        {
            ProjectId = MergeRecoveryJournalValidation.RequireIdentity(projectId, nameof(projectId));
            MergeId = MergeRecoveryJournalValidation.RequireIdentity(mergeId, nameof(mergeId));
            SourceMasterRevisionId = MergeRecoveryJournalValidation.RequireIdentity(
                sourceMasterRevisionId,
                nameof(sourceMasterRevisionId));
            SourceProjectStateRoot = MergeRecoveryJournalValidation.RequireSha256(
                sourceProjectStateRoot,
                nameof(sourceProjectStateRoot));
            BaseByteLength = MergeRecoveryJournalValidation.RequireLength(
                baseByteLength,
                nameof(baseByteLength));
            BaseSha256 = MergeRecoveryJournalValidation.RequireSha256(baseSha256, nameof(baseSha256));
            EvidenceGenerationId = MergeRecoveryJournalValidation.RequireIdentity(
                evidenceGenerationId,
                nameof(evidenceGenerationId));
            MappingGenerationId = MergeRecoveryJournalValidation.RequireIdentity(
                mappingGenerationId,
                nameof(mappingGenerationId));
            DecisionGenerationId = MergeRecoveryJournalValidation.RequireIdentity(
                decisionGenerationId,
                nameof(decisionGenerationId));
            DecisionSetHash = MergeRecoveryJournalValidation.RequireSha256(
                decisionSetHash,
                nameof(decisionSetHash));
            MergePlanHash = MergeRecoveryJournalValidation.RequireSha256(
                mergePlanHash,
                nameof(mergePlanHash));
            DestinationLocatorId = MergeRecoveryJournalValidation.RequireIdentity(
                destinationLocatorId,
                nameof(destinationLocatorId));
            DestinationLeaf = MergeRecoveryJournalValidation.RequireLeaf(
                destinationLeaf,
                nameof(destinationLeaf));
        }

        public string ProjectId { get; }
        public string MergeId { get; }
        public string SourceMasterRevisionId { get; }
        public string SourceProjectStateRoot { get; }
        public long BaseByteLength { get; }
        public string BaseSha256 { get; }
        public string EvidenceGenerationId { get; }
        public string MappingGenerationId { get; }
        public string DecisionGenerationId { get; }
        public string DecisionSetHash { get; }
        public string MergePlanHash { get; }
        public string DestinationLocatorId { get; }
        public string DestinationLeaf { get; }

        internal bool HasSameFrozenFields(MergeRecoveryPreparedBinding other) =>
            other != null &&
            string.Equals(ProjectId, other.ProjectId, StringComparison.Ordinal) &&
            string.Equals(MergeId, other.MergeId, StringComparison.Ordinal) &&
            string.Equals(SourceMasterRevisionId, other.SourceMasterRevisionId, StringComparison.Ordinal) &&
            string.Equals(SourceProjectStateRoot, other.SourceProjectStateRoot, StringComparison.Ordinal) &&
            BaseByteLength == other.BaseByteLength &&
            string.Equals(BaseSha256, other.BaseSha256, StringComparison.Ordinal) &&
            string.Equals(EvidenceGenerationId, other.EvidenceGenerationId, StringComparison.Ordinal) &&
            string.Equals(MappingGenerationId, other.MappingGenerationId, StringComparison.Ordinal) &&
            string.Equals(DecisionGenerationId, other.DecisionGenerationId, StringComparison.Ordinal) &&
            string.Equals(DecisionSetHash, other.DecisionSetHash, StringComparison.Ordinal) &&
            string.Equals(MergePlanHash, other.MergePlanHash, StringComparison.Ordinal) &&
            string.Equals(DestinationLocatorId, other.DestinationLocatorId, StringComparison.Ordinal) &&
            string.Equals(DestinationLeaf, other.DestinationLeaf, StringComparison.Ordinal);
    }

    /// <summary>
    /// Exact validation evidence added by the Validated transition. StageLeaf must not equal the frozen
    /// destination leaf under OrdinalIgnoreCase: the supported Windows NTFS publication boundary deliberately
    /// rejects case-sensitive directories, so differently cased leaves would name the same target occurrence.
    /// </summary>
    public sealed class MergeRecoveryValidationEvidence
    {
        public MergeRecoveryValidationEvidence(
            string stageLeaf,
            string validationContractVersion,
            string validationReceiptHash,
            long expectedOutputLength,
            string expectedOutputSha256)
        {
            StageLeaf = MergeRecoveryJournalValidation.RequireLeaf(stageLeaf, nameof(stageLeaf));
            ValidationContractVersion = MergeRecoveryJournalValidation.RequireIdentity(
                validationContractVersion,
                nameof(validationContractVersion));
            ValidationReceiptHash = MergeRecoveryJournalValidation.RequireSha256(
                validationReceiptHash,
                nameof(validationReceiptHash));
            ExpectedOutputLength = MergeRecoveryJournalValidation.RequireLength(
                expectedOutputLength,
                nameof(expectedOutputLength));
            ExpectedOutputSha256 = MergeRecoveryJournalValidation.RequireSha256(
                expectedOutputSha256,
                nameof(expectedOutputSha256));
        }

        public string StageLeaf { get; }
        public string ValidationContractVersion { get; }
        public string ValidationReceiptHash { get; }
        public long ExpectedOutputLength { get; }
        public string ExpectedOutputSha256 { get; }

        internal bool HasSameFrozenFields(MergeRecoveryValidationEvidence other) =>
            other != null &&
            string.Equals(StageLeaf, other.StageLeaf, StringComparison.Ordinal) &&
            string.Equals(ValidationContractVersion, other.ValidationContractVersion, StringComparison.Ordinal) &&
            string.Equals(ValidationReceiptHash, other.ValidationReceiptHash, StringComparison.Ordinal) &&
            ExpectedOutputLength == other.ExpectedOutputLength &&
            string.Equals(ExpectedOutputSha256, other.ExpectedOutputSha256, StringComparison.Ordinal);
    }

    /// <summary>
    /// Sanitized exact observation returned by the direct publisher or collected during restart reconciliation.
    /// The caller retains any local path and filesystem identity outside this portable object.
    /// </summary>
    public sealed class MergeRecoveryOutputObservation
    {
        public MergeRecoveryOutputObservation(
            string destinationLocatorId,
            string destinationLeaf,
            long byteLength,
            string sha256,
            DateTimeOffset observedUtc)
        {
            DestinationLocatorId = MergeRecoveryJournalValidation.RequireIdentity(
                destinationLocatorId,
                nameof(destinationLocatorId));
            DestinationLeaf = MergeRecoveryJournalValidation.RequireLeaf(
                destinationLeaf,
                nameof(destinationLeaf));
            ByteLength = MergeRecoveryJournalValidation.RequireLength(byteLength, nameof(byteLength));
            Sha256 = MergeRecoveryJournalValidation.RequireSha256(sha256, nameof(sha256));
            ObservedUtc = MergeRecoveryJournalValidation.NormalizeTimestamp(
                observedUtc,
                nameof(observedUtc));
        }

        public string DestinationLocatorId { get; }
        public string DestinationLeaf { get; }
        public long ByteLength { get; }
        public string Sha256 { get; }
        public DateTimeOffset ObservedUtc { get; }
    }

    /// <summary>Portable project-state receipt required to finish the audit commit.</summary>
    public sealed class MergeRecoveryAuditReceipt
    {
        public MergeRecoveryAuditReceipt(
            string auditCommitId,
            string outputRevisionId,
            string committedProjectStateRoot,
            string committedReviewPackageSha256,
            DateTimeOffset committedUtc)
        {
            AuditCommitId = MergeRecoveryJournalValidation.RequireSha256(
                auditCommitId,
                nameof(auditCommitId));
            OutputRevisionId = MergeRecoveryJournalValidation.RequireSha256(
                outputRevisionId,
                nameof(outputRevisionId));
            CommittedProjectStateRoot = MergeRecoveryJournalValidation.RequireSha256(
                committedProjectStateRoot,
                nameof(committedProjectStateRoot));
            CommittedReviewPackageSha256 = MergeRecoveryJournalValidation.RequireSha256(
                committedReviewPackageSha256,
                nameof(committedReviewPackageSha256));
            CommittedUtc = MergeRecoveryJournalValidation.NormalizeTimestamp(
                committedUtc,
                nameof(committedUtc));
        }

        public string AuditCommitId { get; }
        public string OutputRevisionId { get; }
        public string CommittedProjectStateRoot { get; }
        public string CommittedReviewPackageSha256 { get; }
        public DateTimeOffset CommittedUtc { get; }

        internal bool HasSameFrozenFields(MergeRecoveryAuditReceipt other) =>
            other != null &&
            string.Equals(AuditCommitId, other.AuditCommitId, StringComparison.Ordinal) &&
            string.Equals(OutputRevisionId, other.OutputRevisionId, StringComparison.Ordinal) &&
            string.Equals(CommittedProjectStateRoot, other.CommittedProjectStateRoot, StringComparison.Ordinal) &&
            string.Equals(
                CommittedReviewPackageSha256,
                other.CommittedReviewPackageSha256,
                StringComparison.Ordinal) &&
            CommittedUtc.EqualsExact(other.CommittedUtc);
    }

    /// <summary>One immutable cumulative checkpoint in the recovery hash chain.</summary>
    public sealed class MergeRecoveryJournalRecord
    {
        internal MergeRecoveryJournalRecord(
            int contractVersion,
            int sequence,
            MergeRecoveryJournalState state,
            string recordId,
            string priorRecordId,
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation,
            string outputRevisionId,
            string auditCommitId,
            MergePublicationEvidenceKind? publicationEvidenceKind,
            DateTimeOffset? observedUtc,
            MergeRecoveryAuditReceipt auditReceipt)
        {
            if (contractVersion <= 0)
                throw new ArgumentOutOfRangeException(nameof(contractVersion));
            if (sequence < 0 || sequence >= MergeRecoveryJournalContract.MaxRecords)
                throw new ArgumentOutOfRangeException(nameof(sequence));
            if (!Enum.IsDefined(typeof(MergeRecoveryJournalState), state))
                throw new ArgumentOutOfRangeException(nameof(state));

            ContractVersion = contractVersion;
            Sequence = sequence;
            State = state;
            RecordId = MergeRecoveryJournalValidation.RequireSha256(recordId, nameof(recordId));
            PriorRecordId = string.IsNullOrEmpty(priorRecordId)
                ? ""
                : MergeRecoveryJournalValidation.RequireSha256(priorRecordId, nameof(priorRecordId));
            Binding = binding ?? throw new ArgumentNullException(nameof(binding));
            Validation = validation;
            OutputRevisionId = string.IsNullOrEmpty(outputRevisionId)
                ? ""
                : MergeRecoveryJournalValidation.RequireSha256(outputRevisionId, nameof(outputRevisionId));
            AuditCommitId = string.IsNullOrEmpty(auditCommitId)
                ? ""
                : MergeRecoveryJournalValidation.RequireSha256(auditCommitId, nameof(auditCommitId));
            if (publicationEvidenceKind.HasValue &&
                !Enum.IsDefined(typeof(MergePublicationEvidenceKind), publicationEvidenceKind.Value))
                throw new ArgumentOutOfRangeException(nameof(publicationEvidenceKind));
            PublicationEvidenceKind = publicationEvidenceKind;
            ObservedUtc = observedUtc.HasValue
                ? MergeRecoveryJournalValidation.NormalizeTimestamp(observedUtc.Value, nameof(observedUtc))
                : (DateTimeOffset?)null;
            AuditReceipt = auditReceipt;
            ValidateStateShape();
        }

        public int ContractVersion { get; }
        public int Sequence { get; }
        public MergeRecoveryJournalState State { get; }
        public string RecordId { get; }
        public string PriorRecordId { get; }
        public MergeRecoveryPreparedBinding Binding { get; }
        public MergeRecoveryValidationEvidence Validation { get; }
        public string OutputRevisionId { get; }
        public string AuditCommitId { get; }
        public MergePublicationEvidenceKind? PublicationEvidenceKind { get; }
        public DateTimeOffset? ObservedUtc { get; }
        public MergeRecoveryAuditReceipt AuditReceipt { get; }

        // Flattened read-only accessors keep the durable schema easy to inspect without duplicating mutable data.
        public string ProjectId => Binding.ProjectId;
        public string MergeId => Binding.MergeId;
        public string SourceMasterRevisionId => Binding.SourceMasterRevisionId;
        public string SourceProjectStateRoot => Binding.SourceProjectStateRoot;
        public long BaseByteLength => Binding.BaseByteLength;
        public string BaseSha256 => Binding.BaseSha256;
        public string EvidenceGenerationId => Binding.EvidenceGenerationId;
        public string MappingGenerationId => Binding.MappingGenerationId;
        public string DecisionGenerationId => Binding.DecisionGenerationId;
        public string DecisionSetHash => Binding.DecisionSetHash;
        public string MergePlanHash => Binding.MergePlanHash;
        public string DestinationLocatorId => Binding.DestinationLocatorId;
        public string DestinationLeaf => Binding.DestinationLeaf;

        private void ValidateStateShape()
        {
            bool hasValidation = Validation != null;
            bool hasDerivedIds = OutputRevisionId.Length != 0 || AuditCommitId.Length != 0;
            bool hasPublication = PublicationEvidenceKind.HasValue || ObservedUtc.HasValue;
            bool hasAudit = AuditReceipt != null;

            switch (State)
            {
                case MergeRecoveryJournalState.Prepared:
                    if (hasValidation || hasDerivedIds || hasPublication || hasAudit)
                        throw new ArgumentException(
                            "A Prepared checkpoint cannot contain fields belonging to a later state.");
                    break;
                case MergeRecoveryJournalState.Validated:
                    RequireValidationAndDerivedIds(hasValidation);
                    if (hasPublication || hasAudit)
                        throw new ArgumentException(
                            "A Validated checkpoint cannot contain publication or audit fields.");
                    break;
                case MergeRecoveryJournalState.OutputPublished:
                    RequireValidationAndDerivedIds(hasValidation);
                    if (!PublicationEvidenceKind.HasValue || !ObservedUtc.HasValue || hasAudit)
                        throw new ArgumentException(
                            "An OutputPublished checkpoint requires complete publication evidence and no audit receipt.");
                    break;
                case MergeRecoveryJournalState.AuditCommitted:
                    RequireValidationAndDerivedIds(hasValidation);
                    if (!PublicationEvidenceKind.HasValue || !ObservedUtc.HasValue || !hasAudit)
                        throw new ArgumentException(
                            "An AuditCommitted checkpoint requires complete publication and audit evidence.");
                    if (!string.Equals(AuditCommitId, AuditReceipt.AuditCommitId, StringComparison.Ordinal) ||
                        !string.Equals(OutputRevisionId, AuditReceipt.OutputRevisionId, StringComparison.Ordinal))
                        throw new ArgumentException(
                            "The audit receipt must repeat the deterministic validated audit and output identities.");
                    if (AuditReceipt.CommittedUtc < ObservedUtc.Value)
                        throw new ArgumentException(
                            "The audit commit timestamp cannot precede the output observation timestamp.");
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(State));
            }
        }

        private void RequireValidationAndDerivedIds(bool hasValidation)
        {
            if (!hasValidation || OutputRevisionId.Length == 0 || AuditCommitId.Length == 0)
                throw new ArgumentException(
                    "Validated and later checkpoints require validation evidence and deterministic derived identities.");
            if (string.Equals(
                    Binding.DestinationLeaf,
                    Validation.StageLeaf,
                    StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    "A validated stage leaf must not alias the destination leaf under Windows case-insensitive semantics.");
        }
    }

    /// <summary>Validated immutable view returned by replay.</summary>
    public sealed class MergeRecoveryJournalReplay
    {
        internal MergeRecoveryJournalReplay(
            System.Collections.Generic.IReadOnlyList<MergeRecoveryJournalRecord> records)
        {
            Records = records ?? throw new ArgumentNullException(nameof(records));
            Current = records.Count == 0 ? null : records[records.Count - 1];
        }

        public System.Collections.Generic.IReadOnlyList<MergeRecoveryJournalRecord> Records { get; }
        public MergeRecoveryJournalRecord Current { get; }
        public MergeRecoveryJournalState? CurrentState => Current?.State;
        public bool IsAuditCommitted => CurrentState == MergeRecoveryJournalState.AuditCommitted;
    }

    /// <summary>Stable diagnostic identifiers for recovery-ledger rejection.</summary>
    public static class MergeRecoveryJournalCodes
    {
        public const string Empty = "CMP-MERGE-JOURNAL-EMPTY";
        public const string Count = "CMP-MERGE-JOURNAL-COUNT";
        public const string Version = "CMP-MERGE-JOURNAL-VERSION";
        public const string Sequence = "CMP-MERGE-JOURNAL-SEQUENCE";
        public const string State = "CMP-MERGE-JOURNAL-STATE";
        public const string RecordId = "CMP-MERGE-JOURNAL-RECORD-ID";
        public const string PriorChain = "CMP-MERGE-JOURNAL-PRIOR-CHAIN";
        public const string FrozenFieldDrift = "CMP-MERGE-JOURNAL-FROZEN-DRIFT";
        public const string DerivedId = "CMP-MERGE-JOURNAL-DERIVED-ID";
        public const string Transition = "CMP-MERGE-JOURNAL-TRANSITION";
        public const string Observation = "CMP-MERGE-JOURNAL-OBSERVATION";
        public const string AuditReceipt = "CMP-MERGE-JOURNAL-AUDIT-RECEIPT";
        public const string LeafAlias = "CMP-MERGE-JOURNAL-LEAF-ALIAS";
        public const string Codec = "CMP-MERGE-JOURNAL-CODEC";
    }

    public sealed class MergeRecoveryJournalException : IOException
    {
        internal MergeRecoveryJournalException(string code, int sequence, string message)
            : base(message)
        {
            Code = code ?? "";
            Sequence = sequence;
        }

        public string Code { get; }
        public int Sequence { get; }
    }

    internal static class MergeRecoveryJournalValidation
    {
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

        internal static string RequireIdentity(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank portable identity is required.", parameterName);
            if (value.Length > MergeRecoveryJournalContract.MaxIdentityCharacters)
                throw new ArgumentException("A portable identity exceeds its character limit.", parameterName);
            EnsureXml10Text(value, parameterName);
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                bool allowed = (c >= 'A' && c <= 'Z') ||
                               (c >= 'a' && c <= 'z') ||
                               (c >= '0' && c <= '9') ||
                               c == '.' || c == '_' || c == '-';
                if (!allowed)
                    throw new ArgumentException(
                        "A portable identity may contain only ASCII letters, digits, '.', '_' and '-'.",
                        parameterName);
            }
            return value;
        }

        internal static string RequireSha256(string value, string parameterName)
        {
            if (!StableId.IsSha256(value))
                throw new ArgumentException("An exact SHA-256 value is required.", parameterName);
            return value.ToLowerInvariant();
        }

        internal static long RequireLength(long value, string parameterName)
        {
            if (value < 0)
                throw new ArgumentOutOfRangeException(parameterName, "A byte length cannot be negative.");
            return value;
        }

        internal static DateTimeOffset NormalizeTimestamp(DateTimeOffset value, string parameterName)
        {
            if (value == default(DateTimeOffset))
                throw new ArgumentException("A checkpoint timestamp is required.", parameterName);
            return value.ToUniversalTime();
        }

        internal static string RequireLeaf(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank privacy-safe filename leaf is required.", parameterName);
            if (value.Length > MergeRecoveryJournalContract.MaxLeafCharacters)
                throw new ArgumentException("A filename leaf exceeds its character limit.", parameterName);
            EnsureXml10Text(value, parameterName);
            if (value == "." || value == ".." ||
                value[0] == ' ' || value[value.Length - 1] == ' ' ||
                value[value.Length - 1] == '.')
                throw new ArgumentException("The filename leaf is not portable.", parameterName);

            const string invalid = "<>:\"/\\|?*";
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c < 32 || invalid.IndexOf(c) >= 0)
                    throw new ArgumentException(
                        "A filename leaf cannot contain a path separator, control, or reserved character.",
                        parameterName);
            }

            string stem = value;
            int dot = stem.IndexOf('.');
            if (dot >= 0) stem = stem.Substring(0, dot);
            stem = stem.TrimEnd(' ', '.');
            string upper = stem.ToUpperInvariant();
            if (upper == "CON" || upper == "PRN" || upper == "AUX" || upper == "NUL" ||
                upper == "CONIN$" || upper == "CONOUT$" || upper == "CLOCK$" ||
                IsNumberedDevice(upper, "COM") || IsNumberedDevice(upper, "LPT"))
                throw new ArgumentException("A filename leaf cannot use a reserved device name.", parameterName);
            return value;
        }

        internal static string TimestampText(DateTimeOffset value) =>
            NormalizeTimestamp(value, nameof(value)).ToString(
                "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'",
                CultureInfo.InvariantCulture);

        internal static DateTimeOffset ParseCanonicalTimestamp(string text, string parameterName)
        {
            DateTimeOffset parsed;
            if (!DateTimeOffset.TryParseExact(
                    text,
                    "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out parsed) ||
                !string.Equals(TimestampText(parsed), text, StringComparison.Ordinal))
                throw new ArgumentException("A canonical UTC checkpoint timestamp is required.", parameterName);
            return parsed;
        }

        private static bool IsNumberedDevice(string value, string prefix)
        {
            if (value.Length != prefix.Length + 1 ||
                !value.StartsWith(prefix, StringComparison.Ordinal))
                return false;
            char suffix = value[value.Length - 1];
            return (suffix >= '1' && suffix <= '9') ||
                   suffix == '\u00b9' || suffix == '\u00b2' || suffix == '\u00b3';
        }

        private static void EnsureXml10Text(string value, string parameterName)
        {
            try
            {
                StrictUtf8.GetByteCount(value);
            }
            catch (EncoderFallbackException exception)
            {
                throw new ArgumentException(
                    "Portable journal text must be well-formed Unicode.",
                    parameterName,
                    exception);
            }

            // XML 1.0 Fifth Edition Char production. Strict UTF-8 validation above proves that every surrogate is
            // part of one legal supplementary scalar, so paired surrogates may be skipped here as a unit.
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                if (char.IsHighSurrogate(current))
                {
                    index++;
                    continue;
                }
                bool valid = current == '\u0009' || current == '\u000a' || current == '\u000d' ||
                             (current >= '\u0020' && current <= '\ud7ff') ||
                             (current >= '\ue000' && current <= '\ufffd');
                if (!valid)
                    throw new ArgumentException(
                        "Portable journal text contains a character forbidden by XML 1.0.",
                        parameterName);
            }
        }
    }
}
