using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using ExcelPowerTools.Core.Comparison.Publication;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    public static class MergeRecoveryCheckpointStoreContract
    {
        private static readonly string[] FileNames =
        {
            "00-prepared.merge.xml",
            "01-validated.merge.xml",
            "02-output-published.merge.xml",
            "03-audit-committed.merge.xml",
        };

        public static string FileName(MergeRecoveryJournalState state)
        {
            int index = (int)state - 1;
            if (index < 0 || index >= FileNames.Length)
                throw new ArgumentOutOfRangeException(nameof(state));
            return FileNames[index];
        }

        internal static IReadOnlyList<string> AllFileNames => FileNames;
    }

    /// <summary>
    /// Machine-local path binding kept outside the portable journal. It is intentionally not serializable by the
    /// portable codec and contains no fallback path inference.
    /// </summary>
    public sealed class MergeRecoveryLocalBinding
    {
        public MergeRecoveryLocalBinding(
            string projectId,
            string mergeId,
            string destinationLocatorId,
            string checkpointDirectory,
            string stagePath,
            string destinationPath)
        {
            ProjectId = MergeRecoveryJournalValidation.RequireIdentity(projectId, nameof(projectId));
            MergeId = MergeRecoveryJournalValidation.RequireIdentity(mergeId, nameof(mergeId));
            DestinationLocatorId = MergeRecoveryJournalValidation.RequireIdentity(
                destinationLocatorId,
                nameof(destinationLocatorId));
            CheckpointDirectory = RequireExactAbsolutePath(
                checkpointDirectory,
                nameof(checkpointDirectory));
            StagePath = RequireExactAbsolutePath(stagePath, nameof(stagePath));
            DestinationPath = RequireExactAbsolutePath(destinationPath, nameof(destinationPath));

            string stageDirectory = Path.GetDirectoryName(StagePath);
            string destinationDirectory = Path.GetDirectoryName(DestinationPath);
            if (string.IsNullOrEmpty(stageDirectory) || string.IsNullOrEmpty(destinationDirectory) ||
                !string.Equals(stageDirectory, destinationDirectory, StringComparison.Ordinal))
                throw new ArgumentException(
                    "The local stage and destination must be exact paths in the same directory.",
                    nameof(stagePath));
            if (string.Equals(StagePath, DestinationPath, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    "The local stage and destination paths must not alias each other.",
                    nameof(stagePath));
        }

        public string ProjectId { get; }
        public string MergeId { get; }
        public string DestinationLocatorId { get; }
        public string CheckpointDirectory { get; }
        public string StagePath { get; }
        public string DestinationPath { get; }

        internal void ValidateAgainst(MergeRecoveryJournalReplay replay)
        {
            if (replay == null || replay.Current == null)
                throw Failure(
                    MergeRecoveryCheckpointStoreCodes.BindingMismatch,
                    "",
                    -1,
                    "A local binding requires a non-empty portable recovery journal.");

            MergeRecoveryJournalRecord current = replay.Current;
            if (!string.Equals(ProjectId, current.ProjectId, StringComparison.Ordinal) ||
                !string.Equals(MergeId, current.MergeId, StringComparison.Ordinal) ||
                !string.Equals(
                    DestinationLocatorId,
                    current.DestinationLocatorId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    Path.GetFileName(DestinationPath),
                    current.DestinationLeaf,
                    StringComparison.Ordinal) ||
                (current.Validation != null &&
                 !string.Equals(
                     Path.GetFileName(StagePath),
                     current.Validation.StageLeaf,
                     StringComparison.Ordinal)))
                throw Failure(
                    MergeRecoveryCheckpointStoreCodes.BindingMismatch,
                    "",
                    current.Sequence,
                    "The machine-local paths do not exactly match the portable project, merge, locator, and leaves.");
        }

        private static string RequireExactAbsolutePath(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An exact absolute local path is required.", parameterName);
            string full;
            try
            {
                full = Path.GetFullPath(value);
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is NotSupportedException ||
                exception is PathTooLongException)
            {
                throw new ArgumentException("The local path is invalid.", parameterName, exception);
            }

            if (!Path.IsPathRooted(value) ||
                !string.Equals(value, full, StringComparison.Ordinal) ||
                value.StartsWith(@"\\", StringComparison.Ordinal))
                throw new ArgumentException(
                    "The local path must already be an exact normalized absolute DOS path.",
                    parameterName);
            return full;
        }

        private static MergeRecoveryCheckpointStoreException Failure(
            string code,
            string path,
            int sequence,
            string message) =>
            new MergeRecoveryCheckpointStoreException(code, path, sequence, message);
    }

    public static class MergeRecoveryCheckpointStoreCodes
    {
        public const string BindingMismatch = "CMP-MERGE-CHECKPOINT-BINDING-MISMATCH";
        public const string DirectoryUnavailable = "CMP-MERGE-CHECKPOINT-DIRECTORY-UNAVAILABLE";
        public const string Corrupt = "CMP-MERGE-CHECKPOINT-CORRUPT";
        public const string Gap = "CMP-MERGE-CHECKPOINT-GAP";
        public const string Fork = "CMP-MERGE-CHECKPOINT-FORK";
        public const string Sequence = "CMP-MERGE-CHECKPOINT-SEQUENCE";
        public const string StageWrite = "CMP-MERGE-CHECKPOINT-STAGE-WRITE";
        public const string StageVerification = "CMP-MERGE-CHECKPOINT-STAGE-VERIFICATION";
        public const string Publication = "CMP-MERGE-CHECKPOINT-PUBLICATION";
    }

    public sealed class MergeRecoveryCheckpointStoreException : IOException
    {
        internal MergeRecoveryCheckpointStoreException(
            string code,
            string checkpointPath,
            int sequence,
            string message,
            Exception innerException = null)
            : base(message, innerException)
        {
            Code = code ?? "";
            CheckpointPath = checkpointPath ?? "";
            Sequence = sequence;
        }

        public string Code { get; }
        public string CheckpointPath { get; }
        public int Sequence { get; }
    }

    public enum MergeRecoveryCheckpointWriteStatus
    {
        Published = 0,
        AlreadyPresent = 1,
        Blocked = 2,
        PublicationFailed = 3
    }

    public sealed class MergeRecoveryCheckpointWriteResult
    {
        internal MergeRecoveryCheckpointWriteResult(
            MergeRecoveryCheckpointWriteStatus status,
            string code,
            string message,
            string correctiveAction,
            string checkpointPath,
            string retainedStagePath,
            ValidatedStagePublicationResult publication)
        {
            Status = status;
            Code = code ?? "";
            Message = message ?? "";
            CorrectiveAction = correctiveAction ?? "";
            CheckpointPath = checkpointPath ?? "";
            RetainedStagePath = retainedStagePath ?? "";
            Publication = publication;
        }

        public MergeRecoveryCheckpointWriteStatus Status { get; }
        public bool IsSuccess =>
            Status == MergeRecoveryCheckpointWriteStatus.Published ||
            Status == MergeRecoveryCheckpointWriteStatus.AlreadyPresent;
        public string Code { get; }
        public string Message { get; }
        public string CorrectiveAction { get; }
        public string CheckpointPath { get; }
        public string RetainedStagePath { get; }
        public ValidatedStagePublicationResult Publication { get; }
    }

    public sealed class MergeRecoveryCheckpointSnapshot
    {
        internal MergeRecoveryCheckpointSnapshot(
            IReadOnlyList<MergeRecoveryJournalRecord> records,
            IReadOnlyList<string> checkpointPaths)
        {
            Records = records ?? throw new ArgumentNullException(nameof(records));
            CheckpointPaths = checkpointPaths ?? throw new ArgumentNullException(nameof(checkpointPaths));
        }

        public IReadOnlyList<MergeRecoveryJournalRecord> Records { get; }
        public IReadOnlyList<string> CheckpointPaths { get; }
        public MergeRecoveryJournalRecord Current =>
            Records.Count == 0 ? null : Records[Records.Count - 1];
    }

    /// <summary>
    /// Create-new local store for the four portable cumulative checkpoints. It never overwrites or deletes a file.
    /// A failed or raced stage is retained and reported for diagnosis.
    /// </summary>
    public sealed class MergeRecoveryCheckpointStore
    {
        private const int BufferSize = 64 * 1024;
        private readonly MergeRecoveryLocalBinding _binding;

        public MergeRecoveryCheckpointStore(MergeRecoveryLocalBinding binding)
        {
            _binding = binding ?? throw new ArgumentNullException(nameof(binding));
        }

        public MergeRecoveryCheckpointSnapshot Load()
        {
            RequireCheckpointDirectory();
            var checkpointPaths = new List<string>();
            IReadOnlyList<MergeRecoveryJournalRecord> latest =
                new ReadOnlyCollection<MergeRecoveryJournalRecord>(
                    new List<MergeRecoveryJournalRecord>());
            byte[] priorBytes = null;
            bool gapSeen = false;

            for (int index = 0;
                 index < MergeRecoveryCheckpointStoreContract.AllFileNames.Count;
                 index++)
            {
                string path = Path.Combine(
                    _binding.CheckpointDirectory,
                    MergeRecoveryCheckpointStoreContract.AllFileNames[index]);
                byte[] bytes;
                if (!TryReadBounded(path, out bytes))
                {
                    gapSeen = true;
                    continue;
                }
                if (gapSeen)
                    throw Failure(
                        MergeRecoveryCheckpointStoreCodes.Gap,
                        path,
                        index,
                        "A later recovery checkpoint exists after a missing earlier checkpoint.");

                IReadOnlyList<MergeRecoveryJournalRecord> records = DecodeCheckpoint(path, bytes, index);
                _binding.ValidateAgainst(MergeRecoveryJournal.Replay(records));
                if (priorBytes != null)
                {
                    byte[] expectedPrior = MergeRecoveryJournalCodec.Serialize(records.Take(index));
                    if (!BytesEqual(priorBytes, expectedPrior))
                        throw Failure(
                            MergeRecoveryCheckpointStoreCodes.Fork,
                            path,
                            index,
                            "The checkpoint does not contain the exact preceding cumulative journal.");
                }

                checkpointPaths.Add(path);
                latest = records;
                priorBytes = bytes;
            }

            return new MergeRecoveryCheckpointSnapshot(
                latest,
                new ReadOnlyCollection<string>(checkpointPaths));
        }

        public MergeRecoveryCheckpointWriteResult Append(
            IEnumerable<MergeRecoveryJournalRecord> records)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            MergeRecoveryJournalReplay requested = MergeRecoveryJournal.Replay(records);
            _binding.ValidateAgainst(requested);
            int sequence = requested.Current.Sequence;
            string checkpointPath = Path.Combine(
                _binding.CheckpointDirectory,
                MergeRecoveryCheckpointStoreContract.FileName(requested.Current.State));
            byte[] expectedBytes = MergeRecoveryJournalCodec.Serialize(requested.Records);

            try
            {
                MergeRecoveryCheckpointSnapshot existing = Load();
                if (existing.Records.Count >= requested.Records.Count)
                {
                    byte[] existingTarget;
                    if (TryReadBounded(checkpointPath, out existingTarget) &&
                        BytesEqual(existingTarget, expectedBytes))
                        return Result(
                            MergeRecoveryCheckpointWriteStatus.AlreadyPresent,
                            "",
                            "The exact cumulative checkpoint is already present.",
                            "No action is required.",
                            checkpointPath,
                            "",
                            null);
                    return Result(
                        MergeRecoveryCheckpointWriteStatus.Blocked,
                        MergeRecoveryCheckpointStoreCodes.Fork,
                        "A different checkpoint already occupies this sequence.",
                        "Do not overwrite or delete it; inspect the competing recovery chain.",
                        checkpointPath,
                        "",
                        null);
                }

                if (existing.Records.Count != requested.Records.Count - 1)
                    return Result(
                        MergeRecoveryCheckpointWriteStatus.Blocked,
                        MergeRecoveryCheckpointStoreCodes.Sequence,
                        "The requested checkpoint is not the next create-new recovery state.",
                        "Load the current journal and append only its exact next state.",
                        checkpointPath,
                        "",
                        null);

                if (existing.Records.Count > 0)
                {
                    string priorPath = Path.Combine(
                        _binding.CheckpointDirectory,
                        MergeRecoveryCheckpointStoreContract.FileName(existing.Current.State));
                    byte[] actualPrior;
                    byte[] expectedPrior =
                        MergeRecoveryJournalCodec.Serialize(requested.Records.Take(sequence));
                    if (!TryReadBounded(priorPath, out actualPrior) ||
                        !BytesEqual(actualPrior, expectedPrior))
                        return Result(
                            MergeRecoveryCheckpointWriteStatus.Blocked,
                            MergeRecoveryCheckpointStoreCodes.Fork,
                            "The requested journal does not extend the exact stored checkpoint.",
                            "Do not write a fork; reload and investigate the conflicting chain.",
                            checkpointPath,
                            "",
                            null);
                }
            }
            catch (MergeRecoveryCheckpointStoreException exception)
            {
                return Result(
                    MergeRecoveryCheckpointWriteStatus.Blocked,
                    exception.Code,
                    exception.Message,
                    "Do not overwrite, delete, or guess around the checkpoint; inspect the verbose log.",
                    checkpointPath,
                    "",
                    null);
            }

            string stagePath = Path.Combine(
                _binding.CheckpointDirectory,
                Path.GetFileName(checkpointPath) + "." + Guid.NewGuid().ToString("N") + ".stage");
            try
            {
                WriteCreateNewFlushed(stagePath, expectedBytes);
                byte[] stagedBytes;
                if (!TryReadBounded(stagePath, out stagedBytes) ||
                    !BytesEqual(expectedBytes, stagedBytes))
                    return Result(
                        MergeRecoveryCheckpointWriteStatus.Blocked,
                        MergeRecoveryCheckpointStoreCodes.StageVerification,
                        "The flushed checkpoint stage could not be reread byte-for-byte.",
                        "Retain the stage for diagnosis and start a new checkpoint attempt.",
                        checkpointPath,
                        stagePath,
                        null);
                DecodeCheckpoint(stagePath, stagedBytes, sequence);
            }
            catch (Exception exception) when (
                exception is IOException ||
                exception is UnauthorizedAccessException ||
                exception is MergeRecoveryJournalException)
            {
                return Result(
                    MergeRecoveryCheckpointWriteStatus.Blocked,
                    MergeRecoveryCheckpointStoreCodes.StageWrite,
                    "The checkpoint stage could not be created and verified: " + exception.Message,
                    "Retain any stage file for diagnosis and correct the reported filesystem condition.",
                    checkpointPath,
                    File.Exists(stagePath) ? stagePath : "",
                    null);
            }

            string sha256 = ComputeSha256(expectedBytes);
            ValidatedStagePublicationResult publication = ValidatedStageFilePublisher.Publish(
                stagePath,
                checkpointPath,
                expectedBytes.LongLength,
                sha256);
            if (publication.IsPublished)
            {
                try
                {
                    byte[] publishedBytes;
                    if (!TryReadBounded(checkpointPath, out publishedBytes) ||
                        !BytesEqual(expectedBytes, publishedBytes))
                        throw Failure(
                            MergeRecoveryCheckpointStoreCodes.Corrupt,
                            checkpointPath,
                            sequence,
                            "The published checkpoint did not reread as the expected canonical bytes.");
                    Load();
                }
                catch (MergeRecoveryCheckpointStoreException exception)
                {
                    return Result(
                        MergeRecoveryCheckpointWriteStatus.Blocked,
                        exception.Code,
                        exception.Message,
                        "Quarantine the project store and inspect the post-publication verification failure.",
                        checkpointPath,
                        "",
                        publication);
                }
                return Result(
                    MergeRecoveryCheckpointWriteStatus.Published,
                    "",
                    "The cumulative checkpoint was published create-new.",
                    "No action is required.",
                    checkpointPath,
                    "",
                    publication);
            }

            if (publication.Status == ValidatedStagePublicationStatus.DestinationAlreadyExists ||
                publication.Status == ValidatedStagePublicationStatus.DestinationRaceLost)
            {
                try
                {
                    byte[] racedBytes;
                    if (TryReadBounded(checkpointPath, out racedBytes) &&
                        BytesEqual(expectedBytes, racedBytes))
                        return Result(
                            MergeRecoveryCheckpointWriteStatus.AlreadyPresent,
                            "",
                            "A concurrent writer published the exact same cumulative checkpoint.",
                            "No journal action is required; the retained stage may be quarantined manually.",
                            checkpointPath,
                            File.Exists(stagePath) ? stagePath : "",
                            publication);
                }
                catch (MergeRecoveryCheckpointStoreException)
                {
                    // The structured fork result below is deliberately more useful than an incidental read error.
                }
                return Result(
                    MergeRecoveryCheckpointWriteStatus.Blocked,
                    MergeRecoveryCheckpointStoreCodes.Fork,
                    "Publication lost a race to a different or unreadable checkpoint.",
                    "Do not overwrite or delete either artifact; investigate the competing checkpoint.",
                    checkpointPath,
                    File.Exists(stagePath) ? stagePath : "",
                    publication);
            }

            return Result(
                MergeRecoveryCheckpointWriteStatus.PublicationFailed,
                MergeRecoveryCheckpointStoreCodes.Publication,
                publication.Message,
                publication.CorrectiveAction,
                checkpointPath,
                File.Exists(stagePath) ? stagePath : "",
                publication);
        }

        private void RequireCheckpointDirectory()
        {
            try
            {
                if (!Directory.Exists(_binding.CheckpointDirectory))
                    throw Failure(
                        MergeRecoveryCheckpointStoreCodes.DirectoryUnavailable,
                        _binding.CheckpointDirectory,
                        -1,
                        "The exact local checkpoint directory does not exist.");
            }
            catch (Exception exception) when (
                exception is IOException ||
                exception is UnauthorizedAccessException)
            {
                throw Failure(
                    MergeRecoveryCheckpointStoreCodes.DirectoryUnavailable,
                    _binding.CheckpointDirectory,
                    -1,
                    "The exact local checkpoint directory is unavailable.",
                    exception);
            }
        }

        private static IReadOnlyList<MergeRecoveryJournalRecord> DecodeCheckpoint(
            string path,
            byte[] bytes,
            int sequence)
        {
            try
            {
                IReadOnlyList<MergeRecoveryJournalRecord> records =
                    MergeRecoveryJournalCodec.Deserialize(bytes);
                if (records.Count != sequence + 1 ||
                    records[records.Count - 1].Sequence != sequence ||
                    records[records.Count - 1].State !=
                    (MergeRecoveryJournalState)(sequence + 1))
                    throw Failure(
                        MergeRecoveryCheckpointStoreCodes.Sequence,
                        path,
                        sequence,
                        "The checkpoint filename does not match its cumulative journal state.");
                return records;
            }
            catch (MergeRecoveryCheckpointStoreException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is MergeRecoveryJournalException ||
                exception is ArgumentException)
            {
                throw Failure(
                    MergeRecoveryCheckpointStoreCodes.Corrupt,
                    path,
                    sequence,
                    "The checkpoint is not a canonical valid recovery journal: " + exception.Message,
                    exception);
            }
        }

        private static bool TryReadBounded(string path, out byte[] bytes)
        {
            bytes = null;
            try
            {
                using (var stream = new FileStream(
                           path,
                           FileMode.Open,
                           FileAccess.Read,
                           FileShare.Read,
                           BufferSize,
                           FileOptions.SequentialScan))
                {
                    if (stream.Length < 0 ||
                        stream.Length > MergeRecoveryJournalContract.MaxJournalBytes)
                        throw Failure(
                            MergeRecoveryCheckpointStoreCodes.Corrupt,
                            path,
                            -1,
                            "The checkpoint exceeds the bounded portable journal size.");
                    bytes = new byte[(int)stream.Length];
                    int offset = 0;
                    while (offset < bytes.Length)
                    {
                        int read = stream.Read(bytes, offset, bytes.Length - offset);
                        if (read == 0)
                            throw new EndOfStreamException("The checkpoint ended before its declared length.");
                        offset += read;
                    }
                    if (stream.ReadByte() != -1)
                        throw new IOException("The checkpoint changed length while it was read.");
                    return true;
                }
            }
            catch (FileNotFoundException)
            {
                return false;
            }
            catch (DirectoryNotFoundException)
            {
                return false;
            }
            catch (MergeRecoveryCheckpointStoreException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is IOException ||
                exception is UnauthorizedAccessException)
            {
                throw Failure(
                    MergeRecoveryCheckpointStoreCodes.Corrupt,
                    path,
                    -1,
                    "The checkpoint could not be read under an exclusive write/delete denial: " +
                    exception.Message,
                    exception);
            }
        }

        private static void WriteCreateNewFlushed(string path, byte[] bytes)
        {
            using (var stream = new FileStream(
                       path,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       BufferSize,
                       FileOptions.SequentialScan | FileOptions.WriteThrough))
            {
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(true);
            }
        }

        private static string ComputeSha256(byte[] bytes)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] digest = sha.ComputeHash(bytes);
                const string alphabet = "0123456789abcdef";
                var chars = new char[digest.Length * 2];
                for (int index = 0; index < digest.Length; index++)
                {
                    chars[index * 2] = alphabet[digest[index] >> 4];
                    chars[index * 2 + 1] = alphabet[digest[index] & 0x0f];
                }
                return new string(chars);
            }
        }

        private static bool BytesEqual(byte[] left, byte[] right)
        {
            if (ReferenceEquals(left, right)) return true;
            if (left == null || right == null || left.Length != right.Length) return false;
            int difference = 0;
            for (int index = 0; index < left.Length; index++)
                difference |= left[index] ^ right[index];
            return difference == 0;
        }

        private static MergeRecoveryCheckpointWriteResult Result(
            MergeRecoveryCheckpointWriteStatus status,
            string code,
            string message,
            string correctiveAction,
            string checkpointPath,
            string retainedStagePath,
            ValidatedStagePublicationResult publication) =>
            new MergeRecoveryCheckpointWriteResult(
                status,
                code,
                message,
                correctiveAction,
                checkpointPath,
                retainedStagePath,
                publication);

        private static MergeRecoveryCheckpointStoreException Failure(
            string code,
            string path,
            int sequence,
            string message,
            Exception innerException = null) =>
            new MergeRecoveryCheckpointStoreException(
                code,
                path,
                sequence,
                message,
                innerException);
    }
}
