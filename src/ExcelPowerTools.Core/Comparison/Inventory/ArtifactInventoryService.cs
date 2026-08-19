using System;
using System.Globalization;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Inventory
{
    /// <summary>
    /// Inventories an artifact without Excel automation and without modifying it. Exact bytes are hashed through
    /// a bounded buffer, and pre/open/post metadata checks reject files that visibly change during the operation.
    /// </summary>
    public sealed class ArtifactInventoryService
    {
        public async Task<ArtifactInventoryRecord> InventoryAsync(
            string path,
            IProgress<ArtifactInventoryProgress> progress = null,
            CancellationToken cancellationToken = default(CancellationToken),
            ArtifactInventoryOptions options = null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            options = options ?? ArtifactInventoryOptions.Default;
            string normalizedPath = NormalizePath(path);

            FileMetadata beforeOpen = CaptureInitialMetadata(normalizedPath);
            byte[] sample;
            string contentSha256;

            try
            {
                using (var stream = new FileStream(
                    normalizedPath,
                    FileMode.Open,
                    FileAccess.Read,
                    // Deny write and delete while the exact byte identity is being established. Metadata-only
                    // rechecks cannot detect a same-length rewrite whose timestamp is restored.
                    FileShare.Read,
                    options.BufferSize,
                    FileOptions.Asynchronous | FileOptions.SequentialScan))
                {
                    FileMetadata afterOpen = CaptureCurrentMetadata(normalizedPath, beforeOpen, "opening");
                    EnsureStable(normalizedPath, beforeOpen, afterOpen, "between metadata capture and opening");

                    progress?.Report(new ArtifactInventoryProgress(
                        normalizedPath,
                        ArtifactInventoryStage.Hashing,
                        0,
                        beforeOpen.Length));
                    cancellationToken.ThrowIfCancellationRequested();

                    HashResult hash = await HashExactLengthAsync(
                        stream,
                        normalizedPath,
                        beforeOpen.Length,
                        options,
                        progress,
                        cancellationToken).ConfigureAwait(false);
                    sample = hash.Sample;
                    contentSha256 = hash.Sha256;

                    FileMetadata afterHash = CaptureCurrentMetadata(normalizedPath, beforeOpen, "hashing");
                    EnsureStable(normalizedPath, beforeOpen, afterHash, "while hashing");

                    progress?.Report(new ArtifactInventoryProgress(
                        normalizedPath,
                        ArtifactInventoryStage.DetectingFormat,
                        beforeOpen.Length,
                        beforeOpen.Length));
                    cancellationToken.ThrowIfCancellationRequested();

                    ArtifactFormatAssessment format = ArtifactFormatDetector.Assess(
                        stream,
                        sample,
                        Path.GetExtension(normalizedPath),
                        cancellationToken);

                    progress?.Report(new ArtifactInventoryProgress(
                        normalizedPath,
                        ArtifactInventoryStage.VerifyingStability,
                        beforeOpen.Length,
                        beforeOpen.Length));
                    cancellationToken.ThrowIfCancellationRequested();

                    FileMetadata afterDetection = CaptureCurrentMetadata(normalizedPath, beforeOpen, "format detection");
                    EnsureStable(normalizedPath, beforeOpen, afterDetection, "while detecting its format");

                    string inventoryFingerprint = StableId.Compute(
                        "artifact-inventory-fingerprint-v1",
                        NormalizePathForIdentity(normalizedPath),
                        beforeOpen.Length.ToString(CultureInfo.InvariantCulture),
                        beforeOpen.CreationTimeUtc.UtcDateTime.Ticks.ToString(CultureInfo.InvariantCulture),
                        beforeOpen.LastWriteTimeUtc.UtcDateTime.Ticks.ToString(CultureInfo.InvariantCulture),
                        contentSha256);

                    var result = new ArtifactInventoryRecord(
                        inventoryFingerprint,
                        contentSha256,
                        normalizedPath,
                        beforeOpen.Length,
                        beforeOpen.CreationTimeUtc,
                        beforeOpen.LastWriteTimeUtc,
                        format);

                    progress?.Report(new ArtifactInventoryProgress(
                        normalizedPath,
                        ArtifactInventoryStage.Completed,
                        beforeOpen.Length,
                        beforeOpen.Length));
                    return result;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (ArtifactInventoryException)
            {
                throw;
            }
            catch (UnauthorizedAccessException exception)
            {
                throw Unreadable(normalizedPath, exception);
            }
            catch (SecurityException exception)
            {
                throw Unreadable(normalizedPath, exception);
            }
            catch (IOException exception)
            {
                throw Unreadable(normalizedPath, exception);
            }
        }

        private static async Task<HashResult> HashExactLengthAsync(
            Stream stream,
            string normalizedPath,
            long expectedLength,
            ArtifactInventoryOptions options,
            IProgress<ArtifactInventoryProgress> progress,
            CancellationToken cancellationToken)
        {
            var buffer = new byte[options.BufferSize];
            int sampleCapacity = (int)Math.Min((long)options.ContentSampleBytes, expectedLength);
            var sample = new byte[sampleCapacity];
            int sampleLength = 0;
            long totalRead = 0;
            long nextProgress = Math.Min(options.ProgressIntervalBytes, expectedLength);

            using (SHA256 sha = SHA256.Create())
            {
                while (totalRead < expectedLength)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int wanted = (int)Math.Min(buffer.Length, expectedLength - totalRead);
                    int read = await stream.ReadAsync(buffer, 0, wanted, cancellationToken).ConfigureAwait(false);
                    if (read == 0)
                    {
                        throw Changed(normalizedPath,
                            "The artifact became shorter while it was being read. Wait for all writers or sync operations to finish, then retry inventory.");
                    }

                    sha.TransformBlock(buffer, 0, read, buffer, 0);
                    if (sampleLength < sample.Length)
                    {
                        int copy = Math.Min(read, sample.Length - sampleLength);
                        Buffer.BlockCopy(buffer, 0, sample, sampleLength, copy);
                        sampleLength += copy;
                    }

                    totalRead += read;
                    if (progress != null && (totalRead >= nextProgress || totalRead == expectedLength))
                    {
                        progress.Report(new ArtifactInventoryProgress(
                            normalizedPath,
                            ArtifactInventoryStage.Hashing,
                            totalRead,
                            expectedLength));
                        cancellationToken.ThrowIfCancellationRequested();

                        if (nextProgress < expectedLength)
                        {
                            long intervals = totalRead / options.ProgressIntervalBytes + 1;
                            nextProgress = Math.Min(expectedLength, intervals * options.ProgressIntervalBytes);
                        }
                    }
                }

                sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                string sha256 = ToHex(sha.Hash);

                if (sampleLength == sample.Length)
                    return new HashResult(sha256, sample);

                var exactSample = new byte[sampleLength];
                Buffer.BlockCopy(sample, 0, exactSample, 0, sampleLength);
                return new HashResult(sha256, exactSample);
            }
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArtifactInventoryException(
                    ArtifactInventoryErrorCode.InvalidPath,
                    path,
                    "An artifact path is required. Select an existing local file and try again.");
            }

            try
            {
                return Path.GetFullPath(path);
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is NotSupportedException ||
                exception is PathTooLongException ||
                exception is SecurityException)
            {
                throw new ArtifactInventoryException(
                    ArtifactInventoryErrorCode.InvalidPath,
                    path,
                    "The artifact path is invalid or cannot be resolved. Select an existing local file and try again.",
                    exception);
            }
        }

        private static FileMetadata CaptureInitialMetadata(string normalizedPath)
        {
            if (Directory.Exists(normalizedPath))
            {
                throw new ArtifactInventoryException(
                    ArtifactInventoryErrorCode.PathIsNotAFile,
                    normalizedPath,
                    "The selected artifact path is a folder, not a file. Select one workbook or separated-value file.");
            }

            try
            {
                var file = new FileInfo(normalizedPath);
                file.Refresh();
                if (!file.Exists)
                {
                    throw new ArtifactInventoryException(
                        ArtifactInventoryErrorCode.MissingArtifact,
                        normalizedPath,
                        "The selected artifact was not found at '" + normalizedPath +
                        "'. Verify the path or reconnect the source, then try again.");
                }

                return FileMetadata.From(file);
            }
            catch (ArtifactInventoryException)
            {
                throw;
            }
            catch (UnauthorizedAccessException exception)
            {
                throw Unreadable(normalizedPath, exception);
            }
            catch (SecurityException exception)
            {
                throw Unreadable(normalizedPath, exception);
            }
            catch (IOException exception)
            {
                throw Unreadable(normalizedPath, exception);
            }
        }

        private static FileMetadata CaptureCurrentMetadata(
            string normalizedPath,
            FileMetadata baseline,
            string operation)
        {
            try
            {
                var file = new FileInfo(normalizedPath);
                file.Refresh();
                if (!file.Exists)
                {
                    throw Changed(normalizedPath,
                        "The artifact disappeared during " + operation +
                        ". Restore or finish moving/syncing the file, then retry inventory.");
                }
                return FileMetadata.From(file);
            }
            catch (ArtifactInventoryException)
            {
                throw;
            }
            catch (FileNotFoundException exception)
            {
                throw Changed(normalizedPath,
                    "The artifact disappeared during " + operation +
                    ". Restore or finish moving/syncing the file, then retry inventory.", exception);
            }
            catch (DirectoryNotFoundException exception)
            {
                throw Changed(normalizedPath,
                    "The artifact folder disappeared during " + operation +
                    ". Restore or reconnect the folder, then retry inventory.", exception);
            }
            catch (IOException exception)
            {
                throw Changed(normalizedPath,
                    "The artifact could not be rechecked during " + operation +
                    ". Wait for writers or sync operations to finish, then retry inventory.", exception);
            }
        }

        private static void EnsureStable(
            string normalizedPath,
            FileMetadata expected,
            FileMetadata actual,
            string interval)
        {
            if (expected.Length == actual.Length &&
                expected.CreationTimeUtc == actual.CreationTimeUtc &&
                expected.LastWriteTimeUtc == actual.LastWriteTimeUtc)
                return;

            throw Changed(normalizedPath,
                "The artifact's byte length or timestamps changed " + interval +
                ". Wait for all editors and sync operations to finish, then retry so the comparison uses one immutable revision.");
        }

        private static ArtifactInventoryException Unreadable(string normalizedPath, Exception innerException)
        {
            return new ArtifactInventoryException(
                ArtifactInventoryErrorCode.UnreadableArtifact,
                normalizedPath,
                "The artifact at '" + normalizedPath +
                "' could not be read. Close applications holding an exclusive lock, verify file permissions and connectivity, then retry.",
                innerException);
        }

        private static ArtifactInventoryException Changed(
            string normalizedPath,
            string message,
            Exception innerException = null)
        {
            return new ArtifactInventoryException(
                ArtifactInventoryErrorCode.ArtifactChangedDuringRead,
                normalizedPath,
                message,
                innerException);
        }

        private static string NormalizePathForIdentity(string normalizedPath)
        {
            return Path.DirectorySeparatorChar == '\\'
                ? normalizedPath.ToUpperInvariant()
                : normalizedPath;
        }

        private static string ToHex(byte[] bytes)
        {
            var characters = new char[bytes.Length * 2];
            const string hex = "0123456789ABCDEF";
            for (int i = 0; i < bytes.Length; i++)
            {
                characters[i * 2] = hex[bytes[i] >> 4];
                characters[i * 2 + 1] = hex[bytes[i] & 0x0F];
            }
            return new string(characters);
        }

        private struct FileMetadata
        {
            internal FileMetadata(long length, DateTimeOffset creationTimeUtc, DateTimeOffset lastWriteTimeUtc)
            {
                Length = length;
                CreationTimeUtc = creationTimeUtc;
                LastWriteTimeUtc = lastWriteTimeUtc;
            }

            internal long Length { get; }
            internal DateTimeOffset CreationTimeUtc { get; }
            internal DateTimeOffset LastWriteTimeUtc { get; }

            internal static FileMetadata From(FileInfo file)
            {
                return new FileMetadata(
                    file.Length,
                    new DateTimeOffset(file.CreationTimeUtc, TimeSpan.Zero),
                    new DateTimeOffset(file.LastWriteTimeUtc, TimeSpan.Zero));
            }
        }

        private struct HashResult
        {
            internal HashResult(string sha256, byte[] sample)
            {
                Sha256 = sha256;
                Sample = sample;
            }

            internal string Sha256 { get; }
            internal byte[] Sample { get; }
        }
    }
}
