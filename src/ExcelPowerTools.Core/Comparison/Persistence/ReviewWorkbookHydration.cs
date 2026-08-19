using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Microsoft.Win32.SafeHandles;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>
    /// Immutable, read-only hydration of one exact Review.xlsx occurrence. Hydration retains no file handle;
    /// callers must revalidate the current path occurrence before using state after an interaction boundary.
    /// </summary>
    public sealed class ReviewWorkbookHydrationSession
    {
        internal ReviewWorkbookHydrationSession(
            string packageSha256,
            string fileOccurrenceToken,
            ReviewWorkbookReadResult readResult,
            ReviewWorkbookValidationSummary summary)
        {
            PackageSha256 = packageSha256 ?? throw new ArgumentNullException(nameof(packageSha256));
            FileOccurrenceToken = fileOccurrenceToken ??
                                  throw new ArgumentNullException(nameof(fileOccurrenceToken));
            ReadResult = readResult ?? throw new ArgumentNullException(nameof(readResult));
            Summary = summary ?? throw new ArgumentNullException(nameof(summary));
        }

        public string FullPath => Summary.FullPath;
        public string FileName => Summary.FileName;
        public long FileByteLength => Summary.FileByteLength;
        public DateTimeOffset ModifiedUtc => Summary.ModifiedUtc;

        /// <summary>SHA-256 of every package byte read from the same protected handle that was validated.</summary>
        public string PackageSha256 { get; }

        /// <summary>
        /// Stable Windows volume/file identity captured from the validated handle. It is empty only when the
        /// portable non-Windows fallback cannot provide an equivalent identity.
        /// </summary>
        public string FileOccurrenceToken { get; }
        public bool HasStableOccurrenceToken => FileOccurrenceToken.Length != 0;

        /// <summary>
        /// Complete validated state, including StateJson and the deterministically replayed decision generation.
        /// </summary>
        public ReviewWorkbookReadResult ReadResult { get; }

        /// <summary>Privacy-safe metadata suitable for display without exposing embedded project state.</summary>
        public ReviewWorkbookValidationSummary Summary { get; }

        /// <summary>
        /// Hydrates the path again and fails unless its package and validated identities still match this session.
        /// The returned session represents the newly protected occurrence and also retains no handle.
        /// </summary>
        public ReviewWorkbookHydrationSession RevalidateCurrentOccurrence() =>
            ReviewWorkbookHydrationService.RevalidateCurrentOccurrence(this);
    }

    public static class ReviewWorkbookHydrationService
    {
        private const int HashBufferSize = 64 * 1024;

        /// <summary>
        /// Validates and hashes one exact path occurrence under a single read-only handle that shares only reads.
        /// The handle denies ordinary writes, deletes and replacements until all retained data has been captured.
        /// </summary>
        public static ReviewWorkbookHydrationSession Hydrate(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
                throw new ArgumentException("A Review workbook path is required.", nameof(workbookPath));

            string path = Path.GetFullPath(workbookPath);
            using (var stream = new FileStream(
                       path,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.Read,
                       HashBufferSize,
                       FileOptions.RandomAccess))
            {
                long fileBytes = stream.Length;
                if (fileBytes < 0 || fileBytes > ReviewWorkbookFile.MaxReviewPackageBytes)
                    throw new InvalidDataException(
                        "The Review workbook exceeds the safe hydration package-size limit.");

                ReviewWorkbookReadResult readResult = ReviewWorkbookFile.ReadAndValidate(stream);
                string packageSha256 = ComputeBoundedPackageSha256(stream, fileBytes);
                if (stream.Length != fileBytes)
                    throw new IOException("The Review workbook changed while it was being hydrated.");

                ReviewWorkbookHandleMetadata metadata =
                    ReviewWorkbookHandleMetadataReader.Capture(stream, path);
                if (metadata.FileByteLength != fileBytes)
                    throw new IOException("The selected Review workbook occurrence changed during hydration.");
                ReviewWorkbookValidationSummary summary =
                    ReviewWorkbookValidationSummaryService.CreateValidatedSummary(
                        path,
                        fileBytes,
                        metadata.ModifiedUtc,
                        readResult);
                return new ReviewWorkbookHydrationSession(
                    packageSha256,
                    metadata.FileOccurrenceToken,
                    readResult,
                    summary);
            }
        }

        /// <summary>
        /// Reopens the selected path and fails closed unless package, project, state, projection and decision
        /// identities all still describe the hydrated occurrence. Timestamps are retained metadata, not proof.
        /// </summary>
        public static ReviewWorkbookHydrationSession RevalidateCurrentOccurrence(
            ReviewWorkbookHydrationSession expected)
        {
            if (expected == null) throw new ArgumentNullException(nameof(expected));
            ReviewWorkbookHydrationSession current = Hydrate(expected.FullPath);
            if (!HasSameCurrentOccurrence(expected, current))
                throw new InvalidDataException(
                    "The selected Review workbook no longer matches the hydrated package and state identities.");
            return current;
        }

        internal static bool HasSameCurrentOccurrence(
            ReviewWorkbookHydrationSession expected,
            ReviewWorkbookHydrationSession current)
        {
            if (expected == null || current == null) return false;
            if (!string.Equals(expected.FullPath, current.FullPath, StringComparison.OrdinalIgnoreCase) ||
                expected.FileByteLength != current.FileByteLength ||
                !DigestEquals(expected.PackageSha256, current.PackageSha256))
                return false;
            if (expected.HasStableOccurrenceToken != current.HasStableOccurrenceToken ||
                (expected.HasStableOccurrenceToken &&
                 !string.Equals(
                     expected.FileOccurrenceToken,
                     current.FileOccurrenceToken,
                     StringComparison.Ordinal)))
                return false;
            return HasSameValidatedIdentity(expected.ReadResult, current.ReadResult);
        }

        internal static bool HasSameValidatedIdentity(
            ReviewWorkbookReadResult expected,
            ReviewWorkbookReadResult current)
        {
            if (expected == null || current == null) return false;
            if (!string.Equals(expected.ProjectId, current.ProjectId, StringComparison.Ordinal) ||
                !string.Equals(expected.ManifestItemId, current.ManifestItemId, StringComparison.Ordinal) ||
                !string.Equals(expected.StateJson, current.StateJson, StringComparison.Ordinal) ||
                !DigestEquals(expected.StateSha256, current.StateSha256) ||
                !DigestEquals(expected.ProjectionSha256, current.ProjectionSha256) ||
                !expected.ChunkItemIds.SequenceEqual(current.ChunkItemIds, StringComparer.Ordinal) ||
                expected.HasDecisionState != current.HasDecisionState ||
                !string.Equals(expected.DecisionItemId, current.DecisionItemId, StringComparison.Ordinal) ||
                !DigestEquals(expected.DecisionSha256, current.DecisionSha256))
                return false;

            if (!expected.HasDecisionState) return true;
            return string.Equals(
                       expected.DecisionGeneration.ProjectId,
                       current.DecisionGeneration.ProjectId,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       expected.DecisionGeneration.CatalogId,
                       current.DecisionGeneration.CatalogId,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       expected.DecisionGeneration.BaseGenerationId,
                       current.DecisionGeneration.BaseGenerationId,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       expected.DecisionGeneration.MappingGenerationId,
                       current.DecisionGeneration.MappingGenerationId,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       expected.DecisionGeneration.GenerationId,
                       current.DecisionGeneration.GenerationId,
                       StringComparison.Ordinal) &&
                   string.Equals(
                       expected.DecisionGeneration.PriorGenerationId,
                       current.DecisionGeneration.PriorGenerationId,
                       StringComparison.Ordinal);
        }

        internal static string ComputeBoundedPackageSha256(Stream stream, long expectedBytes)
        {
            if (expectedBytes < 0 || expectedBytes > ReviewWorkbookFile.MaxReviewPackageBytes)
                throw new InvalidDataException(
                    "The Review workbook exceeds the safe hydration hash-size limit.");

            long originalPosition = stream.Position;
            try
            {
                stream.Position = 0;
                using (SHA256 sha256 = SHA256.Create())
                {
                    var buffer = new byte[HashBufferSize];
                    long hashed = 0;
                    while (hashed < expectedBytes)
                    {
                        int requested = (int)Math.Min(buffer.Length, expectedBytes - hashed);
                        int read = stream.Read(buffer, 0, requested);
                        if (read == 0)
                            throw new EndOfStreamException(
                                "The Review workbook ended before its protected byte length was hashed.");
                        sha256.TransformBlock(buffer, 0, read, null, 0);
                        hashed = checked(hashed + read);
                    }
                    if (stream.ReadByte() != -1 || stream.Length != expectedBytes)
                        throw new IOException("The Review workbook changed while its package hash was computed.");
                    sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                    return BytesToHex(sha256.Hash);
                }
            }
            finally
            {
                try { stream.Position = originalPosition; }
                catch { }
            }
        }

        private static bool DigestEquals(string expected, string current) =>
            string.Equals(expected ?? "", current ?? "", StringComparison.OrdinalIgnoreCase);

        private static string BytesToHex(byte[] bytes)
        {
            if (bytes == null || bytes.Length != 32)
                throw new InvalidDataException("A complete Review workbook SHA-256 digest was not produced.");
            var chars = new char[bytes.Length * 2];
            const string alphabet = "0123456789abcdef";
            for (int index = 0; index < bytes.Length; index++)
            {
                chars[index * 2] = alphabet[bytes[index] >> 4];
                chars[index * 2 + 1] = alphabet[bytes[index] & 0x0F];
            }
            return new string(chars);
        }
    }

    internal sealed class ReviewWorkbookHandleMetadata
    {
        public ReviewWorkbookHandleMetadata(
            long fileByteLength,
            DateTimeOffset modifiedUtc,
            string fileOccurrenceToken)
        {
            FileByteLength = fileByteLength;
            ModifiedUtc = modifiedUtc;
            FileOccurrenceToken = fileOccurrenceToken ?? "";
        }

        public long FileByteLength { get; }
        public DateTimeOffset ModifiedUtc { get; }
        public string FileOccurrenceToken { get; }
    }

    internal static class ReviewWorkbookHandleMetadataReader
    {
        /// <summary>
        /// Windows supplies size, timestamp and stable volume/file identity from the already-open handle. Core is
        /// also testable on non-Windows platforms, where the portable fallback has to query path metadata and
        /// deliberately returns no occurrence token; package and validated-content identities still apply there.
        /// </summary>
        public static ReviewWorkbookHandleMetadata Capture(FileStream stream, string fallbackPath)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                return CaptureWindows(stream.SafeFileHandle);
            return CapturePortableFallback(stream, fallbackPath);
        }

        private static ReviewWorkbookHandleMetadata CaptureWindows(SafeFileHandle handle)
        {
            if (handle == null || handle.IsInvalid || handle.IsClosed)
                throw new IOException("The Review workbook handle is unavailable for metadata capture.");
            if (!GetFileInformationByHandle(handle, out ByHandleFileInformation information))
            {
                int error = Marshal.GetLastWin32Error();
                throw new IOException(
                    "Windows could not capture metadata from the protected Review workbook handle.",
                    new Win32Exception(error));
            }

            ulong unsignedLength = ((ulong)information.FileSizeHigh << 32) | information.FileSizeLow;
            if (unsignedLength > long.MaxValue)
                throw new IOException("The Review workbook handle returned an invalid file length.");
            ulong fileIndex = ((ulong)information.FileIndexHigh << 32) | information.FileIndexLow;
            string occurrenceToken = fileIndex == 0
                ? ""
                : information.VolumeSerialNumber.ToString("x8", CultureInfo.InvariantCulture) + ":" +
                  fileIndex.ToString("x16", CultureInfo.InvariantCulture);
            long lastWriteFileTime = checked(
                ((long)information.LastWriteTimeHigh << 32) | information.LastWriteTimeLow);
            DateTimeOffset modifiedUtc;
            try
            {
                modifiedUtc = new DateTimeOffset(DateTime.FromFileTimeUtc(lastWriteFileTime));
            }
            catch (ArgumentOutOfRangeException exception)
            {
                throw new IOException(
                    "The Review workbook handle returned an invalid modified timestamp.",
                    exception);
            }
            return new ReviewWorkbookHandleMetadata((long)unsignedLength, modifiedUtc, occurrenceToken);
        }

        private static ReviewWorkbookHandleMetadata CapturePortableFallback(
            FileStream stream,
            string fallbackPath)
        {
            if (string.IsNullOrWhiteSpace(fallbackPath))
                throw new ArgumentException("A fallback Review workbook path is required.", nameof(fallbackPath));
            var file = new FileInfo(fallbackPath);
            file.Refresh();
            if (!file.Exists || file.Length != stream.Length)
                throw new IOException("The selected Review workbook occurrence changed during hydration.");
            return new ReviewWorkbookHandleMetadata(
                stream.Length,
                new DateTimeOffset(file.LastWriteTimeUtc),
                "");
        }

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(
            SafeFileHandle fileHandle,
            out ByHandleFileInformation fileInformation);

        [StructLayout(LayoutKind.Sequential)]
        private struct ByHandleFileInformation
        {
            public uint FileAttributes;
            public uint CreationTimeLow;
            public uint CreationTimeHigh;
            public uint LastAccessTimeLow;
            public uint LastAccessTimeHigh;
            public uint LastWriteTimeLow;
            public uint LastWriteTimeHigh;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }
    }
}
