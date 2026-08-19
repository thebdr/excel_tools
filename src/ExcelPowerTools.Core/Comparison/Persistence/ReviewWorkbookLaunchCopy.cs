using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>
    /// Owns a validated, private, read-only launch copy of a comparison review workbook. The retained read
    /// handle deliberately shares only read access, preventing ordinary writers, deletes and file renames
    /// until the owner disposes this object after Excel has released the workbook.
    /// </summary>
    public sealed class ReviewWorkbookLaunchCopy : IDisposable
    {
        private const int CopyBufferSize = 64 * 1024;
        private const string RootFolderName = "ReviewLaunch";
        private const string RootVersion = "v1";
        private FileStream _protectionStream;
        private bool _disposed;

        private ReviewWorkbookLaunchCopy(
            string sourcePath,
            string launchDirectory,
            string launchPath,
            long byteLength,
            string packageSha256,
            ReviewWorkbookReadResult validation,
            FileStream protectionStream)
        {
            SourcePath = sourcePath;
            LaunchDirectory = launchDirectory;
            LaunchPath = launchPath;
            ByteLength = byteLength;
            PackageSha256 = packageSha256;
            Validation = validation;
            _protectionStream = protectionStream;
        }

        public string SourcePath { get; }
        public string LaunchDirectory { get; }
        public string LaunchPath { get; }
        public long ByteLength { get; }
        public string PackageSha256 { get; }
        public ReviewWorkbookReadResult Validation { get; }

        /// <summary>Returns the per-user root used when a caller does not provide an isolated test root.</summary>
        public static string DefaultRootDirectory
        {
            get
            {
                string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(local)) local = Path.GetTempPath();
                return Path.Combine(local, "ExcelPowerTools", RootFolderName, RootVersion);
            }
        }

        /// <summary>
        /// Copies one exact, write-locked source stream into a random private directory, marks the result
        /// read-only, validates it through the retained protective handle, and returns ownership to the caller.
        /// No source file is changed.
        /// </summary>
        public static ReviewWorkbookLaunchCopy Create(string sourcePath, string rootDirectory = null)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
                throw new ArgumentException("A review workbook source path is required.", nameof(sourcePath));

            string source = Path.GetFullPath(sourcePath);
            if (!File.Exists(source))
                throw new FileNotFoundException("The review workbook source was not found.", source);

            string root = Path.GetFullPath(
                string.IsNullOrWhiteSpace(rootDirectory) ? DefaultRootDirectory : rootDirectory);
            Directory.CreateDirectory(root);
            RejectReparsePoint(root, "The review launch root cannot be a reparse point.");

            string token = Guid.NewGuid().ToString("N");
            string launchDirectory = Path.Combine(root, "launch-" + token);
            string launchPath = Path.Combine(
                launchDirectory,
                BuildLaunchFileName(source, token.Substring(0, 10)));
            FileStream protectionStream = null;

            try
            {
                Directory.CreateDirectory(launchDirectory);
                RejectReparsePoint(
                    launchDirectory,
                    "The private review launch directory unexpectedly became a reparse point.");

                long byteLength;
                string packageSha256;
                using (var input = new FileStream(
                           source,
                           FileMode.Open,
                           FileAccess.Read,
                           FileShare.Read,
                           CopyBufferSize,
                           FileOptions.SequentialScan))
                {
                    byteLength = input.Length;
                    if (byteLength > ReviewWorkbookFile.MaxReviewPackageBytes)
                        throw new InvalidDataException(
                            "The review workbook exceeds the safe launch-copy size limit.");

                    using (var output = new FileStream(
                               launchPath,
                               FileMode.CreateNew,
                               FileAccess.ReadWrite,
                               FileShare.None,
                               CopyBufferSize,
                               FileOptions.SequentialScan | FileOptions.WriteThrough))
                    using (SHA256 sha256 = SHA256.Create())
                    {
                        var buffer = new byte[CopyBufferSize];
                        long copied = 0;
                        int read;
                        while ((read = input.Read(buffer, 0, buffer.Length)) != 0)
                        {
                            output.Write(buffer, 0, read);
                            sha256.TransformBlock(buffer, 0, read, null, 0);
                            copied = checked(copied + read);
                        }
                        sha256.TransformFinalBlock(Array.Empty<byte>(), 0, 0);

                        if (copied != byteLength || input.Length != byteLength)
                            throw new IOException(
                                "The review workbook changed while its protected launch copy was being created.");

                        output.Flush(true);
                        packageSha256 = BytesToHex(sha256.Hash);
                    }
                }

                FileAttributes attributes = File.GetAttributes(launchPath);
                File.SetAttributes(launchPath, attributes | FileAttributes.ReadOnly);

                protectionStream = new FileStream(
                    launchPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    CopyBufferSize,
                    FileOptions.RandomAccess);
                ReviewWorkbookReadResult validation = ReviewWorkbookFile.ReadAndValidate(protectionStream);

                var result = new ReviewWorkbookLaunchCopy(
                    source,
                    launchDirectory,
                    launchPath,
                    byteLength,
                    packageSha256,
                    validation,
                    protectionStream);
                result.VerifyUnchanged();
                protectionStream = null;
                return result;
            }
            catch
            {
                if (protectionStream != null) protectionStream.Dispose();
                TryDeleteFile(launchPath);
                TryDeleteDirectory(launchDirectory);
                throw;
            }
        }

        /// <summary>
        /// Revalidates the exact retained file handle and its whole-package digest. Call this immediately after
        /// Excel returns from Workbooks.Open to verify that the protected launch identity remained unchanged.
        /// </summary>
        public ReviewWorkbookReadResult VerifyUnchanged()
        {
            ThrowIfDisposed();
            ReviewWorkbookReadResult current = ReviewWorkbookFile.ReadAndValidate(_protectionStream);
            string currentPackageSha256 = ComputeSha256(_protectionStream);
            if (!string.Equals(currentPackageSha256, PackageSha256, StringComparison.OrdinalIgnoreCase) ||
                !HasSameValidatedIdentity(Validation, current))
                throw new InvalidDataException(
                    "The protected review launch copy changed after validation.");
            return current;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            FileStream stream = _protectionStream;
            _protectionStream = null;
            if (stream != null) stream.Dispose();
            TryDeleteFile(LaunchPath);
            TryDeleteDirectory(LaunchDirectory);
        }

        internal static bool HasSameValidatedIdentity(
            ReviewWorkbookReadResult expected,
            ReviewWorkbookReadResult actual)
        {
            if (expected == null || actual == null) return false;
            return string.Equals(expected.ProjectId, actual.ProjectId, StringComparison.Ordinal) &&
                   string.Equals(expected.ManifestItemId, actual.ManifestItemId, StringComparison.Ordinal) &&
                   string.Equals(expected.StateSha256, actual.StateSha256, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(expected.ProjectionSha256, actual.ProjectionSha256, StringComparison.OrdinalIgnoreCase) &&
                   expected.ChunkItemIds.SequenceEqual(actual.ChunkItemIds, StringComparer.Ordinal);
        }

        private static string BuildLaunchFileName(string sourcePath, string token)
        {
            string stem = Path.GetFileNameWithoutExtension(sourcePath);
            if (string.IsNullOrWhiteSpace(stem)) stem = "Review";
            if (stem.Length > 80) stem = stem.Substring(0, 80);
            return stem + "-validated-" + token + ".xlsx";
        }

        private static void RejectReparsePoint(string path, string message)
        {
            if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
                throw new IOException(message);
        }

        private static string ComputeSha256(Stream stream)
        {
            long originalPosition = stream.Position;
            try
            {
                stream.Position = 0;
                using (SHA256 sha256 = SHA256.Create()) return BytesToHex(sha256.ComputeHash(stream));
            }
            finally
            {
                try { stream.Position = originalPosition; }
                catch { }
            }
        }

        private static string BytesToHex(byte[] bytes)
        {
            if (bytes == null) throw new InvalidDataException("A SHA-256 digest was not produced.");
            var chars = new char[bytes.Length * 2];
            const string alphabet = "0123456789abcdef";
            for (int index = 0; index < bytes.Length; index++)
            {
                chars[index * 2] = alphabet[bytes[index] >> 4];
                chars[index * 2 + 1] = alphabet[bytes[index] & 0x0F];
            }
            return new string(chars);
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (!File.Exists(path)) return;
                File.SetAttributes(path, FileAttributes.Normal);
                File.Delete(path);
            }
            catch
            {
                // The path is a random private launch artifact and cannot be mistaken for the source. A live
                // Excel read handle or process termination may defer deletion until a later maintenance pass.
            }
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path)) Directory.Delete(path, false);
            }
            catch { }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ReviewWorkbookLaunchCopy));
        }
    }
}
