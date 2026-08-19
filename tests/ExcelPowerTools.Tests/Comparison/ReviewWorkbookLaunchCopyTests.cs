using System;
using System.IO;
using ExcelPowerTools.Core.Comparison.Persistence;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewWorkbookLaunchCopyTests : IDisposable
    {
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.ReviewLaunch.Tests",
            Guid.NewGuid().ToString("N"));

        public ReviewWorkbookLaunchCopyTests()
        {
            Directory.CreateDirectory(_directory);
        }

        [Fact]
        public void Create_retains_an_exact_validated_read_only_copy_until_dispose()
        {
            string source = CreateReview("project-protected", "{\"run\":1}");
            byte[] expectedBytes = File.ReadAllBytes(source);
            string root = Path.Combine(_directory, "launch-root");
            string launchPath;
            string launchDirectory;

            using (ReviewWorkbookLaunchCopy launch = ReviewWorkbookLaunchCopy.Create(source, root))
            {
                launchPath = launch.LaunchPath;
                launchDirectory = launch.LaunchDirectory;

                Assert.Equal(Path.GetFullPath(source), launch.SourcePath);
                Assert.NotEqual(Path.GetFullPath(source), launchPath);
                Assert.StartsWith(Path.GetFullPath(root), launchPath, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(expectedBytes.LongLength, launch.ByteLength);
                Assert.Equal(64, launch.PackageSha256.Length);
                Assert.Equal(expectedBytes, File.ReadAllBytes(launchPath));
                Assert.True((File.GetAttributes(launchPath) & FileAttributes.ReadOnly) != 0);
                Assert.Equal("project-protected", launch.Validation.ProjectId);
                Assert.Equal("project-protected", launch.VerifyUnchanged().ProjectId);

                // The retained lease shares reads only: ordinary writers and delete/rename handles cannot
                // acquire the private launch file before Excel releases it.
                Exception denied = Record.Exception(() => new FileStream(
                    launchPath,
                    FileMode.Open,
                    FileAccess.Write,
                    FileShare.ReadWrite | FileShare.Delete));
                Assert.True(
                    denied is IOException || denied is UnauthorizedAccessException,
                    "The protected read-only launch unexpectedly admitted a writer.");
            }

            Assert.False(File.Exists(launchPath));
            Assert.False(Directory.Exists(launchDirectory));
            Assert.True(Directory.Exists(root));
            Assert.Equal(expectedBytes, File.ReadAllBytes(source));
        }

        [Fact]
        public void Protected_copy_remains_the_validated_snapshot_if_source_changes_later()
        {
            string source = CreateReview("project-snapshot", "{\"version\":1}");
            string root = Path.Combine(_directory, "launch-root");

            using (ReviewWorkbookLaunchCopy launch = ReviewWorkbookLaunchCopy.Create(source, root))
            {
                string expectedDigest = launch.PackageSha256;
                File.WriteAllBytes(source, new byte[] { 1, 2, 3, 4 });

                Assert.Equal(expectedDigest, launch.PackageSha256);
                Assert.Equal("project-snapshot", launch.VerifyUnchanged().ProjectId);
                Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(source));
            }
        }

        [Fact]
        public void Invalid_source_fails_closed_and_removes_private_launch_artifacts()
        {
            string source = Path.Combine(_directory, "invalid.xlsx");
            File.WriteAllBytes(source, new byte[] { 0x50, 0x4b, 0x03, 0x04, 1, 2, 3 });
            string root = Path.Combine(_directory, "launch-root");

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookLaunchCopy.Create(source, root));

            Assert.True(Directory.Exists(root));
            Assert.Empty(Directory.GetFiles(root));
            Assert.Empty(Directory.GetDirectories(root));
            Assert.Equal(new byte[] { 0x50, 0x4b, 0x03, 0x04, 1, 2, 3 }, File.ReadAllBytes(source));
        }

        [Fact]
        public void Stream_validation_uses_the_whole_package_and_restores_caller_position()
        {
            string source = CreateReview("project-stream", "{\"stream\":true}");
            byte[] package = File.ReadAllBytes(source);
            using var stream = new MemoryStream(package, writable: false) { Position = 7 };

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(stream);

            Assert.Equal("project-stream", read.ProjectId);
            Assert.Equal(7, stream.Position);
            Assert.True(stream.CanRead);
        }

        public void Dispose()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        private string CreateReview(string projectId, string stateJson)
        {
            string path = Path.Combine(_directory, projectId + ".xlsx");
            ReviewWorkbookFile.CreateNew(path, new ReviewWorkbookState(projectId, stateJson));
            return path;
        }
    }
}
