using System;
using System.IO;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Files;
using Xunit;

namespace ExcelPowerTools.Tests.Files
{
    public class AtomicFileStoreTests : IDisposable
    {
        private readonly string _dir;

        public AtomicFileStoreTests()
        {
            _dir = Path.Combine(Path.GetTempPath(), "pt-afs-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, recursive: true); } catch { }
        }

        private string PathOf(string name) => Path.Combine(_dir, name);

        [Fact]
        public void Write_creates_then_replaces_with_bak()
        {
            string path = PathOf("data.csv");
            byte[] first = Encoding.UTF8.GetBytes("one");
            byte[] second = Encoding.UTF8.GetBytes("two");

            AtomicFileStore.WriteWithBackup(path, first);
            Assert.Equal(first, File.ReadAllBytes(path));
            Assert.False(File.Exists(path + ".bak"));

            AtomicFileStore.WriteWithBackup(path, second);
            Assert.Equal(second, File.ReadAllBytes(path));
            Assert.Equal(first, File.ReadAllBytes(path + ".bak"));   // one-level backup holds the prior bytes

            Assert.Empty(Directory.GetFiles(_dir, "*.tmp"));         // never a stray temp file
        }

        [Fact]
        public void Write_replaces_a_readonly_bak()
        {
            string path = PathOf("data.csv");
            AtomicFileStore.WriteWithBackup(path, new byte[] { 1 });
            AtomicFileStore.WriteWithBackup(path, new byte[] { 2 });
            File.SetAttributes(path + ".bak", FileAttributes.ReadOnly);

            AtomicFileStore.WriteWithBackup(path, new byte[] { 3 });

            Assert.Equal(new byte[] { 3 }, File.ReadAllBytes(path));
            Assert.Equal(new byte[] { 2 }, File.ReadAllBytes(path + ".bak"));
        }

        [Fact]
        public void ReadSource_classifies_present_missing_unavailable()
        {
            string path = PathOf("data.csv");

            Assert.Equal(SourceReadState.Missing, AtomicFileStore.ReadSource(path, out _, out _));
            Assert.Equal(SourceReadState.Missing,
                AtomicFileStore.ReadSource(PathOf(@"nodir\data.csv"), out _, out _));

            File.WriteAllBytes(path, new byte[] { 9 });
            Assert.Equal(SourceReadState.Present, AtomicFileStore.ReadSource(path, out byte[] bytes, out _));
            Assert.Equal(new byte[] { 9 }, bytes);

            using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                Assert.Equal(SourceReadState.Unavailable, AtomicFileStore.ReadSource(path, out _, out string error));
                Assert.False(string.IsNullOrEmpty(error));
            }
        }

        [Fact]
        public void DiskStillMatches_tracks_content_and_existence()
        {
            string path = PathOf("data.csv");
            byte[] bytes = Encoding.UTF8.GetBytes("payload");

            Assert.True(AtomicFileStore.DiskStillMatches(path, expectedExists: false, expectedBytes: null));
            Assert.False(AtomicFileStore.DiskStillMatches(path, expectedExists: true, expectedBytes: bytes));

            File.WriteAllBytes(path, bytes);
            Assert.True(AtomicFileStore.DiskStillMatches(path, expectedExists: true, expectedBytes: bytes));
            Assert.False(AtomicFileStore.DiskStillMatches(path, expectedExists: false, expectedBytes: null));

            File.WriteAllBytes(path, Encoding.UTF8.GetBytes("changed"));
            Assert.False(AtomicFileStore.DiskStillMatches(path, expectedExists: true, expectedBytes: bytes));
        }

        [Theory]
        [InlineData(new byte[] { 1, 2 }, new byte[] { 1, 2 }, true)]
        [InlineData(new byte[] { 1, 2 }, new byte[] { 1, 3 }, false)]
        [InlineData(new byte[] { 1 }, new byte[] { 1, 2 }, false)]
        [InlineData(new byte[0], new byte[0], true)]
        public void BytesEqual_compares_content(byte[] a, byte[] b, bool expected)
        {
            Assert.Equal(expected, AtomicFileStore.BytesEqual(a, b));
        }

        [Fact]
        public void BytesEqual_treats_null_as_unequal()
        {
            Assert.False(AtomicFileStore.BytesEqual(null, new byte[0]));
            Assert.False(AtomicFileStore.BytesEqual(new byte[0], null));
            Assert.False(AtomicFileStore.BytesEqual(null, null));
        }

        [Fact]
        public void Concurrent_reader_does_not_break_the_replace()
        {
            // A reader holding FileShare-friendly access (the OneDrive/scanner case File.Replace
            // tolerates) must not defeat the bounded retry.
            string path = PathOf("data.csv");
            AtomicFileStore.WriteWithBackup(path, new byte[] { 1 });

            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                AtomicFileStore.WriteWithBackup(path, new byte[] { 2 });
            }
            Assert.Equal(new byte[] { 2 }, File.ReadAllBytes(path));
        }
    }
}
