using System;
using System.IO;
using System.Threading;

namespace ExcelPowerTools.Core.Files
{
    public enum SourceReadState { Present, Missing, Unavailable }

    /// <summary>
    /// Byte-level file primitives shared by every safe-save path: read-with-classification, atomic
    /// backup write, and disk-vs-memory comparison. Lifted from the CSV save pipeline so headless
    /// writers (the IO Bridge apply engine, sidecar persistence) use the identical semantics.
    /// </summary>
    public static class AtomicFileStore
    {
        // OneDrive and antivirus scanners take short-lived handles on freshly changed files; a
        // replace that collides with one is transient, so it is retried briefly before surfacing.
        private const int ReplaceAttempts = 3;
        private const int ReplaceRetryDelayMs = 250;

        public static SourceReadState ReadSource(string path, out byte[] bytes, out string error)
        {
            bytes = null;
            error = null;
            try
            {
                bytes = File.ReadAllBytes(path);
                return SourceReadState.Present;
            }
            catch (FileNotFoundException) { return SourceReadState.Missing; }
            catch (DirectoryNotFoundException) { return SourceReadState.Missing; }
            catch (Exception ex)
            {
                error = ex.Message;
                return SourceReadState.Unavailable;
            }
        }

        public static bool DiskStillMatches(string path, bool expectedExists, byte[] expectedBytes)
        {
            SourceReadState state = ReadSource(path, out byte[] actual, out _);
            if (state == SourceReadState.Unavailable) return false;
            bool exists = state == SourceReadState.Present;
            return exists == expectedExists && (!exists || BytesEqual(actual, expectedBytes));
        }

        public static bool BytesEqual(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            return true;
        }

        /// <summary>Writes bytes with a one-level ".bak" backup and an atomic replace (no partial file on crash).</summary>
        public static void WriteWithBackup(string path, byte[] bytes)
        {
            string directory = Path.GetDirectoryName(path);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
            string tmp = Path.Combine(directory, "." + Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllBytes(tmp, bytes);
                if (File.Exists(path))
                {
                    string bak = path + ".bak";
                    try { if (File.Exists(bak)) File.SetAttributes(bak, FileAttributes.Normal); } catch { }
                    WithTransientRetry(() => File.Replace(tmp, path, bak));   // atomic; original stays intact on failure
                }
                else
                {
                    WithTransientRetry(() => File.Move(tmp, path));
                }
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }   // never leave a stray .tmp
            }
        }

        private static void WithTransientRetry(Action operation)
        {
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    operation();
                    return;
                }
                catch (IOException) when (attempt < ReplaceAttempts)
                {
                    Thread.Sleep(ReplaceRetryDelayMs);
                }
            }
        }
    }
}
