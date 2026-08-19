using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Stores coarse comparison timings under the current user's temporary folder. Bucket names are
    /// hashed before they become filenames and the file content contains only a sample count and a
    /// cumulative duration. Source paths, filenames, workbook labels, and individual run records are
    /// never persisted.
    /// </summary>
    public sealed class ComparisonElapsedTimeEstimateStore
    {
        private static readonly object FileGate = new object();
        private readonly string _storageDirectory;

        public ComparisonElapsedTimeEstimateStore(string storageDirectory = null)
        {
            _storageDirectory = string.IsNullOrWhiteSpace(storageDirectory)
                ? Path.Combine(Path.GetTempPath(), "ExcelPowerTools", "ComparisonTiming", "v1")
                : Path.GetFullPath(storageDirectory);
        }

        public static ComparisonElapsedTimeEstimateStore Default { get; } =
            new ComparisonElapsedTimeEstimateStore();

        /// <summary>
        /// Reads the arithmetic mean of previous successful runs for a non-sensitive workload bucket.
        /// Corrupt, missing, or inaccessible timing data is treated as no estimate.
        /// </summary>
        public bool TryGetEstimate(string workloadBucket, out TimeSpan estimate, out int sampleCount)
        {
            estimate = TimeSpan.Zero;
            sampleCount = 0;
            if (string.IsNullOrWhiteSpace(workloadBucket)) return false;

            try
            {
                lock (FileGate)
                {
                    Aggregate aggregate;
                    if (!TryRead(GetBucketPath(workloadBucket), out aggregate)) return false;
                    double averageMilliseconds = aggregate.TotalMilliseconds / (double)aggregate.SampleCount;
                    if (double.IsNaN(averageMilliseconds) || double.IsInfinity(averageMilliseconds) ||
                        averageMilliseconds < 0 || averageMilliseconds > TimeSpan.MaxValue.TotalMilliseconds)
                        return false;

                    estimate = TimeSpan.FromMilliseconds(averageMilliseconds);
                    sampleCount = aggregate.SampleCount;
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Adds a successful run to a bucket aggregate. This is deliberately best effort: inability to
        /// update a disposable timing estimate must never fail an otherwise successful comparison.
        /// </summary>
        public bool TryRecordSuccessfulRun(string workloadBucket, TimeSpan elapsed)
        {
            if (string.IsNullOrWhiteSpace(workloadBucket) || elapsed < TimeSpan.Zero) return false;

            try
            {
                lock (FileGate)
                {
                    Directory.CreateDirectory(_storageDirectory);
                    string path = GetBucketPath(workloadBucket);
                    Aggregate aggregate;
                    if (!TryRead(path, out aggregate)) aggregate = new Aggregate(0, 0);

                    long elapsedMilliseconds = Math.Max(0L, (long)Math.Round(elapsed.TotalMilliseconds));
                    long remaining = long.MaxValue - aggregate.TotalMilliseconds;
                    if (elapsedMilliseconds > remaining) return false;
                    if (aggregate.SampleCount == int.MaxValue) return false;

                    var updated = new Aggregate(
                        aggregate.SampleCount + 1,
                        aggregate.TotalMilliseconds + elapsedMilliseconds);
                    WriteAtomically(path, updated);
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private string GetBucketPath(string workloadBucket)
        {
            byte[] input = Encoding.UTF8.GetBytes(workloadBucket.Trim());
            byte[] digest;
            using (SHA256 algorithm = SHA256.Create()) digest = algorithm.ComputeHash(input);

            var name = new StringBuilder(digest.Length * 2);
            for (int index = 0; index < digest.Length; index++)
                name.Append(digest[index].ToString("x2", CultureInfo.InvariantCulture));
            return Path.Combine(_storageDirectory, name + ".timing");
        }

        private static bool TryRead(string path, out Aggregate aggregate)
        {
            aggregate = default(Aggregate);
            if (!File.Exists(path)) return false;

            string content = File.ReadAllText(path, Encoding.ASCII).Trim();
            string[] parts = content.Split('|');
            int count;
            long totalMilliseconds;
            if (parts.Length != 2 ||
                !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out count) ||
                !long.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out totalMilliseconds) ||
                count <= 0 || totalMilliseconds < 0)
                return false;

            aggregate = new Aggregate(count, totalMilliseconds);
            return true;
        }

        private static void WriteAtomically(string destination, Aggregate aggregate)
        {
            string temporary = destination + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                string content = aggregate.SampleCount.ToString(CultureInfo.InvariantCulture) + "|" +
                                 aggregate.TotalMilliseconds.ToString(CultureInfo.InvariantCulture);
                File.WriteAllText(temporary, content, Encoding.ASCII);

                if (File.Exists(destination))
                {
                    try
                    {
                        File.Replace(temporary, destination, null, true);
                        temporary = null;
                    }
                    catch (PlatformNotSupportedException)
                    {
                        File.Copy(temporary, destination, true);
                    }
                    catch (IOException)
                    {
                        File.Copy(temporary, destination, true);
                    }
                }
                else
                {
                    File.Move(temporary, destination);
                    temporary = null;
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(temporary))
                {
                    try { File.Delete(temporary); }
                    catch { }
                }
            }
        }

        private struct Aggregate
        {
            internal Aggregate(int sampleCount, long totalMilliseconds)
            {
                SampleCount = sampleCount;
                TotalMilliseconds = totalMilliseconds;
            }

            internal int SampleCount { get; }
            internal long TotalMilliseconds { get; }
        }
    }

    /// <summary>
    /// Produces coarse, non-sensitive timing buckets. Exact workbook sizes and counts are intentionally
    /// not encoded, so callers can obtain useful estimates without using paths or document names as keys.
    /// </summary>
    public static class ComparisonWorkloadBucket
    {
        public static string Create(
            int revisionCount,
            long totalBytes,
            int totalSheetCount,
            long approximateUsedCellCount = 0)
        {
            if (revisionCount < 0) throw new ArgumentOutOfRangeException(nameof(revisionCount));
            if (totalBytes < 0) throw new ArgumentOutOfRangeException(nameof(totalBytes));
            if (totalSheetCount < 0) throw new ArgumentOutOfRangeException(nameof(totalSheetCount));
            if (approximateUsedCellCount < 0) throw new ArgumentOutOfRangeException(nameof(approximateUsedCellCount));

            return "comparison-v1" +
                   "|revisions:" + Band(revisionCount, 2, 5, 10, 20) +
                   "|bytes:" + Band(totalBytes, 10L << 20, 50L << 20, 200L << 20, 500L << 20) +
                   "|sheets:" + Band(totalSheetCount, 5, 15, 30, 75) +
                   "|cells:" + Band(approximateUsedCellCount, 150000L, 1000000L, 5000000L, 20000000L);
        }

        private static string Band(long value, long first, long second, long third, long fourth)
        {
            if (value <= first) return "xs";
            if (value <= second) return "s";
            if (value <= third) return "m";
            if (value <= fourth) return "l";
            return "xl";
        }
    }
}
