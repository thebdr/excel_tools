using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelPowerTools.Core.Analysis
{
    /// <summary>Per-column profile: counts, numeric range, lengths, and the most frequent values.</summary>
    public sealed class ColumnStatsResult
    {
        public int Total;
        public int NonEmpty;
        public int Empty;
        public int Distinct;
        public int NumericCount;
        public double? Min;
        public double? Max;
        public int MinLength;        // over non-empty values
        public int MaxLength;
        public List<KeyValuePair<string, int>> TopValues = new List<KeyValuePair<string, int>>();
    }

    public static class ColumnStats
    {
        /// <summary>Profiles a column of cell texts (invariant "." decimals count as numeric).</summary>
        public static ColumnStatsResult Compute(IEnumerable<string> values, int topN = 10)
        {
            var r = new ColumnStatsResult { MinLength = int.MaxValue };
            var counts = new Dictionary<string, int>();

            foreach (string raw in values)
            {
                string v = raw ?? "";
                r.Total++;

                if (v.Length == 0)
                {
                    r.Empty++;
                    continue;
                }

                r.NonEmpty++;
                if (v.Length < r.MinLength) r.MinLength = v.Length;
                if (v.Length > r.MaxLength) r.MaxLength = v.Length;

                counts.TryGetValue(v, out int c);
                counts[v] = c + 1;

                if (double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                {
                    r.NumericCount++;
                    if (r.Min == null || d < r.Min) r.Min = d;
                    if (r.Max == null || d > r.Max) r.Max = d;
                }
            }

            if (r.NonEmpty == 0) r.MinLength = 0;
            r.Distinct = counts.Count;
            r.TopValues = counts
                .OrderByDescending(kv => kv.Value)
                .ThenBy(kv => kv.Key, System.StringComparer.Ordinal)
                .Take(topN)
                .ToList();
            return r;
        }
    }
}
