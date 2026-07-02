using ExcelPowerTools.Core.Analysis;
using Xunit;

namespace ExcelPowerTools.Tests.Analysis
{
    public class ColumnStatsTests
    {
        [Fact]
        public void Profiles_a_mixed_column()
        {
            var r = ColumnStats.Compute(new[] { "A", "A", "B", "", "1.5", "2", null, "007" });

            Assert.Equal(8, r.Total);
            Assert.Equal(6, r.NonEmpty);
            Assert.Equal(2, r.Empty);
            Assert.Equal(5, r.Distinct);        // A, B, 1.5, 2, 007
            Assert.Equal(3, r.NumericCount);    // 1.5, 2, 007
            Assert.Equal(1.5, r.Min);
            Assert.Equal(7.0, r.Max);           // "007" parses as 7
            Assert.Equal(1, r.MinLength);
            Assert.Equal(3, r.MaxLength);

            Assert.Equal("A", r.TopValues[0].Key);
            Assert.Equal(2, r.TopValues[0].Value);
        }

        [Fact]
        public void Empty_column_is_safe()
        {
            var r = ColumnStats.Compute(new string[] { "", null });
            Assert.Equal(2, r.Total);
            Assert.Equal(0, r.NonEmpty);
            Assert.Equal(0, r.Distinct);
            Assert.Null(r.Min);
            Assert.Equal(0, r.MinLength);
            Assert.Empty(r.TopValues);
        }

        [Fact]
        public void Top_values_capped_and_ordered()
        {
            var values = new System.Collections.Generic.List<string>();
            for (int v = 0; v < 20; v++)
                for (int n = 0; n <= v; n++)
                    values.Add("V" + v);

            var r = ColumnStats.Compute(values, topN: 5);
            Assert.Equal(5, r.TopValues.Count);
            Assert.Equal("V19", r.TopValues[0].Key);   // most frequent first
            Assert.Equal(20, r.TopValues[0].Value);
        }
    }
}
