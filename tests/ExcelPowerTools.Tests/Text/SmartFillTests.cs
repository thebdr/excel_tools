using ExcelPowerTools.Core.Text;
using Xunit;

namespace ExcelPowerTools.Tests.Text
{
    public class SmartFillTests
    {
        [Fact]
        public void Trailing_number_increments_with_padding_kept()
        {
            var seq = FillSequence.FromSeeds("AUTO1_Z07");
            Assert.False(seq.TwoSeedsConsumed);
            Assert.Equal("AUTO1_Z07", seq.ValueAt(0));
            Assert.Equal("AUTO1_Z08", seq.ValueAt(1));
            Assert.Equal("AUTO1_Z10", seq.ValueAt(3));   // 07 -> 10, width 2 kept
        }

        [Fact]
        public void Padding_grows_naturally_past_its_width()
        {
            var seq = FillSequence.FromSeeds("Z99");
            Assert.Equal("Z100", seq.ValueAt(1));
        }

        [Fact]
        public void Ip_last_octet_is_the_trailing_run()
        {
            var seq = FillSequence.FromSeeds("192.168.0.5");
            Assert.Equal("192.168.0.6", seq.ValueAt(1));
            Assert.Equal("192.168.0.15", seq.ValueAt(10));
        }

        [Fact]
        public void Two_seeds_infer_the_step()
        {
            var seq = FillSequence.FromSeeds("DB10", "DB20");
            Assert.True(seq.TwoSeedsConsumed);
            Assert.Equal("DB10", seq.ValueAt(0));
            Assert.Equal("DB20", seq.ValueAt(1));
            Assert.Equal("DB30", seq.ValueAt(2));
            Assert.Equal("DB40", seq.ValueAt(3));
        }

        [Fact]
        public void Second_seed_with_different_prefix_is_ignored()
        {
            var seq = FillSequence.FromSeeds("Z01", "Q05");
            Assert.False(seq.TwoSeedsConsumed);
            Assert.Equal("Z02", seq.ValueAt(1));   // step stays 1 from the first seed
        }

        [Fact]
        public void Negative_step_clamps_at_zero()
        {
            var seq = FillSequence.FromSeeds("N3", "N2");
            Assert.True(seq.TwoSeedsConsumed);
            Assert.Equal("N1", seq.ValueAt(2));
            Assert.Equal("N0", seq.ValueAt(3));
            Assert.Equal("N0", seq.ValueAt(4));   // clamped
        }

        [Fact]
        public void No_trailing_digits_means_plain_copy()
        {
            var seq = FillSequence.FromSeeds("HEADER");
            Assert.Equal("HEADER", seq.ValueAt(0));
            Assert.Equal("HEADER", seq.ValueAt(5));
        }

        [Fact]
        public void Pure_number_seed_works_with_empty_prefix()
        {
            var seq = FillSequence.FromSeeds("007");
            Assert.Equal("008", seq.ValueAt(1));
            Assert.Equal("012", seq.ValueAt(5));
        }

        [Fact]
        public void Digits_in_the_middle_do_not_count_as_trailing()
        {
            var seq = FillSequence.FromSeeds("Z01_backup");
            Assert.Equal("Z01_backup", seq.ValueAt(1));   // plain copy
        }

        [Theory]
        [InlineData("AUTO1_Z07", "AUTO1_Z", 7, 2)]
        [InlineData("007", "", 7, 3)]
        [InlineData("192.168.0.5", "192.168.0.", 5, 1)]
        public void TryParseTrailingNumber_splits_correctly(string input, string prefix, long value, int width)
        {
            Assert.True(FillSequence.TryParseTrailingNumber(input, out string p, out long v, out int w));
            Assert.Equal(prefix, p);
            Assert.Equal(value, v);
            Assert.Equal(width, w);
        }

        [Fact]
        public void TryParseTrailingNumber_rejects_no_digits_and_empty()
        {
            Assert.False(FillSequence.TryParseTrailingNumber("abc", out _, out _, out _));
            Assert.False(FillSequence.TryParseTrailingNumber("", out _, out _, out _));
            Assert.False(FillSequence.TryParseTrailingNumber(null, out _, out _, out _));
        }
    }
}
