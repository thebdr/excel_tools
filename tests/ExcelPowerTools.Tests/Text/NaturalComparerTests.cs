using System;
using System.Linq;
using ExcelPowerTools.Core.Text;
using Xunit;

namespace ExcelPowerTools.Tests.Text
{
    public class NaturalComparerTests
    {
        private static int Cmp(string a, string b) => NaturalComparer.Instance.Compare(a, b);

        [Theory]
        [InlineData("DI2", "DI10")]
        [InlineData("Slot_2", "Slot_10")]
        [InlineData("Z9", "Z100")]
        [InlineData("AUTO1_Z01", "AUTO1_Z02")]
        [InlineData("A", "B")]
        [InlineData("", "A")]
        [InlineData("192.168.0.2", "192.168.0.10")]
        public void Orders_naturally(string smaller, string bigger)
        {
            Assert.True(Cmp(smaller, bigger) < 0);
            Assert.True(Cmp(bigger, smaller) > 0);
        }

        [Fact]
        public void Case_insensitive_and_equal()
        {
            Assert.Equal(0, Cmp("abc", "ABC"));
            Assert.Equal(0, Cmp("Z01", "Z01"));
        }

        [Fact]
        public void Numbers_sort_before_text()
        {
            Assert.True(Cmp("5", "a") < 0);
        }

        [Fact]
        public void Padded_equal_numbers_are_deterministic()
        {
            // 07 and 7 are numerically equal; padding breaks the tie consistently (never 0).
            Assert.True(Cmp("Z07", "Z7") != 0);
            Assert.Equal(-Math.Sign(Cmp("Z7", "Z07")), Math.Sign(Cmp("Z07", "Z7")));
        }

        [Fact]
        public void Sorting_a_realistic_column_is_stable_and_natural()
        {
            var input = new[] { "DI10", "DI2", "DI1", "DI21", "DI3" };
            var sorted = input.OrderBy(x => x, NaturalComparer.Instance).ToArray();
            Assert.Equal(new[] { "DI1", "DI2", "DI3", "DI10", "DI21" }, sorted);
        }
    }
}
