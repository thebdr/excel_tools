using System;
using ExcelPowerTools.Core.Comparison.Canonical;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class CanonicalValueTests
    {
        [Fact]
        public void Blank_and_empty_text_remain_distinct()
        {
            Assert.NotEqual(CanonicalValue.Blank, CanonicalValue.Text(""));
            Assert.Equal(CanonicalValueKind.Blank, CanonicalValue.Blank.Kind);
            Assert.Equal(CanonicalValueKind.Text, CanonicalValue.Text("").Kind);
        }

        [Fact]
        public void Numbers_use_invariant_round_trip_identity()
        {
            Assert.Equal(CanonicalValue.Number(1), CanonicalValue.Number(1.0));
            Assert.Equal("1.25", CanonicalValue.Number(1.25).ExactText);
        }

        [Fact]
        public void Date_system_is_part_of_exact_identity()
        {
            CanonicalValue date1900 = CanonicalValue.DateSerial(45000, ExcelDateSystem.Date1900);
            CanonicalValue date1904 = CanonicalValue.DateSerial(45000, ExcelDateSystem.Date1904);

            Assert.NotEqual(date1900, date1904);
        }

        [Fact]
        public void Iso_date_time_is_not_tied_to_an_excel_serial_date_system()
        {
            CanonicalValue value = CanonicalValue.IsoDateTime("2026-07-16T14:15:16.123");

            Assert.Equal(CanonicalValueKind.IsoDateTime, value.Kind);
            Assert.Equal(ExcelDateSystem.NotApplicable, value.DateSystem);
            Assert.Equal("2026-07-16T14:15:16.123", value.ExactText);
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        public void Non_finite_numbers_are_rejected(double value)
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => CanonicalValue.Number(value));
        }

        [Fact]
        public void Snapshot_rejects_duplicate_logical_subjects()
        {
            var cells = new[]
            {
                CanonicalCellState.Constant("sheet/row/column", CanonicalValue.Text("a")),
                CanonicalCellState.Constant("sheet/row/column", CanonicalValue.Text("b")),
            };

            Assert.Throws<ArgumentException>(() => new RevisionSnapshot("revision", cells));
        }
    }
}
