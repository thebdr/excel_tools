using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The first-line and group numbers, read off the SWP's own tabs. The fixtures are the real
    /// FVZ induction workbooks: G1 carries IL01_* plus ILGRP01_*, G2 carries IL02_*, IL03_* and
    /// ILGRP02_CABINET.
    /// </summary>
    public class SwpPlantNumberingTests
    {
        private static readonly string[] InductionG1 =
        {
            "Cover", "IL01_ZONE", "IL01_CABINET", "IL01_BELT", "IL01_MODULATOR", "IL01_COIL",
            "ILGRP01_SENSOR", "ILGRP01_CABINET",
        };

        private static readonly string[] InductionG2 =
        {
            "Cover", "IL02_ZONE", "IL02_CABINET", "IL02_BELT", "IL02_MODULATOR", "IL02_COIL",
            "IL03_ZONE", "IL03_CABINET", "IL03_BELT", "IL03_MODULATOR", "IL03_COIL",
            "ILGRP02_CABINET",
        };

        [Fact]
        public void The_first_line_is_the_lowest_numbered_sheet()
        {
            SwpPlantNumbering.Result g1 = SwpPlantNumbering.FromSheetNames(InductionG1, SwpPlantType.Ispc);
            Assert.True(g1.HasFirstLine);
            Assert.Equal(1, g1.FirstLineId);
            Assert.True(g1.HasGroup);
            Assert.Equal(1, g1.GroupIndex);

            // G2 holds TWO lines: the map computes firstLineId + line - 1, so the first is the
            // lowest — reading the last sheet instead would send every row one sheet too far.
            SwpPlantNumbering.Result g2 = SwpPlantNumbering.FromSheetNames(InductionG2, SwpPlantType.Ispc);
            Assert.Equal(2, g2.FirstLineId);
            Assert.Equal(2, g2.GroupIndex);
        }

        [Fact]
        public void The_group_sheet_is_not_read_as_a_line()
        {
            // "ILGRP02_CABINET" also starts with "IL": reading its "GR" as a line number is the
            // silent misparse this replaces. A file with ONLY a group sheet yields no line.
            SwpPlantNumbering.Result result =
                SwpPlantNumbering.FromSheetNames(new[] { "Cover", "ILGRP02_CABINET" }, SwpPlantType.Ispc);

            Assert.False(result.HasFirstLine);
            Assert.True(result.HasGroup);
            Assert.Equal(2, result.GroupIndex);
            Assert.Equal("ILGRP02_CABINET", result.Evidence);
        }

        [Fact]
        public void Upstream_reads_its_system_id_and_leaves_the_conveyor_group_alone()
        {
            // UP tabs carry the system id; the conveyor group is a PRODUCTION-tree fact
            // (ConveyorSys\SysGroupNN), so nothing in the SWP can supply it.
            SwpPlantNumbering.Result result = SwpPlantNumbering.FromSheetNames(
                new[] { "Cover", "UP04_CABINET", "UP05_BELT", "UP06_ZONE" }, SwpPlantType.Ups);

            Assert.True(result.HasFirstLine);
            Assert.Equal(4, result.FirstLineId);
            Assert.False(result.HasGroup);
        }

        [Theory]
        [InlineData(SwpPlantType.Msc)]
        [InlineData(SwpPlantType.Safety)]
        public void Plants_without_numbered_sheets_derive_nothing(SwpPlantType plantType)
        {
            // One cabinet table, no per-line sheets: the dialog's own values must stand.
            SwpPlantNumbering.Result result = SwpPlantNumbering.FromSheetNames(
                new[] { "Cover", "CABINET", "CHUTE", "IL01_BELT" }, plantType);

            Assert.False(result.Any);
            Assert.Equal("", result.Evidence);
        }

        [Fact]
        public void Unnumbered_and_missing_input_derive_nothing()
        {
            Assert.False(SwpPlantNumbering.FromSheetNames(
                new[] { "Cover", "Notes", "ILBELT" }, SwpPlantType.Ispc).Any);
            Assert.False(SwpPlantNumbering.FromSheetNames(new string[0], SwpPlantType.Ispc).Any);
            Assert.False(SwpPlantNumbering.FromSheetNames(null, SwpPlantType.Ispc).Any);
            Assert.False(SwpPlantNumbering.FromSheetNames(
                new[] { "Cover", null, "", "IL07_BELT" }, SwpPlantType.Ispc).HasGroup);
        }

        [Fact]
        public void Sheet_names_are_matched_case_insensitively()
        {
            SwpPlantNumbering.Result result = SwpPlantNumbering.FromSheetNames(
                new[] { "il09_belt", "ilgrp03_cabinet" }, SwpPlantType.Ispc);

            Assert.Equal(9, result.FirstLineId);
            Assert.Equal(3, result.GroupIndex);
        }

        [Fact]
        public void The_evidence_names_the_sheets_the_numbers_came_from()
        {
            // The dialog shows this instead of a number appearing on its own.
            Assert.Equal("IL02_ZONE, ILGRP02_CABINET",
                SwpPlantNumbering.FromSheetNames(InductionG2, SwpPlantType.Ispc).Evidence);
        }
    }
}
