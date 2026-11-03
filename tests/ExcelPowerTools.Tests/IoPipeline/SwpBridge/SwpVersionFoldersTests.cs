using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The dated version folders: which one a session comes back to, and what a new one takes
    /// along. The file list is the real induction folder — two SWPs one IO list drives, their
    /// CONFIG_PLANT, and the debris a working folder collects.
    /// </summary>
    public class SwpVersionFoldersTests
    {
        private const string G2 = "SWP_02_Induction_generated_G2.xlsx";
        private const string Config = "CONFIG_PLANT_02_Induction.xlsx";

        private static readonly string[] InductionFolder =
        {
            "CONFIG_PLANT_02_Induction.xlsx",
            "SWP_02_Induction_generated_G1.xlsx",
            "SWP_02_Induction_generated_G1.xlsx.bak",
            "SWP_02_Induction_generated_G2.xlsx",
            "SWP_02_Induction_generated_G2.xlsx.bak",
            "~$SWP_02_Induction_generated_G1.xlsx",
            "SwpBridge_Log.txt",
        };

        [Theory]
        [InlineData("06_20-Aug-2026", 6)]
        [InlineData("01_23-Giu-2026", 1)]
        [InlineData("10_10-set-2026", 10)]
        public void A_dated_folder_reads_its_number(string name, int expected) =>
            Assert.Equal((int?)expected, SwpVersionFolders.Number(name));

        [Theory]
        [InlineData("archive")]
        [InlineData("_23-Giu-2026")]
        [InlineData("v2_23-Giu-2026")]
        [InlineData("")]
        [InlineData(null)]
        public void Any_other_folder_has_no_number(string name) =>
            Assert.Null(SwpVersionFolders.Number(name));

        [Fact]
        public void The_newest_folder_is_the_highest_number_not_the_last_name()
        {
            Assert.Equal("10_10-set-2026", SwpVersionFolders.Newest(new[] { "9_01-set-2026", "10_10-set-2026", "archive" }));
            Assert.Equal("02_10-set-2026", SwpVersionFolders.Newest(new[] { "02_10-set-2026", "01_23-Giu-2026" }));
            Assert.Null(SwpVersionFolders.Newest(new[] { "archive", "Sessions" }));
            Assert.Null(SwpVersionFolders.Newest(null));
        }

        [Fact]
        public void The_other_swp_an_io_list_drives_is_taken_along()
        {
            List<string> copy = SwpVersionFolders.FilesToCopy(InductionFolder, @"^SWP_.*\.xlsx$", G2, Config, out string error);

            // G2 is the SWP being applied — written patched, never copied over itself.
            Assert.Equal(new[] { "SWP_02_Induction_generated_G1.xlsx" }, copy);
            Assert.Equal("", error);
        }

        [Fact]
        public void A_loose_pattern_still_leaves_safety_copies_and_lock_files_behind()
        {
            // "g1" also matches G1's .bak and Excel's ~$ lock file — neither belongs in a new version.
            List<string> copy = SwpVersionFolders.FilesToCopy(InductionFolder, "g1", G2, Config, out _);

            Assert.Equal(new[] { "SWP_02_Induction_generated_G1.xlsx" }, copy);
        }

        [Fact]
        public void The_swp_and_config_plant_are_never_listed_even_when_matched()
        {
            List<string> copy = SwpVersionFolders.FilesToCopy(InductionFolder, @"\.xlsx$", G2, Config, out _);

            Assert.Equal(new[] { "SWP_02_Induction_generated_G1.xlsx" }, copy);
        }

        [Fact]
        public void Files_are_listed_in_folder_order()
        {
            List<string> copy = SwpVersionFolders.FilesToCopy(InductionFolder, @"G1\.xlsx$|\.txt$", G2, Config, out _);

            Assert.Equal(new[] { "SWP_02_Induction_generated_G1.xlsx", "SwpBridge_Log.txt" }, copy);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void An_empty_pattern_takes_nothing(string pattern)
        {
            Assert.Empty(SwpVersionFolders.FilesToCopy(InductionFolder, pattern, G2, Config, out string error));
            Assert.Equal("", error);
        }

        [Fact]
        public void A_pattern_that_does_not_parse_takes_nothing_and_says_why()
        {
            Assert.Empty(SwpVersionFolders.FilesToCopy(InductionFolder, "([unclosed", G2, Config, out string error));
            Assert.NotEqual("", error);
        }
    }
}
