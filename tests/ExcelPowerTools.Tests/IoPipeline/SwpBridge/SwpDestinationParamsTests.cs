using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpDestinationParamsTests
    {
        [Fact]
        public void Ordinal_counts_only_active_lines()
        {
            var log = new List<SwpLogEntry>();
            var map = SwpDestinationParams.ParseChuteMap(new[]
            {
                "#;header;line",
                "@;101;CHUTE A",
                "!;include;directive",
                "@;102;CHUTE B",
                "",
                "@;205;CHUTE C",
            }, log);

            Assert.Equal(0, map["101"]);
            Assert.Equal(1, map["102"]);
            Assert.Equal(2, map["205"]);
            Assert.Empty(log);
        }

        [Fact]
        public void Comma_separator_and_numeric_normalization_work()
        {
            var map = SwpDestinationParams.ParseChuteMap(new[] { "@,007,X" }, new List<SwpLogEntry>());
            Assert.Equal(0, map["7"]);
        }

        [Fact]
        public void First_occurrence_wins_on_duplicate_indexes()
        {
            var map = SwpDestinationParams.ParseChuteMap(new[] { "@;5;A", "@;5;B" }, new List<SwpLogEntry>());
            Assert.Equal(0, map["5"]);
        }

        [Fact]
        public void Invalid_separator_logs_but_the_ordinal_still_advances()
        {
            var log = new List<SwpLogEntry>();
            var map = SwpDestinationParams.ParseChuteMap(new[]
            {
                "@|1|BROKEN",
                "@;2;GOOD",
            }, log);

            Assert.Contains(log, l => l.Severity == SwpLogSeverity.Error && l.Message.Contains("separator"));
            Assert.Equal(1, map["2"]);
        }

        [Fact]
        public void Key_formats_the_msc_index_invariantly()
        {
            Assert.Equal("205", SwpDestinationParams.Key(205));
        }
    }
}
