using System;
using ExcelPowerTools.Core.Text;
using Xunit;

namespace ExcelPowerTools.Tests.Text
{
    public class RegexReplacerTests
    {
        private static bool Run(RegexReplaceSpec spec, string input, out string output)
        {
            var rx = RegexReplacer.Build(spec);
            return RegexReplacer.TryReplace(rx, RegexReplacer.EffectiveReplacement(spec), input, out output);
        }

        [Fact]
        public void Capture_groups_work()
        {
            var spec = new RegexReplaceSpec { Pattern = @"(\d+)\.(\d+)", Replacement = "$2.$1" };
            Assert.True(Run(spec, "I902.3", out string outp));
            Assert.Equal("I3.902", outp);
        }

        [Fact]
        public void Ignore_case_option()
        {
            var spec = new RegexReplaceSpec { Pattern = "auto", Replacement = "MAN", IgnoreCase = true };
            Assert.True(Run(spec, "AUTO1_Z01", out string outp));
            Assert.Equal("MAN1_Z01", outp);

            spec.IgnoreCase = false;
            Assert.False(Run(spec, "AUTO1_Z01", out _));
        }

        [Fact]
        public void Literal_mode_escapes_pattern_and_replacement()
        {
            // "." must be a literal dot, and "$1" in the replacement must stay literal text.
            var spec = new RegexReplaceSpec { Pattern = "192.168", Replacement = "$10.0", LiteralPattern = true };
            Assert.True(Run(spec, "192.168.0.5", out string outp));
            Assert.Equal("$10.0.0.5", outp);

            // Without literal mode, "." matches any char — "192x168" would match too.
            var regexSpec = new RegexReplaceSpec { Pattern = "192.168", Replacement = "X" };
            Assert.True(Run(regexSpec, "192x168", out _));
            Assert.False(Run(spec, "192x168", out _));   // literal mode: no match
        }

        [Fact]
        public void Whole_cell_anchors_the_pattern()
        {
            var spec = new RegexReplaceSpec { Pattern = "7", Replacement = "9", WholeCell = true };
            Assert.False(Run(spec, "007", out _));   // substring only — no whole-cell match

            var spec2 = new RegexReplaceSpec { Pattern = "0*7", Replacement = "9", WholeCell = true };
            Assert.True(Run(spec2, "007", out string outp));
            Assert.Equal("9", outp);
        }

        [Fact]
        public void No_change_and_empty_input_return_false()
        {
            var spec = new RegexReplaceSpec { Pattern = "a", Replacement = "a" };
            Assert.False(Run(spec, "banana", out _));   // matches but nothing changes
            Assert.False(Run(spec, "", out _));
            Assert.False(Run(spec, null, out _));
        }

        [Fact]
        public void Invalid_pattern_throws_argument_exception()
        {
            var spec = new RegexReplaceSpec { Pattern = "([unclosed" };
            Assert.ThrowsAny<ArgumentException>(() => RegexReplacer.Build(spec));
        }

        [Fact]
        public void Multiple_matches_in_one_cell_all_replaced()
        {
            var spec = new RegexReplaceSpec { Pattern = @"\bZ0(\d)\b", Replacement = "Z1$1" };
            Assert.True(Run(spec, "Z01;Z02;Z03", out string outp));
            Assert.Equal("Z11;Z12;Z13", outp);
        }
    }
}
