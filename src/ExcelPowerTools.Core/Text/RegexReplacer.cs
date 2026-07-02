using System;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.Text
{
    /// <summary>Everything the user chose in the find/replace dialog.</summary>
    public sealed class RegexReplaceSpec
    {
        public string Pattern { get; set; } = "";
        public string Replacement { get; set; } = "";
        public bool IgnoreCase { get; set; }

        /// <summary>Plain-text mode: pattern AND replacement are taken literally (no regex, no $1).</summary>
        public bool LiteralPattern { get; set; }

        /// <summary>The pattern must match the entire cell, not a substring.</summary>
        public bool WholeCell { get; set; }
    }

    /// <summary>One pending cell replacement (1-based sheet coordinates).</summary>
    public sealed class RegexCellChange
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Before { get; set; }
        public string After { get; set; }
    }

    /// <summary>The pure regex semantics behind the find/replace command.</summary>
    public static class RegexReplacer
    {
        /// <summary>Compiles the spec (throws ArgumentException on an invalid pattern). 2s timeout per cell.</summary>
        public static Regex Build(RegexReplaceSpec spec)
        {
            string pattern = spec.LiteralPattern ? Regex.Escape(spec.Pattern ?? "") : (spec.Pattern ?? "");
            if (spec.WholeCell) pattern = "^(?:" + pattern + ")$";

            var options = RegexOptions.CultureInvariant;
            if (spec.IgnoreCase) options |= RegexOptions.IgnoreCase;

            return new Regex(pattern, options, TimeSpan.FromSeconds(2));
        }

        /// <summary>The replacement string to feed Regex.Replace ("$" escaped in plain-text mode).</summary>
        public static string EffectiveReplacement(RegexReplaceSpec spec)
        {
            string r = spec.Replacement ?? "";
            return spec.LiteralPattern ? r.Replace("$", "$$") : r;
        }

        /// <summary>
        /// True (with the new value) when <paramref name="input"/> matches and the replacement actually
        /// changes it. Empty cells never match — a replace can edit values, not fill blanks.
        /// </summary>
        public static bool TryReplace(Regex regex, string replacement, string input, out string output)
        {
            output = null;
            if (string.IsNullOrEmpty(input)) return false;
            if (!regex.IsMatch(input)) return false;

            string result = regex.Replace(input, replacement);
            if (result == input) return false;

            output = result;
            return true;
        }
    }
}
