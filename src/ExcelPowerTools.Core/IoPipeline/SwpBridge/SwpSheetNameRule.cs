using System;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The regex mechanics shared by every workspace default that may carry ONE destination-sheet
    /// exception (the FLD treatment, the CONFIG_PLANT treatment). One pattern is enough because
    /// each such setting is binary: the regex names the sheets that take the OTHER mode, and
    /// everything else follows the default.
    ///
    /// Matching is case-insensitive — SWP sheet names are upper-case by convention and a rule that
    /// silently missed on case would apply the wrong treatment to a whole sheet — and carries a
    /// timeout, so a pathological pattern reads as "no match" instead of hanging a build. A pattern
    /// that does not compile is INERT, never fatal: <see cref="Error"/> is set for the caller to
    /// report and every sheet follows the default, because losing a build over a typo is worse
    /// than treating the typo as "no exception".
    /// </summary>
    public sealed class SwpSheetNameRule
    {
        public static readonly TimeSpan MatchTimeout = TimeSpan.FromMilliseconds(250);

        /// <summary>No exception configured — every sheet follows the workspace default.</summary>
        public static readonly SwpSheetNameRule None = new SwpSheetNameRule("", null, "");

        private readonly Regex _regex;

        private SwpSheetNameRule(string pattern, Regex regex, string error)
        {
            Pattern = pattern ?? "";
            _regex = regex;
            Error = error ?? "";
        }

        public string Pattern { get; }

        /// <summary>Why the pattern was rejected ("" when it compiled or none was given).</summary>
        public string Error { get; }

        /// <summary>true when a usable pattern is in force; false for "none" and for a bad one.</summary>
        public bool IsActive => _regex != null;

        public static SwpSheetNameRule Create(string pattern)
        {
            string trimmed = (pattern ?? "").Trim();
            if (trimmed.Length == 0) return None;
            try
            {
                return new SwpSheetNameRule(trimmed,
                    new Regex(trimmed, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, MatchTimeout), "");
            }
            catch (ArgumentException ex)
            {
                return new SwpSheetNameRule(trimmed, null, ex.Message);
            }
        }

        public bool Matches(string sheetName)
        {
            if (_regex == null || string.IsNullOrEmpty(sheetName)) return false;
            try { return _regex.IsMatch(sheetName); }
            catch (RegexMatchTimeoutException) { return false; }
        }
    }
}
