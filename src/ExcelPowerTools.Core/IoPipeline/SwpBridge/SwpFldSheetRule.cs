using System;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The workspace's sheet-name exception to the FLD default: on destination (SWP) sheets whose
    /// name matches <see cref="Pattern"/>, a FU/Location mismatch takes <see cref="Mode"/> instead
    /// of the workspace default. ONE pattern is enough because the treatment is binary — the regex
    /// names the set of sheets that take the OTHER mode, and everything else follows the default.
    ///
    /// The mode is stored as the mode it chose, never as a "flip" bool (the same rule the per-row
    /// [Append FLD] / [Overwrite FLD] override follows): changing the workspace default later can
    /// never silently reverse what a recorded rule does to a sheet.
    ///
    /// The FLD rule holds either way — a matched sheet still carries the IO list's functional unit,
    /// location and device, only the place it lands changes. A pattern that does not compile is
    /// INERT, never fatal: the build falls back to the workspace default and says so in the log,
    /// because losing a build over a typo is worse than treating the typo as "no exception".
    /// </summary>
    public sealed class SwpFldSheetRule
    {
        /// <summary>A pathological pattern must not hang a build; a timeout reads as "no match".</summary>
        public static readonly TimeSpan MatchTimeout = TimeSpan.FromMilliseconds(250);

        /// <summary>No exception configured — every sheet follows the workspace default.</summary>
        public static readonly SwpFldSheetRule None =
            new SwpFldSheetRule("", SwpMismatchMode.AppendToDescription, null, "");

        private readonly Regex _regex;

        private SwpFldSheetRule(string pattern, SwpMismatchMode mode, Regex regex, string error)
        {
            Pattern = pattern ?? "";
            Mode = mode;
            _regex = regex;
            Error = error ?? "";
        }

        public string Pattern { get; }

        /// <summary>The treatment a MATCHING sheet takes — the explicit mode, not "the opposite".</summary>
        public SwpMismatchMode Mode { get; }

        /// <summary>Why the pattern was rejected ("" when it compiled or none was given).</summary>
        public string Error { get; }

        /// <summary>true when a usable pattern is in force; false for "none" and for a bad pattern.</summary>
        public bool IsActive => _regex != null;

        /// <summary>
        /// Compiles a rule. An empty/blank pattern yields <see cref="None"/>; a pattern that does
        /// not parse yields an inactive rule carrying <see cref="Error"/> for the caller to report.
        /// Matching is case-insensitive: SWP sheet names are upper-case by convention, and a rule
        /// that silently missed because of case would apply the wrong FLD treatment to a whole sheet.
        /// </summary>
        public static SwpFldSheetRule Create(string pattern, SwpMismatchMode mode)
        {
            string trimmed = (pattern ?? "").Trim();
            if (trimmed.Length == 0) return None;
            try
            {
                var regex = new Regex(trimmed,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, MatchTimeout);
                return new SwpFldSheetRule(trimmed, mode, regex, "");
            }
            catch (ArgumentException ex)
            {
                return new SwpFldSheetRule(trimmed, mode, null, ex.Message);
            }
        }

        /// <summary>
        /// The treatment this rule imposes on <paramref name="sheetName"/>, or null when the rule
        /// is inactive or the sheet does not match — null means "the caller's default stands".
        /// </summary>
        public SwpMismatchMode? ModeFor(string sheetName)
        {
            if (_regex == null || string.IsNullOrEmpty(sheetName)) return null;
            try { return _regex.IsMatch(sheetName) ? Mode : (SwpMismatchMode?)null; }
            catch (RegexMatchTimeoutException) { return null; }
        }
    }
}
