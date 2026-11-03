namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// What a row does with CONFIG_PLANT's template text. The two answers are exclusive — a row
    /// cannot both keep the template whole and cut its tail off — which is why they are one mode
    /// and not two independent switches.
    /// </summary>
    public enum SwpConfigMode
    {
        /// <summary>Keep CONFIG_PLANT's text: the Q/R formula is left alone and T/U receives only
        /// the FLD label. SCADA shows the template followed by that label.</summary>
        KeepConfigText,

        /// <summary>Replace the template's trailing text with the row's description: Q/R is rebuilt
        /// to keep the template up to its last placeholder and show T/U in the tail's place.</summary>
        ReplaceTemplateTail,
    }

    /// <summary>
    /// The workspace's sheet-name exception to the CONFIG_PLANT default — the same shape as
    /// <see cref="SwpFldSheetRule"/>: on destination (SWP) sheets whose name matches
    /// <see cref="Pattern"/>, rows take <see cref="Mode"/> instead of the workspace default.
    ///
    /// The mode is stored as the mode it chose, never as a "flip" bool, so changing the workspace
    /// default later can never silently reverse what a recorded rule does to a sheet. A pattern
    /// that does not compile is inert with a warning, never fatal.
    /// </summary>
    public sealed class SwpConfigSheetRule
    {
        public static readonly SwpConfigSheetRule None =
            new SwpConfigSheetRule(SwpSheetNameRule.None, SwpConfigMode.KeepConfigText);

        private readonly SwpSheetNameRule _rule;

        private SwpConfigSheetRule(SwpSheetNameRule rule, SwpConfigMode mode)
        {
            _rule = rule;
            Mode = mode;
        }

        /// <summary>The treatment a MATCHING sheet takes — the explicit mode, not "the opposite".</summary>
        public SwpConfigMode Mode { get; }

        public string Pattern => _rule.Pattern;
        public string Error => _rule.Error;
        public bool IsActive => _rule.IsActive;

        public static SwpConfigSheetRule Create(string pattern, SwpConfigMode mode)
        {
            SwpSheetNameRule rule = SwpSheetNameRule.Create(pattern);
            return ReferenceEquals(rule, SwpSheetNameRule.None)
                ? None
                : new SwpConfigSheetRule(rule, mode);
        }

        /// <summary>
        /// The treatment this rule imposes on <paramref name="sheetName"/>, or null when the rule
        /// is inactive or the sheet does not match — null means "the caller's default stands".
        /// </summary>
        public SwpConfigMode? ModeFor(string sheetName) =>
            _rule.Matches(sheetName) ? Mode : (SwpConfigMode?)null;
    }
}
