using System;
using System.Globalization;

namespace ExcelPowerTools.Core.Text
{
    /// <summary>
    /// The sequence behind smart fill-down: a seed whose TRAILING digit run increments per step —
    /// "AUTO1_Z07" → Z08, Z09, Z10 (zero-padding preserved), "192.168.0.5" → .6, .7 (the last octet is
    /// the trailing run). Two seeds with the same prefix infer the step (10, 20 → 30, 40). A seed with
    /// no trailing digits degrades to plain copy.
    /// </summary>
    public sealed class FillSequence
    {
        private readonly string _prefix;
        private readonly long _start;
        private readonly int _width;
        private readonly long _step;
        private readonly string _copyValue;   // non-null => plain-copy mode

        /// <summary>True when the second seed matched and defined the step — the fill then starts below it.</summary>
        public bool TwoSeedsConsumed { get; }

        private FillSequence(string prefix, long start, int width, long step, bool twoSeeds)
        {
            _prefix = prefix;
            _start = start;
            _width = width;
            _step = step;
            TwoSeedsConsumed = twoSeeds;
        }

        private FillSequence(string copyValue) => _copyValue = copyValue;

        /// <summary>Builds the sequence from the first cell (and optionally the second, for step inference).</summary>
        public static FillSequence FromSeeds(string seed1, string seed2 = null)
        {
            if (!TryParseTrailingNumber(seed1, out string prefix, out long v1, out int width))
                return new FillSequence(seed1 ?? "");

            if (seed2 != null
                && TryParseTrailingNumber(seed2, out string prefix2, out long v2, out int width2)
                && string.Equals(prefix, prefix2, StringComparison.Ordinal))
            {
                return new FillSequence(prefix, v1, Math.Max(width, width2), v2 - v1, twoSeeds: true);
            }

            return new FillSequence(prefix, v1, width, step: 1, twoSeeds: false);
        }

        /// <summary>The value at position <paramref name="k"/> (k = 0 is the first seed). Never below 0.</summary>
        public string ValueAt(int k)
        {
            if (_copyValue != null) return _copyValue;

            long value = _start + _step * k;
            if (value < 0) value = 0;
            return Compose(_prefix, value, _width);
        }

        /// <summary>Splits "AUTO1_Z07" into prefix "AUTO1_Z", value 7, width 2. False when no trailing digits.</summary>
        public static bool TryParseTrailingNumber(string s, out string prefix, out long value, out int width)
        {
            prefix = null; value = 0; width = 0;
            if (string.IsNullOrEmpty(s)) return false;

            int end = s.Length;
            int start = end;
            while (start > 0 && s[start - 1] >= '0' && s[start - 1] <= '9') start--;
            if (start == end) return false;                    // no trailing digits
            if (end - start > 18) return false;                // avoid long overflow on absurd runs

            prefix = s.Substring(0, start);
            width = end - start;
            value = long.Parse(s.Substring(start), CultureInfo.InvariantCulture);
            return true;
        }

        /// <summary>Rebuilds prefix + value, zero-padded to at least <paramref name="width"/> digits.</summary>
        public static string Compose(string prefix, long value, int width)
        {
            string digits = value.ToString(CultureInfo.InvariantCulture);
            if (digits.Length < width) digits = digits.PadLeft(width, '0');
            return prefix + digits;
        }
    }
}
