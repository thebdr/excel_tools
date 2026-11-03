using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The dated version folders of an SWP system — "NN_dd-MMM-yyyy" siblings under its root, NN
    /// counting up. Two decisions stand on them: which folder a session comes back to, and which
    /// files a new folder takes along from the current one.
    /// </summary>
    public static class SwpVersionFolders
    {
        private static readonly Regex Numbered = new Regex(@"^(\d+)_", RegexOptions.CultureInvariant);

        /// <summary>A pathological pattern must not hang an apply; a timeout reads as "no match".</summary>
        private static readonly TimeSpan MatchTimeout = TimeSpan.FromMilliseconds(250);

        /// <summary>NN of a dated version folder ("06_20-Aug-2026" → 6); null for any other name.</summary>
        public static int? Number(string folderName)
        {
            Match m = Numbered.Match(folderName ?? "");
            return m.Success && int.TryParse(m.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int n)
                ? n
                : (int?)null;
        }

        /// <summary>
        /// The newest of these folders by NN — numerically, so "10_…" beats "9_…" — or null when none
        /// of them is a dated version folder.
        /// </summary>
        public static string Newest(IEnumerable<string> folderNames)
        {
            if (folderNames == null) return null;
            string best = null;
            int bestNumber = -1;
            foreach (string name in folderNames)
            {
                int? n = Number(name);
                if (n.HasValue && n.Value > bestNumber)
                {
                    best = name;
                    bestNumber = n.Value;
                }
            }
            return best;
        }

        /// <summary>
        /// The files a new dated folder takes along from the current one, besides what apply puts
        /// there itself: every name the pattern matches (case-insensitively), in folder order.
        /// Never the SWP — apply writes the patched copy — and never CONFIG_PLANT, which is always
        /// copied; never Excel's "~$" lock files, and never ".bak" safety copies, which belong to
        /// the folder they were made in. An empty pattern takes nothing; one that does not parse
        /// takes nothing and says why in <paramref name="error"/> — inert, never fatal.
        /// </summary>
        public static List<string> FilesToCopy(IEnumerable<string> fileNames, string pattern,
            string swpFileName, string configFileName, out string error)
        {
            error = "";
            var result = new List<string>();
            string trimmed = (pattern ?? "").Trim();
            if (trimmed.Length == 0 || fileNames == null) return result;

            Regex regex;
            try
            {
                regex = new Regex(trimmed, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, MatchTimeout);
            }
            catch (ArgumentException ex)
            {
                error = ex.Message;
                return result;
            }

            foreach (string name in fileNames)
            {
                if (string.IsNullOrEmpty(name)) continue;
                if (name.StartsWith("~$", StringComparison.Ordinal)) continue;
                if (name.EndsWith(".bak", StringComparison.OrdinalIgnoreCase)) continue;
                if (string.Equals(name, swpFileName, StringComparison.OrdinalIgnoreCase)) continue;
                if (string.Equals(name, configFileName, StringComparison.OrdinalIgnoreCase)) continue;

                bool matches;
                try { matches = regex.IsMatch(name); }
                catch (RegexMatchTimeoutException) { matches = false; }
                if (matches) result.Add(name);
            }
            return result;
        }
    }
}
