using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Diagnostics
{
    public sealed class FindingEvidence
    {
        public FindingEvidence()
        {
        }

        public FindingEvidence(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public string Key { get; set; } = "";
        public string Value { get; set; } = "";
    }

    /// <summary>
    /// Builds an aggregation identity from stable semantic evidence. Severity, localized prose, filenames and
    /// volatile physical row numbers are deliberately not accepted by this API.
    /// </summary>
    public static class FindingUidBuilder
    {
        public static string Build(string code, string semanticTarget, params FindingEvidence[] evidence)
        {
            return Build(code, semanticTarget, (IEnumerable<FindingEvidence>)evidence);
        }

        public static string Build(
            string code,
            string semanticTarget,
            IEnumerable<FindingEvidence> evidence)
        {
            if (string.IsNullOrWhiteSpace(code))
                throw new ArgumentException("A diagnostic code is required.", nameof(code));
            if (string.IsNullOrWhiteSpace(semanticTarget))
                throw new ArgumentException("A semantic target is required.", nameof(semanticTarget));

            var ordered = new List<FindingEvidence>();
            if (evidence != null)
            {
                foreach (FindingEvidence item in evidence)
                {
                    if (item == null)
                        throw new ArgumentException("Finding evidence cannot contain null entries.", nameof(evidence));
                    if (string.IsNullOrWhiteSpace(item.Key))
                        throw new ArgumentException("Every finding evidence entry requires a key.", nameof(evidence));
                    ordered.Add(item);
                }
            }

            ordered.Sort(CompareEvidence);

            var components = new string[3 + ordered.Count * 2];
            components[0] = code.Trim().ToUpperInvariant();
            components[1] = semanticTarget;
            components[2] = ordered.Count.ToString(CultureInfo.InvariantCulture);
            for (int i = 0; i < ordered.Count; i++)
            {
                components[3 + i * 2] = ordered[i].Key;
                components[4 + i * 2] = ordered[i].Value;
            }

            return StableId.Compute("diagnostic-finding-v1", components);
        }

        private static int CompareEvidence(FindingEvidence left, FindingEvidence right)
        {
            int key = string.Compare(left.Key, right.Key, StringComparison.Ordinal);
            return key != 0 ? key : string.Compare(left.Value, right.Value, StringComparison.Ordinal);
        }
    }
}
