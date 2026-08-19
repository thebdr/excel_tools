using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.Comparison.Revision
{
    public enum RevisionTokenKind
    {
        ExplicitPrefix,
        TrailingWholeNumber,
        TrailingSplitDecimal,
    }

    public sealed class RevisionNumber : IComparable<RevisionNumber>, IEquatable<RevisionNumber>
    {
        private readonly int[] _components;
        private readonly IReadOnlyList<int> _readOnlyComponents;

        public RevisionNumber(params int[] components)
        {
            if (components == null || components.Length == 0)
                throw new ArgumentException("At least one revision-number component is required.", nameof(components));
            for (int i = 0; i < components.Length; i++)
                if (components[i] < 0) throw new ArgumentOutOfRangeException(nameof(components));
            _components = (int[])components.Clone();
            _readOnlyComponents = new ReadOnlyCollection<int>(_components);
        }

        public IReadOnlyList<int> Components => _readOnlyComponents;

        public int CompareTo(RevisionNumber other)
        {
            if (other == null) return 1;
            int length = Math.Max(_components.Length, other._components.Length);
            for (int i = 0; i < length; i++)
            {
                int left = i < _components.Length ? _components[i] : 0;
                int right = i < other._components.Length ? other._components[i] : 0;
                int comparison = left.CompareTo(right);
                if (comparison != 0) return comparison;
            }
            return 0;
        }

        public bool Equals(RevisionNumber other) => CompareTo(other) == 0;
        public override bool Equals(object obj) => Equals(obj as RevisionNumber);

        public override int GetHashCode()
        {
            unchecked
            {
                int last = _components.Length - 1;
                while (last > 0 && _components[last] == 0) last--;
                int hash = 17;
                for (int i = 0; i <= last; i++) hash = hash * 31 + _components[i];
                return hash;
            }
        }

        public override string ToString() => string.Join(".", _components.Select(
            value => value.ToString(CultureInfo.InvariantCulture)));
    }

    public sealed class RevisionTokenCandidate
    {
        internal RevisionTokenCandidate(
            string rawToken,
            RevisionNumber number,
            RevisionTokenKind kind,
            double confidence,
            int startIndex,
            int length,
            string evidence)
        {
            RawToken = rawToken;
            Number = number;
            Kind = kind;
            Confidence = confidence;
            StartIndex = startIndex;
            Length = length;
            Evidence = evidence;
        }

        public string RawToken { get; }
        public RevisionNumber Number { get; }
        public RevisionTokenKind Kind { get; }
        public double Confidence { get; }
        public int StartIndex { get; }
        public int Length { get; }
        public string Evidence { get; }
    }

    public sealed class RevisionLabelInferenceResult
    {
        internal RevisionLabelInferenceResult(
            string baseName,
            IReadOnlyList<RevisionTokenCandidate> candidates,
            RevisionTokenCandidate strongCandidate)
        {
            BaseName = baseName;
            Candidates = new ReadOnlyCollection<RevisionTokenCandidate>(
                new List<RevisionTokenCandidate>(candidates ?? Array.Empty<RevisionTokenCandidate>()));
            StrongCandidate = strongCandidate;
        }

        public string BaseName { get; }
        public IReadOnlyList<RevisionTokenCandidate> Candidates { get; }
        public RevisionTokenCandidate StrongCandidate { get; }
        public bool HasStrongCandidate => StrongCandidate != null;
    }

    /// <summary>
    /// Conservative filename evidence only. It deliberately returns competing interpretations for terminal
    /// numbers (for example 31 and 3.1) rather than silently deciding the revision order.
    /// </summary>
    public static class RevisionLabelInference
    {
        private static readonly Regex ExplicitToken = new Regex(
            @"(?:^|[^A-Za-z0-9])(?<token>(?<prefix>revision|version|rev|ver|r|v)[\s_-]*(?<version>\d+(?:[._]\d+)*))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex TrailingDigits = new Regex(
            @"(?<digits>\d+)$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static RevisionLabelInferenceResult Infer(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("A file name is required.", nameof(fileName));

            string baseName = Path.GetFileNameWithoutExtension(fileName);
            var candidates = new List<RevisionTokenCandidate>();

            foreach (Match match in ExplicitToken.Matches(baseName))
            {
                Group token = match.Groups["token"];
                Group version = match.Groups["version"];
                RevisionNumber number;
                if (!TryParseComponents(version.Value, out number)) continue;

                string prefix = match.Groups["prefix"].Value;
                double confidence = prefix.Length > 1 ? 0.98 : 0.95;
                candidates.Add(new RevisionTokenCandidate(
                    token.Value,
                    number,
                    RevisionTokenKind.ExplicitPrefix,
                    confidence,
                    token.Index,
                    token.Length,
                    "explicit " + prefix.ToLowerInvariant() + " revision token"));
            }

            bool explicitAtEnd = candidates.Any(candidate =>
                candidate.Kind == RevisionTokenKind.ExplicitPrefix &&
                candidate.StartIndex + candidate.Length == baseName.Length);
            if (!explicitAtEnd)
                AddTrailingCandidates(baseName, candidates);

            candidates.Sort(CompareCandidates);
            RevisionTokenCandidate strong = SelectStrongCandidate(candidates);
            return new RevisionLabelInferenceResult(baseName, candidates, strong);
        }

        private static void AddTrailingCandidates(string baseName, List<RevisionTokenCandidate> candidates)
        {
            Match match = TrailingDigits.Match(baseName);
            if (!match.Success) return;

            string digits = match.Groups["digits"].Value;
            int whole;
            if (!int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out whole)) return;

            bool separated = match.Index > 0 && !char.IsLetterOrDigit(baseName[match.Index - 1]);
            double wholeConfidence = separated ? 0.45 : 0.30;
            candidates.Add(new RevisionTokenCandidate(
                digits,
                new RevisionNumber(whole),
                RevisionTokenKind.TrailingWholeNumber,
                wholeConfidence,
                match.Index,
                match.Length,
                "terminal digits interpreted as a whole revision number"));

            if (digits.Length < 2) return;
            string majorText = digits.Substring(0, digits.Length - 1);
            string minorText = digits.Substring(digits.Length - 1);
            int major;
            int minor;
            if (!int.TryParse(majorText, NumberStyles.None, CultureInfo.InvariantCulture, out major) ||
                !int.TryParse(minorText, NumberStyles.None, CultureInfo.InvariantCulture, out minor))
                return;

            candidates.Add(new RevisionTokenCandidate(
                digits,
                new RevisionNumber(major, minor),
                RevisionTokenKind.TrailingSplitDecimal,
                wholeConfidence - 0.02,
                match.Index,
                match.Length,
                "terminal digits split before the final digit (for example 31 as 3.1)"));
        }

        private static RevisionTokenCandidate SelectStrongCandidate(List<RevisionTokenCandidate> candidates)
        {
            if (candidates.Count == 0 || candidates[0].Kind != RevisionTokenKind.ExplicitPrefix) return null;
            if (candidates.Count == 1) return candidates[0];
            return candidates[0].Confidence - candidates[1].Confidence >= 0.10 ? candidates[0] : null;
        }

        private static int CompareCandidates(RevisionTokenCandidate left, RevisionTokenCandidate right)
        {
            int confidence = right.Confidence.CompareTo(left.Confidence);
            if (confidence != 0) return confidence;
            int position = right.StartIndex.CompareTo(left.StartIndex);
            if (position != 0) return position;
            return left.Kind.CompareTo(right.Kind);
        }

        private static bool TryParseComponents(string text, out RevisionNumber number)
        {
            string[] pieces = text.Split('.', '_');
            var components = new int[pieces.Length];
            for (int i = 0; i < pieces.Length; i++)
            {
                if (!int.TryParse(pieces[i], NumberStyles.None, CultureInfo.InvariantCulture, out components[i]))
                {
                    number = null;
                    return false;
                }
            }
            number = new RevisionNumber(components);
            return true;
        }
    }
}
