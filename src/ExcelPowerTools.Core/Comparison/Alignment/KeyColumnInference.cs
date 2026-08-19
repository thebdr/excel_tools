using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    public enum KeyInferenceDecision
    {
        HeaderRequired,
        None,
        Ambiguous,
        StrongKey,
        ConfirmedKey,
    }

    public sealed class KeyInferenceOptions
    {
        public int MaxCompositeColumns { get; set; } = 2;
        public int MaxCandidateHeaderColumns { get; set; } = 32;
        public double StrongScoreThreshold { get; set; } = 0.82D;
        public double StrongRunnerUpMargin { get; set; } = 0.06D;
        public double MinimumCoverage { get; set; } = 0.95D;
        public double MinimumUniqueness { get; set; } = 0.98D;
        public double MinimumCrossRevisionOverlap { get; set; } = 0.60D;

        /// <summary>
        /// Optional explicit mapping. Both lists must be present, non-empty, and have equal length. Confirmation
        /// authorizes the columns as a key but does not hide duplicate-key ambiguity during row alignment.
        /// </summary>
        public IReadOnlyList<int> ConfirmedBeforeColumns { get; set; }
        public IReadOnlyList<int> ConfirmedAfterColumns { get; set; }
    }

    public sealed class KeyCandidate
    {
        internal KeyCandidate(
            IEnumerable<int> beforeColumns,
            IEnumerable<int> afterColumns,
            double score,
            AlignmentConfidence confidence,
            double beforeCoverage,
            double afterCoverage,
            double beforeUniqueness,
            double afterUniqueness,
            double crossRevisionOverlap,
            int beforeDuplicateCount,
            int afterDuplicateCount,
            IEnumerable<AlignmentReason> reasons)
        {
            BeforeColumns = Freeze(beforeColumns);
            AfterColumns = Freeze(afterColumns);
            Score = score;
            Confidence = confidence;
            BeforeCoverage = beforeCoverage;
            AfterCoverage = afterCoverage;
            BeforeUniqueness = beforeUniqueness;
            AfterUniqueness = afterUniqueness;
            CrossRevisionOverlap = crossRevisionOverlap;
            BeforeDuplicateCount = beforeDuplicateCount;
            AfterDuplicateCount = afterDuplicateCount;
            Reasons = new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
        }

        public IReadOnlyList<int> BeforeColumns { get; }
        public IReadOnlyList<int> AfterColumns { get; }
        public double Score { get; }
        public AlignmentConfidence Confidence { get; }
        public double BeforeCoverage { get; }
        public double AfterCoverage { get; }
        public double BeforeUniqueness { get; }
        public double AfterUniqueness { get; }
        public double CrossRevisionOverlap { get; }
        public int BeforeDuplicateCount { get; }
        public int AfterDuplicateCount { get; }
        public bool HasDuplicates => BeforeDuplicateCount > 0 || AfterDuplicateCount > 0;
        public IReadOnlyList<AlignmentReason> Reasons { get; }

        private static IReadOnlyList<int> Freeze(IEnumerable<int> values) =>
            new ReadOnlyCollection<int>(new List<int>(values ?? Array.Empty<int>()));
    }

    public sealed class KeyInferenceResult
    {
        internal KeyInferenceResult(
            KeyInferenceDecision decision,
            KeyCandidate selectedCandidate,
            IEnumerable<KeyCandidate> candidates,
            double runnerUpMargin,
            IEnumerable<AlignmentReason> reasons)
        {
            Decision = decision;
            SelectedCandidate = selectedCandidate;
            Candidates = new ReadOnlyCollection<KeyCandidate>(
                new List<KeyCandidate>(candidates ?? Array.Empty<KeyCandidate>()));
            RunnerUpMargin = runnerUpMargin;
            Reasons = new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
        }

        public KeyInferenceDecision Decision { get; }
        public KeyCandidate SelectedCandidate { get; }
        public IReadOnlyList<KeyCandidate> Candidates { get; }
        public double RunnerUpMargin { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }
        public bool HasUsableKey => SelectedCandidate != null &&
                                    (Decision == KeyInferenceDecision.StrongKey ||
                                     Decision == KeyInferenceDecision.ConfirmedKey);
    }

    /// <summary>
    /// Scores single and small composite keys only after both revisions have a strong/confirmed header. Header
    /// labels map physical columns; exact, length-prefixed cell tuples establish row keys without trimming,
    /// case-folding, or otherwise discarding source evidence.
    /// </summary>
    public static class KeyColumnInference
    {
        public static KeyInferenceResult Infer(
            AlignmentTable before,
            AlignmentTable after,
            HeaderDetectionResult beforeHeader,
            HeaderDetectionResult afterHeader,
            KeyInferenceOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (after == null) throw new ArgumentNullException(nameof(after));
            if (beforeHeader == null) throw new ArgumentNullException(nameof(beforeHeader));
            if (afterHeader == null) throw new ArgumentNullException(nameof(afterHeader));
            options = options ?? new KeyInferenceOptions();
            ValidateOptions(options);
            cancellationToken.ThrowIfCancellationRequested();

            bool hasConfirmed = options.ConfirmedBeforeColumns != null || options.ConfirmedAfterColumns != null;
            if (hasConfirmed)
            {
                ValidateConfirmedColumns(before, after, options);
                KeyCandidate confirmed = Evaluate(
                    before,
                    after,
                    beforeHeader.DataStartRowIndex,
                    afterHeader.DataStartRowIndex,
                    options.ConfirmedBeforeColumns,
                    options.ConfirmedAfterColumns,
                    1D,
                    options,
                    true,
                    cancellationToken);
                return new KeyInferenceResult(
                    KeyInferenceDecision.ConfirmedKey,
                    confirmed,
                    new[] { confirmed },
                    1D,
                    new[]
                    {
                        new AlignmentReason(
                            "key.user-confirmed",
                            confirmed.HasDuplicates
                                ? "The user confirmed these key columns; duplicate values remain explicit ambiguities."
                                : "The user explicitly confirmed these key columns."),
                    });
            }

            if (!beforeHeader.HasAcceptedHeader || !afterHeader.HasAcceptedHeader)
                return new KeyInferenceResult(
                    KeyInferenceDecision.HeaderRequired,
                    null,
                    Array.Empty<KeyCandidate>(),
                    0D,
                    new[]
                    {
                        new AlignmentReason(
                            "key.header-required",
                            "Key inference did not run because both revisions need a strong or confirmed header."),
                    });

            IReadOnlyList<ColumnPair> pairs = CommonUniqueHeaderColumns(
                before,
                after,
                beforeHeader,
                afterHeader,
                cancellationToken);
            if (pairs.Count == 0)
                return new KeyInferenceResult(
                    KeyInferenceDecision.None,
                    null,
                    Array.Empty<KeyCandidate>(),
                    0D,
                    new[]
                    {
                        new AlignmentReason(
                            "key.no-common-unique-headers",
                            "No uniquely named header columns could be mapped across the two revisions."),
                    });

            int limit = Math.Min(pairs.Count, options.MaxCandidateHeaderColumns);
            var candidates = new List<KeyCandidate>();
            for (int index = 0; index < limit; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ColumnPair pair = pairs[index];
                candidates.Add(Evaluate(
                    before,
                    after,
                    beforeHeader.DataStartRowIndex,
                    afterHeader.DataStartRowIndex,
                    new[] { pair.Before },
                    new[] { pair.After },
                    HeaderDetector.KeyAliasStrength(pair.NormalizedHeader),
                    options,
                    false,
                    cancellationToken));
            }

            bool hasNearlyUniqueSingle = candidates.Any(candidate =>
                candidate.BeforeCoverage >= options.MinimumCoverage &&
                candidate.AfterCoverage >= options.MinimumCoverage &&
                candidate.BeforeUniqueness >= options.MinimumUniqueness &&
                candidate.AfterUniqueness >= options.MinimumUniqueness &&
                candidate.CrossRevisionOverlap >= options.MinimumCrossRevisionOverlap);

            if (!hasNearlyUniqueSingle && options.MaxCompositeColumns >= 2)
            {
                for (int left = 0; left < limit; left++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int right = left + 1; right < limit; right++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        ColumnPair first = pairs[left];
                        ColumnPair second = pairs[right];
                        double alias = (HeaderDetector.KeyAliasStrength(first.NormalizedHeader) +
                                        HeaderDetector.KeyAliasStrength(second.NormalizedHeader)) / 2D;
                        candidates.Add(Evaluate(
                            before,
                            after,
                            beforeHeader.DataStartRowIndex,
                            afterHeader.DataStartRowIndex,
                            new[] { first.Before, second.Before },
                            new[] { first.After, second.After },
                            alias,
                            options,
                            false,
                            cancellationToken));
                    }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            candidates.Sort(CompareCandidates);
            cancellationToken.ThrowIfCancellationRequested();
            KeyCandidate best = candidates.Count == 0 ? null : candidates[0];
            double runner = candidates.Count > 1 ? candidates[1].Score : 0D;
            double margin = best == null ? 0D : Math.Max(0D, best.Score - runner);
            bool bestPasses = best != null && PassesStrongGates(best, options);

            if (bestPasses && margin >= options.StrongRunnerUpMargin)
                return new KeyInferenceResult(
                    KeyInferenceDecision.StrongKey,
                    best,
                    candidates,
                    margin,
                    new[]
                    {
                        new AlignmentReason(
                            "key.strong",
                            "The top key passed coverage, uniqueness, overlap, score, duplicate, and margin gates.",
                            best.Score),
                    });

            if (best != null && (best.Score >= 0.55D || best.HasDuplicates || bestPasses))
                return new KeyInferenceResult(
                    KeyInferenceDecision.Ambiguous,
                    null,
                    candidates,
                    margin,
                    new[]
                    {
                        new AlignmentReason(
                            best.HasDuplicates ? "key.duplicates" : "key.runner-up-ambiguity",
                            best.HasDuplicates
                                ? "The best key candidate contains duplicate values and was not inferred as a unique key."
                                : "No candidate was sufficiently separated from alternatives; row alignment must fall back."),
                    });

            return new KeyInferenceResult(
                KeyInferenceDecision.None,
                null,
                candidates,
                margin,
                new[]
                {
                    new AlignmentReason(
                        "key.no-strong-candidate",
                        "No key candidate met the conservative inference gates."),
                });
        }

        private static KeyCandidate Evaluate(
            AlignmentTable before,
            AlignmentTable after,
            int beforeStart,
            int afterStart,
            IReadOnlyList<int> beforeColumns,
            IReadOnlyList<int> afterColumns,
            double aliasStrength,
            KeyInferenceOptions options,
            bool confirmed,
            CancellationToken cancellationToken)
        {
            KeyStats beforeStats = Analyze(before, beforeStart, beforeColumns, cancellationToken);
            KeyStats afterStats = Analyze(after, afterStart, afterColumns, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            int intersection = beforeStats.Keys.Intersect(afterStats.Keys, StringComparer.Ordinal).Count();
            cancellationToken.ThrowIfCancellationRequested();
            int denominator = Math.Max(beforeStats.Keys.Count, afterStats.Keys.Count);
            double overlap = denominator == 0 ? 0D : (double)intersection / denominator;
            double typeConsistency = TypeConsistency(
                before,
                after,
                beforeStart,
                afterStart,
                beforeColumns,
                afterColumns,
                cancellationToken);
            double coverage = Math.Min(beforeStats.Coverage, afterStats.Coverage);
            double uniqueness = Math.Min(beforeStats.Uniqueness, afterStats.Uniqueness);
            double score = Clamp(
                0.22D * coverage +
                0.28D * uniqueness +
                0.30D * overlap +
                0.08D * typeConsistency +
                0.12D * aliasStrength);

            AlignmentConfidence confidence = confirmed
                ? AlignmentConfidence.Confirmed
                : score >= options.StrongScoreThreshold &&
                  beforeStats.DuplicateCount == 0 && afterStats.DuplicateCount == 0
                    ? AlignmentConfidence.Strong
                    : score >= 0.65D
                        ? AlignmentConfidence.Medium
                        : score >= 0.45D
                            ? AlignmentConfidence.Weak
                            : AlignmentConfidence.None;

            var reasons = new List<AlignmentReason>
            {
                Metric("key.coverage", "Minimum nonblank coverage", coverage, 0.22D),
                Metric("key.uniqueness", "Minimum within-revision uniqueness", uniqueness, 0.28D),
                Metric("key.cross-revision-overlap", "Exact key overlap across revisions", overlap, 0.30D),
                Metric("key.type-consistency", "Column type consistency", typeConsistency, 0.08D),
                Metric("key.header-alias", "Key-like header evidence", aliasStrength, 0.12D),
                new AlignmentReason(
                    "key.duplicate-count",
                    "Duplicate surplus: Before=" + beforeStats.DuplicateCount.ToString(CultureInfo.InvariantCulture) +
                    ", After=" + afterStats.DuplicateCount.ToString(CultureInfo.InvariantCulture) + "."),
                new AlignmentReason(
                    "key.exact-tuples",
                    "Key matching uses exact length-prefixed cell tuples; normalization is limited to header mapping."),
            };

            return new KeyCandidate(
                beforeColumns,
                afterColumns,
                score,
                confidence,
                beforeStats.Coverage,
                afterStats.Coverage,
                beforeStats.Uniqueness,
                afterStats.Uniqueness,
                overlap,
                beforeStats.DuplicateCount,
                afterStats.DuplicateCount,
                reasons);
        }

        private static KeyStats Analyze(
            AlignmentTable table,
            int dataStart,
            IReadOnlyList<int> columns,
            CancellationToken cancellationToken)
        {
            int dataRows = Math.Max(0, table.RowCount - dataStart);
            int populated = 0;
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int row = dataStart; row < table.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                bool complete = true;
                foreach (int column in columns)
                {
                    if (!table.TryGetCell(row, column, out string value) || AlignmentText.IsBlank(value))
                    {
                        complete = false;
                        break;
                    }
                }
                if (!complete) continue;
                populated++;
                string key = AlignmentText.ExactTuple(table, row, columns);
                counts.TryGetValue(key, out int count);
                counts[key] = count + 1;
            }

            int duplicates = 0;
            foreach (int count in counts.Values)
                if (count > 1) duplicates += count - 1;
            double coverage = dataRows == 0 ? 0D : (double)populated / dataRows;
            double uniqueness = populated == 0 ? 0D : (double)counts.Count / populated;
            return new KeyStats(coverage, uniqueness, duplicates, new HashSet<string>(counts.Keys, StringComparer.Ordinal));
        }

        private static double TypeConsistency(
            AlignmentTable before,
            AlignmentTable after,
            int beforeStart,
            int afterStart,
            IReadOnlyList<int> beforeColumns,
            IReadOnlyList<int> afterColumns,
            CancellationToken cancellationToken)
        {
            double total = 0D;
            for (int index = 0; index < beforeColumns.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                TypeStats beforeType = AnalyzeType(
                    before,
                    beforeStart,
                    beforeColumns[index],
                    cancellationToken);
                TypeStats afterType = AnalyzeType(
                    after,
                    afterStart,
                    afterColumns[index],
                    cancellationToken);
                double consistency = Math.Min(beforeType.MajorityRatio, afterType.MajorityRatio);
                if (!string.Equals(beforeType.MajorityKind, afterType.MajorityKind, StringComparison.Ordinal))
                    consistency *= 0.25D;
                total += consistency;
            }
            return beforeColumns.Count == 0 ? 0D : total / beforeColumns.Count;
        }

        private static TypeStats AnalyzeType(
            AlignmentTable table,
            int dataStart,
            int column,
            CancellationToken cancellationToken)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            int observed = 0;
            for (int row = dataStart; row < table.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!table.TryGetCell(row, column, out string value) || AlignmentText.IsBlank(value)) continue;
                observed++;
                string kind = ValueKind(value);
                counts.TryGetValue(kind, out int count);
                counts[kind] = count + 1;
            }
            if (observed == 0) return new TypeStats("none", 0D);
            KeyValuePair<string, int> best = counts
                .OrderByDescending(pair => pair.Value)
                .ThenBy(pair => pair.Key, StringComparer.Ordinal)
                .First();
            return new TypeStats(best.Key, (double)best.Value / observed);
        }

        private static string ValueKind(string value)
        {
            string trimmed = value.Trim();
            if (double.TryParse(
                trimmed,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture,
                out _)) return "number";
            if (bool.TryParse(trimmed, out _)) return "boolean";
            if (Guid.TryParse(trimmed, out _)) return "guid";
            if (trimmed.Any(char.IsDigit) && trimmed.Any(char.IsLetter)) return "alphanumeric";
            return "text";
        }

        private static IReadOnlyList<ColumnPair> CommonUniqueHeaderColumns(
            AlignmentTable before,
            AlignmentTable after,
            HeaderDetectionResult beforeHeader,
            HeaderDetectionResult afterHeader,
            CancellationToken cancellationToken)
        {
            Dictionary<string, List<int>> beforeColumns = HeaderColumns(
                before,
                beforeHeader.AcceptedHeaderRowIndex.Value,
                cancellationToken);
            Dictionary<string, List<int>> afterColumns = HeaderColumns(
                after,
                afterHeader.AcceptedHeaderRowIndex.Value,
                cancellationToken);
            var labels = new SortedSet<string>(beforeColumns.Keys, StringComparer.Ordinal);
            labels.IntersectWith(afterColumns.Keys);
            var result = new List<ColumnPair>();
            foreach (string label in labels)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (label.Length == 0 || beforeColumns[label].Count != 1 || afterColumns[label].Count != 1) continue;
                result.Add(new ColumnPair(beforeColumns[label][0], afterColumns[label][0], label));
            }
            return new ReadOnlyCollection<ColumnPair>(result);
        }

        private static Dictionary<string, List<int>> HeaderColumns(
            AlignmentTable table,
            int headerRow,
            CancellationToken cancellationToken)
        {
            var result = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            AlignmentRow row = table.Rows[headerRow];
            for (int column = 0; column < row.Count; column++)
            {
                if ((column & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                string normalized = AlignmentText.NormalizeLabel(row.Values[column]);
                if (!result.TryGetValue(normalized, out List<int> indices))
                {
                    indices = new List<int>();
                    result.Add(normalized, indices);
                }
                indices.Add(column);
            }
            return result;
        }

        private static bool PassesStrongGates(KeyCandidate candidate, KeyInferenceOptions options) =>
            candidate.Score >= options.StrongScoreThreshold &&
            candidate.BeforeCoverage >= options.MinimumCoverage &&
            candidate.AfterCoverage >= options.MinimumCoverage &&
            candidate.BeforeUniqueness >= options.MinimumUniqueness &&
            candidate.AfterUniqueness >= options.MinimumUniqueness &&
            candidate.CrossRevisionOverlap >= options.MinimumCrossRevisionOverlap &&
            !candidate.HasDuplicates;

        private static int CompareCandidates(KeyCandidate left, KeyCandidate right)
        {
            int score = right.Score.CompareTo(left.Score);
            if (score != 0) return score;
            int width = left.BeforeColumns.Count.CompareTo(right.BeforeColumns.Count);
            if (width != 0) return width;
            for (int index = 0; index < Math.Min(left.BeforeColumns.Count, right.BeforeColumns.Count); index++)
            {
                int column = left.BeforeColumns[index].CompareTo(right.BeforeColumns[index]);
                if (column != 0) return column;
            }
            return 0;
        }

        private static void ValidateConfirmedColumns(
            AlignmentTable before,
            AlignmentTable after,
            KeyInferenceOptions options)
        {
            if (options.ConfirmedBeforeColumns == null || options.ConfirmedAfterColumns == null ||
                options.ConfirmedBeforeColumns.Count == 0 ||
                options.ConfirmedBeforeColumns.Count != options.ConfirmedAfterColumns.Count)
                throw new ArgumentException(
                    "Confirmed Before/After key columns must both be non-empty and have equal length.",
                    nameof(options));
            if (options.ConfirmedBeforeColumns.Count > options.MaxCompositeColumns)
                throw new ArgumentException("The confirmed key exceeds the configured composite-key width.", nameof(options));
            foreach (int column in options.ConfirmedBeforeColumns)
                if (column < 0 || column >= before.ColumnCount)
                    throw new ArgumentOutOfRangeException(nameof(options), "A confirmed Before key column is outside the table.");
            foreach (int column in options.ConfirmedAfterColumns)
                if (column < 0 || column >= after.ColumnCount)
                    throw new ArgumentOutOfRangeException(nameof(options), "A confirmed After key column is outside the table.");
        }

        private static void ValidateOptions(KeyInferenceOptions options)
        {
            // Automatic inference intentionally considers at most two columns, but explicit confirmations may
            // retain wider domain keys (for example group/item/channel) without being silently truncated.
            if (options.MaxCompositeColumns <= 0 || options.MaxCompositeColumns > 4)
                throw new ArgumentOutOfRangeException(nameof(options.MaxCompositeColumns));
            if (options.MaxCandidateHeaderColumns <= 0 || options.MaxCandidateHeaderColumns > 128)
                throw new ArgumentOutOfRangeException(nameof(options.MaxCandidateHeaderColumns));
            ValidateUnit(options.StrongScoreThreshold, nameof(options.StrongScoreThreshold));
            ValidateUnit(options.StrongRunnerUpMargin, nameof(options.StrongRunnerUpMargin));
            ValidateUnit(options.MinimumCoverage, nameof(options.MinimumCoverage));
            ValidateUnit(options.MinimumUniqueness, nameof(options.MinimumUniqueness));
            ValidateUnit(options.MinimumCrossRevisionOverlap, nameof(options.MinimumCrossRevisionOverlap));
        }

        private static void ValidateUnit(double value, string name)
        {
            if (value < 0D || value > 1D) throw new ArgumentOutOfRangeException(name);
        }

        private static AlignmentReason Metric(string code, string label, double value, double weight) =>
            new AlignmentReason(
                code,
                label + ": " + value.ToString("0.000", CultureInfo.InvariantCulture) + ".",
                value * weight);

        private static double Clamp(double value) => Math.Max(0D, Math.Min(1D, value));

        private sealed class ColumnPair
        {
            internal ColumnPair(int before, int after, string normalizedHeader)
            {
                Before = before;
                After = after;
                NormalizedHeader = normalizedHeader;
            }

            internal int Before { get; }
            internal int After { get; }
            internal string NormalizedHeader { get; }
        }

        private sealed class KeyStats
        {
            internal KeyStats(
                double coverage,
                double uniqueness,
                int duplicateCount,
                HashSet<string> keys)
            {
                Coverage = coverage;
                Uniqueness = uniqueness;
                DuplicateCount = duplicateCount;
                Keys = keys;
            }

            internal double Coverage { get; }
            internal double Uniqueness { get; }
            internal int DuplicateCount { get; }
            internal HashSet<string> Keys { get; }
        }

        private sealed class TypeStats
        {
            internal TypeStats(string majorityKind, double majorityRatio)
            {
                MajorityKind = majorityKind;
                MajorityRatio = majorityRatio;
            }

            internal string MajorityKind { get; }
            internal double MajorityRatio { get; }
        }
    }
}
