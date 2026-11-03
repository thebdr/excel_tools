using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    public sealed class TabularAlignmentOptions
    {
        public int MaxNeighborBoundedGap { get; set; } = 16;
        public bool AllowEqualShapePositionalFallback { get; set; } = true;
        public double MinimumPositionalSimilarity { get; set; } = 0.80D;
    }

    /// <summary>
    /// Conservative physical row/column mapping. Exact unique keys or content anchors are resolved; duplicate
    /// evidence becomes an explicit ambiguity block. Changed rows may be paired positionally only inside two
    /// monotonic exact anchors, or as a weak equal-shape fallback when every row is highly similar.
    /// </summary>
    public static class TabularAlignmentEngine
    {
        // Exact sequence alignment is deliberately bounded. Useful ranges can contain up to 100,000 rows, so an
        // unconstrained quadratic LCS would be unsafe inside Excel. Gaps above this budget remain explicit
        // ambiguities instead of being guessed.
        private const long MaxExactSequenceMatrixCells = 8000000L;
        private const double MinimumBlankHeaderColumnSimilarity = 0.90D;
        private const double MinimumBlankHeaderCandidateMargin = 0.05D;

        public static AxisAlignmentResult AlignRows(
            AlignmentTable before,
            AlignmentTable after,
            HeaderDetectionResult beforeHeader = null,
            HeaderDetectionResult afterHeader = null,
            KeyInferenceResult keyInference = null,
            TabularAlignmentOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (after == null) throw new ArgumentNullException(nameof(after));
            options = options ?? new TabularAlignmentOptions();
            ValidateOptions(options);
            cancellationToken.ThrowIfCancellationRequested();

            int beforeStart = beforeHeader?.DataStartRowIndex ?? 0;
            int afterStart = afterHeader?.DataStartRowIndex ?? 0;
            if (keyInference != null && keyInference.HasUsableKey)
                return AlignRowsByKey(
                    before,
                    after,
                    beforeStart,
                    afterStart,
                    keyInference,
                    cancellationToken);

            return AlignRowsBySequence(
                before,
                after,
                beforeStart,
                afterStart,
                options,
                cancellationToken);
        }

        public static AxisAlignmentResult AlignColumns(
            AlignmentTable before,
            AlignmentTable after,
            HeaderDetectionResult beforeHeader = null,
            HeaderDetectionResult afterHeader = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (after == null) throw new ArgumentNullException(nameof(after));
            cancellationToken.ThrowIfCancellationRequested();

            if (beforeHeader != null && afterHeader != null &&
                beforeHeader.HasAcceptedHeader && afterHeader.HasAcceptedHeader)
                return AlignColumnsByHeader(
                    before,
                    after,
                    beforeHeader,
                    afterHeader,
                    cancellationToken);

            int beforeStart = beforeHeader?.DataStartRowIndex ?? 0;
            int afterStart = afterHeader?.DataStartRowIndex ?? 0;
            var beforeTokens = new Dictionary<int, string>();
            var afterTokens = new Dictionary<int, string>();
            for (int column = 0; column < before.ColumnCount; column++)
            {
                if ((column & 0x1F) == 0) cancellationToken.ThrowIfCancellationRequested();
                beforeTokens.Add(column, AlignmentText.ExactColumnFingerprint(before, column, beforeStart));
            }
            for (int column = 0; column < after.ColumnCount; column++)
            {
                if ((column & 0x1F) == 0) cancellationToken.ThrowIfCancellationRequested();
                afterTokens.Add(column, AlignmentText.ExactColumnFingerprint(after, column, afterStart));
            }
            return AlignSequenceTokens(
                beforeTokens,
                afterTokens,
                AxisAlignmentStrategy.UniqueContentAnchors,
                AlignmentConfidence.Strong,
                "column.exact-content-anchor",
                "A unique exact column-content fingerprint matched across revisions.",
                "axis.duplicate-content-ambiguity",
                "Competing non-monotonic exact column evidence cannot be paired without changing identity.",
                cancellationToken);
        }

        private static AxisAlignmentResult AlignRowsByKey(
            AlignmentTable before,
            AlignmentTable after,
            int beforeStart,
            int afterStart,
            KeyInferenceResult keyInference,
            CancellationToken cancellationToken)
        {
            KeyCandidate key = keyInference.SelectedCandidate;
            var beforeTokens = new Dictionary<int, string>();
            var afterTokens = new Dictionary<int, string>();
            var beforeMissing = new List<int>();
            var afterMissing = new List<int>();
            for (int row = beforeStart; row < before.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (HasCompleteKey(before, row, key.BeforeColumns))
                    beforeTokens.Add(row, AlignmentText.ExactTuple(before, row, key.BeforeColumns));
                else beforeMissing.Add(row);
            }
            for (int row = afterStart; row < after.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (HasCompleteKey(after, row, key.AfterColumns))
                    afterTokens.Add(row, AlignmentText.ExactTuple(after, row, key.AfterColumns));
                else afterMissing.Add(row);
            }

            AxisAlignmentStrategy strategy = keyInference.Decision == KeyInferenceDecision.ConfirmedKey
                ? AxisAlignmentStrategy.ConfirmedKey
                : AxisAlignmentStrategy.InferredKey;
            AlignmentConfidence confidence = keyInference.Decision == KeyInferenceDecision.ConfirmedKey
                ? AlignmentConfidence.Confirmed
                : AlignmentConfidence.Strong;
            AxisAlignmentResult resolved = AlignExactTokens(
                beforeTokens,
                afterTokens,
                strategy,
                confidence,
                "row.exact-key",
                "An exact unique key tuple matched across revisions.",
                cancellationToken);

            var links = new List<AxisAlignmentLink>(resolved.Links);
            foreach (int row in beforeMissing)
                links.Add(OnlyBefore(
                    row,
                    strategy,
                    "row.missing-key",
                    "The row has an absent or blank key component and was not key-matched."));
            foreach (int row in afterMissing)
                links.Add(OnlyAfter(
                    row,
                    strategy,
                    "row.missing-key",
                    "The row has an absent or blank key component and was not key-matched."));
            SortLinks(links);

            var reasons = new List<AlignmentReason>(resolved.Reasons)
            {
                new AlignmentReason(
                    "row.key-strategy",
                    keyInference.Decision == KeyInferenceDecision.ConfirmedKey
                        ? "Rows were aligned by user-confirmed exact key columns."
                        : "Rows were aligned by conservatively inferred exact key columns."),
            };
            return new AxisAlignmentResult(strategy, links, resolved.Ambiguities, reasons);
        }

        private static AxisAlignmentResult AlignRowsBySequence(
            AlignmentTable before,
            AlignmentTable after,
            int beforeStart,
            int afterStart,
            TabularAlignmentOptions options,
            CancellationToken cancellationToken)
        {
            var beforeTokens = new Dictionary<int, string>();
            var afterTokens = new Dictionary<int, string>();
            for (int row = beforeStart; row < before.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                beforeTokens.Add(row, AlignmentText.ExactRowFingerprint(before, row));
            }
            for (int row = afterStart; row < after.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                afterTokens.Add(row, AlignmentText.ExactRowFingerprint(after, row));
            }

            var pairs = new List<IndexPair>();
            var ambiguities = new List<AxisAlignmentAmbiguity>();
            var consumedBefore = new HashSet<int>();
            var consumedAfter = new HashSet<int>();
            bool usedRepeatedSequence = ResolveMonotonicExactTokens(
                beforeTokens,
                afterTokens,
                AxisAlignmentStrategy.UniqueContentAnchors,
                AlignmentConfidence.Strong,
                "row.unique-content-anchor",
                "An exact row fingerprint was paired in deterministic monotonic sequence order.",
                "row.duplicate-content-ambiguity",
                "Competing non-monotonic exact row evidence cannot be paired without changing row identity.",
                pairs,
                ambiguities,
                consumedBefore,
                consumedAfter,
                cancellationToken);

            // Exact pairs may include unique moves. Only their monotonic subset is safe as two-sided sequence bounds.
            IReadOnlyList<IndexPair> monotonicAnchors = LongestMonotonicSubset(pairs, cancellationToken);
            bool usedNeighborBounds = PairBoundedGaps(
                before,
                after,
                beforeStart,
                afterStart,
                monotonicAnchors,
                consumedBefore,
                consumedAfter,
                pairs,
                options.MaxNeighborBoundedGap,
                cancellationToken);

            bool usedPositional = false;
            if (pairs.Count == 0 && ambiguities.Count == 0 && options.AllowEqualShapePositionalFallback)
                usedPositional = PairEqualShapePositionally(
                    before,
                    after,
                    beforeStart,
                    afterStart,
                    consumedBefore,
                    consumedAfter,
                    pairs,
                    options.MinimumPositionalSimilarity,
                    cancellationToken);

            var links = BuildMatchedLinks(pairs, cancellationToken);
            for (int row = beforeStart; row < before.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!consumedBefore.Contains(row))
                    links.Add(OnlyBefore(
                        row,
                        AxisAlignmentStrategy.UniqueContentAnchors,
                        "row.unmatched-before",
                        "No conservative key, unique content, or bounded-neighbor match was available."));
            }
            for (int row = afterStart; row < after.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!consumedAfter.Contains(row))
                    links.Add(OnlyAfter(
                        row,
                        AxisAlignmentStrategy.UniqueContentAnchors,
                        "row.unmatched-after",
                        "No conservative key, unique content, or bounded-neighbor match was available."));
            }
            SortLinks(links);

            AxisAlignmentStrategy resultStrategy = usedNeighborBounds
                ? AxisAlignmentStrategy.NeighborBoundedSequence
                : usedPositional
                    ? AxisAlignmentStrategy.Positional
                    : AxisAlignmentStrategy.UniqueContentAnchors;
            var reasons = new List<AlignmentReason>
            {
                new AlignmentReason(
                    "row.sequence-fallback",
                    "No strong/confirmed key was supplied; exact content anchors and bounded sequence evidence were used."),
            };
            if (usedRepeatedSequence)
                reasons.Add(new AlignmentReason(
                    "row.repeated-exact-sequence",
                    "Repeated exact rows were paired monotonically; surplus occurrences remain explicit insertions or deletions."));
            if (ambiguities.Count > 0)
                reasons.Add(new AlignmentReason(
                    "row.ambiguities-retained",
                    "Duplicate occurrence sets were retained as unresolved ambiguity blocks."));
            return new AxisAlignmentResult(resultStrategy, links, ambiguities, reasons);
        }

        private static AxisAlignmentResult AlignColumnsByHeader(
            AlignmentTable before,
            AlignmentTable after,
            HeaderDetectionResult beforeHeader,
            HeaderDetectionResult afterHeader,
            CancellationToken cancellationToken)
        {
            int beforeRow = beforeHeader.AcceptedHeaderRowIndex.Value;
            int afterRow = afterHeader.AcceptedHeaderRowIndex.Value;
            var beforeTokens = new Dictionary<int, string>();
            var afterTokens = new Dictionary<int, string>();
            var beforeBlankHeaders = new HashSet<int>();
            var afterBlankHeaders = new HashSet<int>();
            for (int column = 0; column < before.ColumnCount; column++)
            {
                if ((column & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                before.TryGetCell(beforeRow, column, out string value);
                string normalized = AlignmentText.NormalizeLabel(value);
                if (normalized.Length == 0) beforeBlankHeaders.Add(column);
                beforeTokens.Add(column, ColumnSequenceToken(
                    before,
                    column,
                    beforeHeader.DataStartRowIndex,
                    normalized));
            }
            for (int column = 0; column < after.ColumnCount; column++)
            {
                if ((column & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                after.TryGetCell(afterRow, column, out string value);
                string normalized = AlignmentText.NormalizeLabel(value);
                if (normalized.Length == 0) afterBlankHeaders.Add(column);
                afterTokens.Add(column, ColumnSequenceToken(
                    after,
                    column,
                    afterHeader.DataStartRowIndex,
                    normalized));
            }

            AxisAlignmentResult exact = AlignSequenceTokens(
                beforeTokens,
                afterTokens,
                AxisAlignmentStrategy.HeaderLabels,
                AlignmentConfidence.Strong,
                "column.unique-header-label",
                "An exact normalized header or blank-header content fingerprint was paired monotonically.",
                "column.duplicate-header-ambiguity",
                "Competing non-monotonic header evidence cannot be paired without changing column identity.",
                cancellationToken);
            exact = PreserveNamedHeaderAmbiguities(
                beforeTokens,
                afterTokens,
                exact,
                cancellationToken);
            return PairSimilarBlankHeaderColumns(
                before,
                after,
                beforeHeader.DataStartRowIndex,
                afterHeader.DataStartRowIndex,
                beforeBlankHeaders,
                afterBlankHeaders,
                exact,
                cancellationToken);
        }

        private static AxisAlignmentResult PreserveNamedHeaderAmbiguities(
            IReadOnlyDictionary<int, string> beforeTokens,
            IReadOnlyDictionary<int, string> afterTokens,
            AxisAlignmentResult source,
            CancellationToken cancellationToken)
        {
            OccurrenceSet occurrences = BuildOccurrences(beforeTokens, afterTokens, cancellationToken);
            List<string> duplicateNamedHeaders = occurrences.AllTokens.Where(token =>
                    token.StartsWith("\u0002header:", StringComparison.Ordinal) &&
                    occurrences.Before[token].Count > 0 &&
                    occurrences.After[token].Count > 0 &&
                    (occurrences.Before[token].Count > 1 || occurrences.After[token].Count > 1))
                .ToList();
            if (duplicateNamedHeaders.Count == 0) return source;

            var beforeBlocked = new HashSet<int>(duplicateNamedHeaders.SelectMany(token => occurrences.Before[token]));
            var afterBlocked = new HashSet<int>(duplicateNamedHeaders.SelectMany(token => occurrences.After[token]));
            var evidenceTokens = new HashSet<string>(
                duplicateNamedHeaders.Select(EvidenceToken),
                StringComparer.Ordinal);
            List<AxisAlignmentLink> links = source.Links.Where(item =>
                    !item.BeforeIndex.HasValue || !beforeBlocked.Contains(item.BeforeIndex.Value))
                .Where(item => !item.AfterIndex.HasValue || !afterBlocked.Contains(item.AfterIndex.Value))
                .ToList();
            List<AxisAlignmentAmbiguity> ambiguities = source.Ambiguities
                .Where(item => !evidenceTokens.Contains(item.EvidenceToken))
                .ToList();
            foreach (string token in duplicateNamedHeaders)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ambiguities.Add(new AxisAlignmentAmbiguity(
                    occurrences.Before[token],
                    occurrences.After[token],
                    EvidenceToken(token),
                    new[]
                    {
                        new AlignmentReason(
                            "column.duplicate-header-ambiguity",
                            "A repeated nonblank header label is not sufficient to identify changed columns by occurrence."),
                    }));
            }
            SortLinks(links);
            var reasons = new List<AlignmentReason>(source.Reasons)
            {
                new AlignmentReason(
                    "column.named-duplicates-retained",
                    "Repeated nonblank header labels remain unresolved unless an exact whole-table path proves identity."),
            };
            return new AxisAlignmentResult(source.Strategy, links, ambiguities, reasons);
        }

        private static string ColumnSequenceToken(
            AlignmentTable table,
            int column,
            int dataStartRow,
            string normalizedHeader)
        {
            if (!string.IsNullOrEmpty(normalizedHeader)) return "\u0002header:" + normalizedHeader;
            string exactContent = AlignmentText.ExactColumnFingerprint(table, column, dataStartRow);
            return "\u0001blank-content:" + StableId.Compute(
                "alignment-blank-header-column-v1",
                exactContent);
        }

        /// <summary>
        /// Resolves exact axis tokens with global one-to-one anchors plus a bounded monotonic sequence pass for
        /// repeated occurrences. Unlike key alignment, surplus repeated content is a structural insertion/deletion;
        /// only competing cross-anchor or non-monotonic evidence remains ambiguous.
        /// </summary>
        private static AxisAlignmentResult AlignSequenceTokens(
            IReadOnlyDictionary<int, string> beforeTokens,
            IReadOnlyDictionary<int, string> afterTokens,
            AxisAlignmentStrategy strategy,
            AlignmentConfidence confidence,
            string matchCode,
            string matchMessage,
            string ambiguityCode,
            string ambiguityMessage,
            CancellationToken cancellationToken)
        {
            var pairs = new List<IndexPair>();
            var ambiguities = new List<AxisAlignmentAmbiguity>();
            var consumedBefore = new HashSet<int>();
            var consumedAfter = new HashSet<int>();
            bool repeated = ResolveMonotonicExactTokens(
                beforeTokens,
                afterTokens,
                strategy,
                confidence,
                matchCode,
                matchMessage,
                ambiguityCode,
                ambiguityMessage,
                pairs,
                ambiguities,
                consumedBefore,
                consumedAfter,
                cancellationToken);

            var links = BuildMatchedLinks(pairs, cancellationToken);
            foreach (int index in beforeTokens.Keys.OrderBy(value => value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!consumedBefore.Contains(index))
                    links.Add(OnlyBefore(index, strategy, "axis.only-before", "No monotonic exact evidence exists in After."));
            }
            foreach (int index in afterTokens.Keys.OrderBy(value => value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!consumedAfter.Contains(index))
                    links.Add(OnlyAfter(index, strategy, "axis.only-after", "No monotonic exact evidence exists in Before."));
            }
            SortLinks(links);

            var reasons = new List<AlignmentReason>
            {
                new AlignmentReason(
                    "axis.sequence-strategy",
                    "Resolved exact evidence uses bounded monotonic sequence alignment; competing mappings remain explicit."),
            };
            if (repeated)
                reasons.Add(new AlignmentReason(
                    "axis.repeated-exact-sequence",
                    "Repeated exact occurrences were paired in monotonic order and surplus occurrences stayed structural."));
            return new AxisAlignmentResult(strategy, links, ambiguities, reasons);
        }

        private static bool ResolveMonotonicExactTokens(
            IReadOnlyDictionary<int, string> beforeTokens,
            IReadOnlyDictionary<int, string> afterTokens,
            AxisAlignmentStrategy strategy,
            AlignmentConfidence confidence,
            string matchCode,
            string matchMessage,
            string ambiguityCode,
            string ambiguityMessage,
            List<IndexPair> pairs,
            List<AxisAlignmentAmbiguity> ambiguities,
            ISet<int> consumedBefore,
            ISet<int> consumedAfter,
            CancellationToken cancellationToken)
        {
            OccurrenceSet occurrences = BuildOccurrences(beforeTokens, afterTokens, cancellationToken);
            var uniquePairs = new List<IndexPair>();
            foreach (string token in occurrences.AllTokens)
            {
                cancellationToken.ThrowIfCancellationRequested();
                List<int> beforeIndices = occurrences.Before[token];
                List<int> afterIndices = occurrences.After[token];
                if (beforeIndices.Count != 1 || afterIndices.Count != 1) continue;
                var pair = new IndexPair(
                    beforeIndices[0],
                    afterIndices[0],
                    confidence,
                    strategy,
                    new[] { new AlignmentReason(matchCode, matchMessage) });
                uniquePairs.Add(pair);
                pairs.Add(pair);
                consumedBefore.Add(pair.Before);
                consumedAfter.Add(pair.After);
            }

            IReadOnlyList<IndexPair> anchors = LongestMonotonicSubset(uniquePairs, cancellationToken);
            bool usedRepeated = false;
            long remainingWork = MaxExactSequenceMatrixCells;
            List<SequenceItem> orderedBefore = OrderedSequenceItems(beforeTokens, cancellationToken);
            List<SequenceItem> orderedAfter = OrderedSequenceItems(afterTokens, cancellationToken);
            int beforeCursor = 0;
            int afterCursor = 0;
            for (int gap = 0; gap <= anchors.Count; gap++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int beforeLast = gap == anchors.Count ? int.MaxValue : anchors[gap].Before - 1;
                int afterLast = gap == anchors.Count ? int.MaxValue : anchors[gap].After - 1;
                var beforeGap = new List<SequenceItem>();
                while (beforeCursor < orderedBefore.Count && orderedBefore[beforeCursor].Index <= beforeLast)
                {
                    SequenceItem item = orderedBefore[beforeCursor++];
                    if (!consumedBefore.Contains(item.Index)) beforeGap.Add(item);
                }
                var afterGap = new List<SequenceItem>();
                while (afterCursor < orderedAfter.Count && orderedAfter[afterCursor].Index <= afterLast)
                {
                    SequenceItem item = orderedAfter[afterCursor++];
                    if (!consumedAfter.Contains(item.Index)) afterGap.Add(item);
                }
                if (beforeGap.Count == 0 || afterGap.Count == 0) continue;

                TryAlignExactSequence(
                    beforeGap,
                    afterGap,
                    ref remainingWork,
                    cancellationToken,
                    out List<SequenceMatch> matches);

                var crossingTokens = new HashSet<string>(StringComparer.Ordinal);
                foreach (SequenceMatch match in matches)
                {
                    if (uniquePairs.Any(unique => Crosses(unique.Before, unique.After, match.Before, match.After)))
                        crossingTokens.Add(match.Token);
                }

                HashSet<string> competing = CompetingSequenceTokens(
                    beforeGap,
                    afterGap,
                    matches,
                    crossingTokens,
                    occurrences,
                    cancellationToken);
                foreach (string token in competing)
                {
                    AddSequenceAmbiguity(
                        token,
                        beforeGap.Where(item => string.Equals(item.Token, token, StringComparison.Ordinal))
                            .Select(item => item.Index),
                        afterGap.Where(item => string.Equals(item.Token, token, StringComparison.Ordinal))
                            .Select(item => item.Index),
                        ambiguityCode,
                        ambiguityMessage,
                        ambiguities,
                        consumedBefore,
                        consumedAfter);
                }

                foreach (SequenceMatch match in matches)
                {
                    if (competing.Contains(match.Token)) continue;
                    pairs.Add(new IndexPair(
                        match.Before,
                        match.After,
                        confidence,
                        strategy,
                        new[] { new AlignmentReason(matchCode, matchMessage) }));
                    consumedBefore.Add(match.Before);
                    consumedAfter.Add(match.After);
                    if (occurrences.Before[match.Token].Count > 1 || occurrences.After[match.Token].Count > 1)
                        usedRepeated = true;
                }

                // Common evidence skipped inside the same anchor-bounded gap represents a competing LCS path,
                // a non-monotonic mapping, or an exhausted work budget.
                foreach (string token in beforeGap.Select(item => item.Token)
                             .Intersect(afterGap.Select(item => item.Token), StringComparer.Ordinal)
                             .Distinct(StringComparer.Ordinal))
                {
                    List<int> beforeRemaining = beforeGap
                        .Where(item => string.Equals(item.Token, token, StringComparison.Ordinal) &&
                                       !consumedBefore.Contains(item.Index))
                        .Select(item => item.Index)
                        .ToList();
                    List<int> afterRemaining = afterGap
                        .Where(item => string.Equals(item.Token, token, StringComparison.Ordinal) &&
                                       !consumedAfter.Contains(item.Index))
                        .Select(item => item.Index)
                        .ToList();
                    if (beforeRemaining.Count == 0 || afterRemaining.Count == 0) continue;
                    AddSequenceAmbiguity(
                        token,
                        beforeRemaining,
                        afterRemaining,
                        ambiguityCode,
                        ambiguityMessage,
                        ambiguities,
                        consumedBefore,
                        consumedAfter);
                }
            }

            // A repeated token that survives on both sides only because its occurrences landed in different
            // unique-anchor gaps crossed a proven anchor. Treating that as an unrelated delete + insert would
            // silently choose an identity for a reordered duplicate group, so retain it as ambiguity.
            foreach (string token in occurrences.AllTokens)
            {
                cancellationToken.ThrowIfCancellationRequested();
                List<int> beforeRemaining = occurrences.Before[token]
                    .Where(index => !consumedBefore.Contains(index))
                    .ToList();
                List<int> afterRemaining = occurrences.After[token]
                    .Where(index => !consumedAfter.Contains(index))
                    .ToList();
                if (beforeRemaining.Count == 0 || afterRemaining.Count == 0) continue;
                AddSequenceAmbiguity(
                    token,
                    beforeRemaining,
                    afterRemaining,
                    ambiguityCode,
                    ambiguityMessage,
                    ambiguities,
                    consumedBefore,
                    consumedAfter);
            }
            return usedRepeated;
        }

        private static List<SequenceItem> OrderedSequenceItems(
            IReadOnlyDictionary<int, string> tokens,
            CancellationToken cancellationToken)
        {
            var result = new List<SequenceItem>();
            foreach (KeyValuePair<int, string> pair in tokens.OrderBy(item => item.Key))
            {
                cancellationToken.ThrowIfCancellationRequested();
                result.Add(new SequenceItem(pair.Key, pair.Value));
            }
            return result;
        }

        private static HashSet<string> CompetingSequenceTokens(
            IReadOnlyList<SequenceItem> before,
            IReadOnlyList<SequenceItem> after,
            IReadOnlyList<SequenceMatch> matches,
            IEnumerable<string> rejectedTokens,
            OccurrenceSet allOccurrences,
            CancellationToken cancellationToken)
        {
            var beforeByToken = before.GroupBy(item => item.Token, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
            var afterByToken = after.GroupBy(item => item.Token, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
            var competing = new HashSet<string>(rejectedTokens ?? Array.Empty<string>(), StringComparer.Ordinal);
            var matchedBefore = new HashSet<int>(matches
                .Where(match => !competing.Contains(match.Token))
                .Select(match => match.Before));
            var matchedAfter = new HashSet<int>(matches
                .Where(match => !competing.Contains(match.Token))
                .Select(match => match.After));
            foreach (string token in beforeByToken.Keys.Intersect(afterByToken.Keys, StringComparer.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (beforeByToken[token].Any(item => !matchedBefore.Contains(item.Index)) &&
                    afterByToken[token].Any(item => !matchedAfter.Contains(item.Index)))
                    competing.Add(token);
            }

            // If an alternate LCS can replace a matched repeated token with unresolved non-identical evidence,
            // both token groups are competing. Expand that relation to a fixed point, but cap the examination so
            // adversarial 100k-row inputs cannot turn ambiguity detection into another unbounded quadratic pass.
            long checks = 0;
            bool changed;
            do
            {
                changed = false;
                string[] unresolved = competing.ToArray();
                foreach (SequenceMatch match in matches)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (competing.Contains(match.Token) ||
                        allOccurrences.Before[match.Token].Count <= 1 &&
                        allOccurrences.After[match.Token].Count <= 1)
                        continue;
                    foreach (string token in unresolved)
                    {
                        if (++checks > MaxExactSequenceMatrixCells)
                        {
                            foreach (SequenceMatch remaining in matches)
                                if (allOccurrences.Before[remaining.Token].Count > 1 ||
                                    allOccurrences.After[remaining.Token].Count > 1)
                                    competing.Add(remaining.Token);
                            return competing;
                        }
                        if (!beforeByToken.TryGetValue(token, out List<SequenceItem> beforeItems) ||
                            !afterByToken.TryGetValue(token, out List<SequenceItem> afterItems))
                            continue;
                        int beforeMinimum = beforeItems[0].Index;
                        int beforeMaximum = beforeItems[beforeItems.Count - 1].Index;
                        int afterMinimum = afterItems[0].Index;
                        int afterMaximum = afterItems[afterItems.Count - 1].Index;
                        if (beforeMinimum < match.Before && afterMaximum > match.After ||
                            beforeMaximum > match.Before && afterMinimum < match.After)
                        {
                            competing.Add(match.Token);
                            changed = true;
                            break;
                        }
                    }
                }
            } while (changed);
            return competing;
        }

        private static bool TryAlignExactSequence(
            IReadOnlyList<SequenceItem> before,
            IReadOnlyList<SequenceItem> after,
            ref long remainingWork,
            CancellationToken cancellationToken,
            out List<SequenceMatch> matches)
        {
            matches = new List<SequenceMatch>();
            int prefix = 0;
            while (prefix < before.Count && prefix < after.Count &&
                   string.Equals(before[prefix].Token, after[prefix].Token, StringComparison.Ordinal))
            {
                if ((prefix & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                matches.Add(new SequenceMatch(
                    before[prefix].Index,
                    after[prefix].Index,
                    before[prefix].Token));
                prefix++;
            }

            int beforeEnd = before.Count - 1;
            int afterEnd = after.Count - 1;
            var suffix = new List<SequenceMatch>();
            while (beforeEnd >= prefix && afterEnd >= prefix &&
                   string.Equals(before[beforeEnd].Token, after[afterEnd].Token, StringComparison.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                suffix.Add(new SequenceMatch(
                    before[beforeEnd].Index,
                    after[afterEnd].Index,
                    before[beforeEnd].Token));
                beforeEnd--;
                afterEnd--;
            }

            int beforeLength = beforeEnd - prefix + 1;
            int afterLength = afterEnd - prefix + 1;
            if (beforeLength > 0 && afterLength > 0)
            {
                long cells = checked((long)(beforeLength + 1) * (afterLength + 1));
                if (cells > remainingWork || cells > int.MaxValue)
                {
                    suffix.Reverse();
                    matches.AddRange(suffix);
                    return false;
                }
                remainingWork -= cells;
                int stride = afterLength + 1;
                var lengths = new int[(int)cells];
                for (int left = beforeLength - 1; left >= 0; left--)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int row = left * stride;
                    int next = (left + 1) * stride;
                    for (int right = afterLength - 1; right >= 0; right--)
                    {
                        if ((right & 0xFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                        if (string.Equals(
                                before[prefix + left].Token,
                                after[prefix + right].Token,
                                StringComparison.Ordinal))
                            lengths[row + right] = lengths[next + right + 1] + 1;
                        else
                            lengths[row + right] = Math.Max(lengths[next + right], lengths[row + right + 1]);
                    }
                }

                int beforeCursor = 0;
                int afterCursor = 0;
                while (beforeCursor < beforeLength && afterCursor < afterLength)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    SequenceItem left = before[prefix + beforeCursor];
                    SequenceItem right = after[prefix + afterCursor];
                    if (string.Equals(left.Token, right.Token, StringComparison.Ordinal))
                    {
                        matches.Add(new SequenceMatch(left.Index, right.Index, left.Token));
                        beforeCursor++;
                        afterCursor++;
                        continue;
                    }
                    int down = lengths[(beforeCursor + 1) * stride + afterCursor];
                    int across = lengths[beforeCursor * stride + afterCursor + 1];
                    if (down > across) beforeCursor++;
                    else afterCursor++;
                }
            }

            suffix.Reverse();
            matches.AddRange(suffix);
            return true;
        }

        private static bool Crosses(int beforeA, int afterA, int beforeB, int afterB) =>
            beforeA < beforeB && afterA > afterB || beforeA > beforeB && afterA < afterB;

        private static void AddSequenceAmbiguity(
            string token,
            IEnumerable<int> beforeIndices,
            IEnumerable<int> afterIndices,
            string code,
            string message,
            ICollection<AxisAlignmentAmbiguity> ambiguities,
            ISet<int> consumedBefore,
            ISet<int> consumedAfter)
        {
            List<int> before = beforeIndices.Distinct().OrderBy(index => index).ToList();
            List<int> after = afterIndices.Distinct().OrderBy(index => index).ToList();
            if (before.Count == 0 || after.Count == 0) return;
            ambiguities.Add(new AxisAlignmentAmbiguity(
                before,
                after,
                EvidenceToken(token),
                new[] { new AlignmentReason(code, message) }));
            foreach (int index in before) consumedBefore.Add(index);
            foreach (int index in after) consumedAfter.Add(index);
        }

        private static AxisAlignmentResult AlignExactTokens(
            IReadOnlyDictionary<int, string> beforeTokens,
            IReadOnlyDictionary<int, string> afterTokens,
            AxisAlignmentStrategy strategy,
            AlignmentConfidence confidence,
            string matchCode,
            string matchMessage,
            CancellationToken cancellationToken)
        {
            OccurrenceSet occurrences = BuildOccurrences(beforeTokens, afterTokens, cancellationToken);
            var pairs = new List<IndexPair>();
            var ambiguities = new List<AxisAlignmentAmbiguity>();
            var consumedBefore = new HashSet<int>();
            var consumedAfter = new HashSet<int>();

            foreach (string token in occurrences.AllTokens)
            {
                cancellationToken.ThrowIfCancellationRequested();
                List<int> beforeIndices = occurrences.Before[token];
                List<int> afterIndices = occurrences.After[token];
                if (beforeIndices.Count == 1 && afterIndices.Count == 1)
                {
                    pairs.Add(new IndexPair(
                        beforeIndices[0],
                        afterIndices[0],
                        confidence,
                        strategy,
                        new[] { new AlignmentReason(matchCode, matchMessage) }));
                    consumedBefore.Add(beforeIndices[0]);
                    consumedAfter.Add(afterIndices[0]);
                }
                else if (beforeIndices.Count > 0 && afterIndices.Count > 0)
                {
                    ambiguities.Add(new AxisAlignmentAmbiguity(
                        beforeIndices,
                        afterIndices,
                        EvidenceToken(token),
                        new[]
                        {
                            new AlignmentReason(
                                strategy == AxisAlignmentStrategy.HeaderLabels
                                    ? "column.duplicate-header-ambiguity"
                                    : strategy == AxisAlignmentStrategy.InferredKey ||
                                      strategy == AxisAlignmentStrategy.ConfirmedKey
                                        ? "row.duplicate-key-ambiguity"
                                        : "axis.duplicate-content-ambiguity",
                                "The same alignment evidence occurs more than once; no occurrence pairing was guessed."),
                        }));
                    foreach (int index in beforeIndices) consumedBefore.Add(index);
                    foreach (int index in afterIndices) consumedAfter.Add(index);
                }
            }

            var links = BuildMatchedLinks(pairs, cancellationToken);
            foreach (int index in beforeTokens.Keys.OrderBy(value => value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!consumedBefore.Contains(index))
                    links.Add(OnlyBefore(index, strategy, "axis.only-before", "No matching evidence exists in After."));
            }
            foreach (int index in afterTokens.Keys.OrderBy(value => value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!consumedAfter.Contains(index))
                    links.Add(OnlyAfter(index, strategy, "axis.only-after", "No matching evidence exists in Before."));
            }
            SortLinks(links);

            return new AxisAlignmentResult(
                strategy,
                links,
                ambiguities,
                new[]
                {
                    new AlignmentReason(
                        "axis.strategy",
                        "Resolved mappings use " + strategy + "; duplicate evidence remains explicit."),
                });
        }

        private static AxisAlignmentResult PairSimilarBlankHeaderColumns(
            AlignmentTable before,
            AlignmentTable after,
            int beforeDataStart,
            int afterDataStart,
            ISet<int> beforeBlankHeaders,
            ISet<int> afterBlankHeaders,
            AxisAlignmentResult source,
            CancellationToken cancellationToken)
        {
            var links = new List<AxisAlignmentLink>(source.Links);
            var ambiguousBefore = new HashSet<int>(source.Ambiguities.SelectMany(item => item.BeforeIndices));
            var ambiguousAfter = new HashSet<int>(source.Ambiguities.SelectMany(item => item.AfterIndices));
            var matchedBefore = new HashSet<int>(links
                .Where(item => item.Kind == AxisAlignmentKind.Matched)
                .Select(item => item.BeforeIndex.Value));
            var matchedAfter = new HashSet<int>(links
                .Where(item => item.Kind == AxisAlignmentKind.Matched)
                .Select(item => item.AfterIndex.Value));
            List<int> unresolvedBefore = beforeBlankHeaders
                .Where(index => !matchedBefore.Contains(index) && !ambiguousBefore.Contains(index))
                .OrderBy(index => index)
                .ToList();
            var unresolvedAfter = new HashSet<int>(afterBlankHeaders
                .Where(index => !matchedAfter.Contains(index) && !ambiguousAfter.Contains(index)));
            if (unresolvedBefore.Count == 0 || unresolvedAfter.Count == 0) return source;

            var exactPairs = links
                .Where(item => item.Kind == AxisAlignmentKind.Matched)
                .Select(item => new IndexPair(
                    item.BeforeIndex.Value,
                    item.AfterIndex.Value,
                    item.Confidence,
                    item.Strategy,
                    item.Reasons))
                .ToList();
            IReadOnlyList<IndexPair> anchors = LongestMonotonicSubset(exactPairs, cancellationToken);
            var proposals = new List<BlankColumnCandidate>();
            foreach (int beforeColumn in unresolvedBefore)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int interval = AnchorInterval(beforeColumn, anchors, true);
                int leftBefore = interval == 0 ? -1 : anchors[interval - 1].Before;
                int leftAfter = interval == 0 ? -1 : anchors[interval - 1].After;
                int rightBefore = interval == anchors.Count ? before.ColumnCount : anchors[interval].Before;
                int rightAfter = interval == anchors.Count ? after.ColumnCount : anchors[interval].After;
                var positionalCandidates = new HashSet<int>
                {
                    leftAfter + (beforeColumn - leftBefore),
                    rightAfter - (rightBefore - beforeColumn),
                };
                var viable = new List<BlankColumnCandidate>();
                foreach (int afterColumn in positionalCandidates.OrderBy(index => index))
                {
                    if (!unresolvedAfter.Contains(afterColumn) || afterColumn <= leftAfter || afterColumn >= rightAfter)
                        continue;
                    double similarity = ColumnSimilarity(
                        before,
                        after,
                        beforeColumn,
                        afterColumn,
                        beforeDataStart,
                        afterDataStart,
                        cancellationToken,
                        out int beforeEvidence,
                        out int afterEvidence);
                    if (beforeEvidence == 0 || afterEvidence == 0 ||
                        similarity < MinimumBlankHeaderColumnSimilarity)
                        continue;
                    viable.Add(new BlankColumnCandidate(beforeColumn, afterColumn, similarity));
                }
                viable.Sort((left, right) =>
                {
                    int similarity = right.Similarity.CompareTo(left.Similarity);
                    return similarity != 0 ? similarity : left.After.CompareTo(right.After);
                });
                if (viable.Count == 0) continue;
                if (viable.Count > 1 &&
                    viable[0].Similarity - viable[1].Similarity < MinimumBlankHeaderCandidateMargin)
                    continue;
                proposals.Add(viable[0]);
            }

            HashSet<int> competingTargets = new HashSet<int>(proposals
                .GroupBy(item => item.After)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key));
            List<BlankColumnCandidate> accepted = proposals
                .Where(item => !competingTargets.Contains(item.After))
                .OrderBy(item => item.Before)
                .ToList();
            int previousAfter = -1;
            bool changed = false;
            foreach (BlankColumnCandidate candidate in accepted)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (candidate.After <= previousAfter || exactPairs.Any(pair => Crosses(
                        pair.Before,
                        pair.After,
                        candidate.Before,
                        candidate.After)))
                    continue;
                previousAfter = candidate.After;
                links.RemoveAll(item =>
                    item.Kind == AxisAlignmentKind.OnlyBefore && item.BeforeIndex == candidate.Before ||
                    item.Kind == AxisAlignmentKind.OnlyAfter && item.AfterIndex == candidate.After);
                links.Add(new AxisAlignmentLink(
                    candidate.Before,
                    candidate.After,
                    AxisAlignmentKind.Matched,
                    AlignmentConfidence.Medium,
                    AxisAlignmentStrategy.NeighborBoundedSequence,
                    new[]
                    {
                        new AlignmentReason(
                            "column.blank-header-neighbor-similarity",
                            "Blank-header columns occupy the same monotonic neighbor-bounded position and have " +
                            "at least 90% exact same-row cell similarity."),
                    }));
                changed = true;
            }
            if (!changed) return source;

            SortLinks(links);
            var reasons = new List<AlignmentReason>(source.Reasons)
            {
                new AlignmentReason(
                    "column.blank-header-neighbor-bounded",
                    "A uniquely positioned high-similarity blank-header column was aligned with Medium confidence."),
            };
            return new AxisAlignmentResult(source.Strategy, links, source.Ambiguities, reasons);
        }

        private static int AnchorInterval(
            int index,
            IReadOnlyList<IndexPair> anchors,
            bool before)
        {
            int low = 0;
            int high = anchors.Count;
            while (low < high)
            {
                int middle = low + (high - low) / 2;
                int value = before ? anchors[middle].Before : anchors[middle].After;
                if (value < index) low = middle + 1;
                else high = middle;
            }
            return low;
        }

        private static double ColumnSimilarity(
            AlignmentTable before,
            AlignmentTable after,
            int beforeColumn,
            int afterColumn,
            int beforeStart,
            int afterStart,
            CancellationToken cancellationToken,
            out int beforeEvidence,
            out int afterEvidence)
        {
            int beforeHeight = Math.Max(0, before.RowCount - beforeStart);
            int afterHeight = Math.Max(0, after.RowCount - afterStart);
            int height = Math.Max(beforeHeight, afterHeight);
            int equal = 0;
            beforeEvidence = 0;
            afterEvidence = 0;
            for (int offset = 0; offset < height; offset++)
            {
                if ((offset & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                string beforeValue = null;
                string afterValue = null;
                bool beforePresent = offset < beforeHeight &&
                                     before.TryGetCell(beforeStart + offset, beforeColumn, out beforeValue);
                bool afterPresent = offset < afterHeight &&
                                    after.TryGetCell(afterStart + offset, afterColumn, out afterValue);
                if (beforePresent && !AlignmentText.IsBlank(beforeValue)) beforeEvidence++;
                if (afterPresent && !AlignmentText.IsBlank(afterValue)) afterEvidence++;
                if (beforePresent == afterPresent &&
                    string.Equals(beforeValue, afterValue, StringComparison.Ordinal))
                    equal++;
            }
            return height == 0 ? 1D : (double)equal / height;
        }

        private static bool PairBoundedGaps(
            AlignmentTable before,
            AlignmentTable after,
            int beforeStart,
            int afterStart,
            IReadOnlyList<IndexPair> anchors,
            ISet<int> consumedBefore,
            ISet<int> consumedAfter,
            ICollection<IndexPair> pairs,
            int maxGap,
            CancellationToken cancellationToken)
        {
            if (anchors.Count < 2) return false;
            bool used = false;
            for (int anchor = 0; anchor < anchors.Count - 1; anchor++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                IndexPair left = anchors[anchor];
                IndexPair right = anchors[anchor + 1];
                var beforeGap = new List<int>();
                var afterGap = new List<int>();
                for (int row = Math.Max(beforeStart, left.Before + 1); row < right.Before; row++)
                {
                    if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (!consumedBefore.Contains(row)) beforeGap.Add(row);
                }
                for (int row = Math.Max(afterStart, left.After + 1); row < right.After; row++)
                {
                    if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (!consumedAfter.Contains(row)) afterGap.Add(row);
                }
                if (beforeGap.Count == 0 || beforeGap.Count != afterGap.Count || beforeGap.Count > maxGap) continue;

                bool safe = beforeGap.Count == 1;
                if (!safe)
                {
                    safe = true;
                    for (int index = 0; index < beforeGap.Count; index++)
                    {
                        if (RowSimilarity(
                                before,
                                after,
                                beforeGap[index],
                                afterGap[index],
                                cancellationToken) < 0.50D)
                        {
                            safe = false;
                            break;
                        }
                    }
                }
                if (!safe) continue;

                for (int index = 0; index < beforeGap.Count; index++)
                {
                    pairs.Add(new IndexPair(
                        beforeGap[index],
                        afterGap[index],
                        AlignmentConfidence.Medium,
                        AxisAlignmentStrategy.NeighborBoundedSequence,
                        new[]
                        {
                            new AlignmentReason(
                                "row.two-sided-neighbor-bounds",
                                "The changed row position is bounded by monotonic unique exact rows on both sides."),
                        }));
                    consumedBefore.Add(beforeGap[index]);
                    consumedAfter.Add(afterGap[index]);
                }
                used = true;
            }
            return used;
        }

        private static bool PairEqualShapePositionally(
            AlignmentTable before,
            AlignmentTable after,
            int beforeStart,
            int afterStart,
            ISet<int> consumedBefore,
            ISet<int> consumedAfter,
            ICollection<IndexPair> pairs,
            double minimumSimilarity,
            CancellationToken cancellationToken)
        {
            int beforeCount = before.RowCount - beforeStart;
            int afterCount = after.RowCount - afterStart;
            if (beforeCount <= 0 || beforeCount != afterCount) return false;
            for (int offset = 0; offset < beforeCount; offset++)
                if (RowSimilarity(
                        before,
                        after,
                        beforeStart + offset,
                        afterStart + offset,
                        cancellationToken) < minimumSimilarity)
                    return false;

            for (int offset = 0; offset < beforeCount; offset++)
            {
                if ((offset & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                int beforeRow = beforeStart + offset;
                int afterRow = afterStart + offset;
                pairs.Add(new IndexPair(
                    beforeRow,
                    afterRow,
                    AlignmentConfidence.Weak,
                    AxisAlignmentStrategy.Positional,
                    new[]
                    {
                        new AlignmentReason(
                            "row.equal-shape-positional",
                            "Equal row counts and high same-position exact-cell similarity support a weak positional mapping."),
                    }));
                consumedBefore.Add(beforeRow);
                consumedAfter.Add(afterRow);
            }
            return true;
        }

        private static double RowSimilarity(
            AlignmentTable before,
            AlignmentTable after,
            int beforeRow,
            int afterRow,
            CancellationToken cancellationToken)
        {
            int width = Math.Max(before.Rows[beforeRow].Count, after.Rows[afterRow].Count);
            if (width == 0) return 1D;
            int equal = 0;
            for (int column = 0; column < width; column++)
            {
                if ((column & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                bool beforePresent = before.TryGetCell(beforeRow, column, out string beforeValue);
                bool afterPresent = after.TryGetCell(afterRow, column, out string afterValue);
                if (beforePresent == afterPresent && string.Equals(beforeValue, afterValue, StringComparison.Ordinal)) equal++;
            }
            return (double)equal / width;
        }

        private static IReadOnlyList<IndexPair> LongestMonotonicSubset(
            IEnumerable<IndexPair> pairs,
            CancellationToken cancellationToken)
        {
            List<IndexPair> ordered = pairs.OrderBy(pair => pair.Before).ThenBy(pair => pair.After).ToList();
            if (ordered.Count <= 1) return new ReadOnlyCollection<IndexPair>(ordered);

            var tails = new int[ordered.Count];
            var tailIndices = new int[ordered.Count];
            var previous = new int[ordered.Count];
            int length = 0;
            for (int index = 0; index < ordered.Count; index++)
            {
                if ((index & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                int value = ordered[index].After;
                int low = 0;
                int high = length;
                while (low < high)
                {
                    int middle = low + (high - low) / 2;
                    if (tails[middle] < value) low = middle + 1;
                    else high = middle;
                }
                tails[low] = value;
                tailIndices[low] = index;
                previous[index] = low > 0 ? tailIndices[low - 1] : -1;
                if (low == length) length++;
            }

            var result = new IndexPair[length];
            int cursor = tailIndices[length - 1];
            for (int index = length - 1; index >= 0; index--)
            {
                result[index] = ordered[cursor];
                cursor = previous[cursor];
            }
            return new ReadOnlyCollection<IndexPair>(result);
        }

        private static List<AxisAlignmentLink> BuildMatchedLinks(
            IReadOnlyList<IndexPair> pairs,
            CancellationToken cancellationToken)
        {
            List<IndexPair> ordered = pairs.OrderBy(pair => pair.Before).ThenBy(pair => pair.After).ToList();
            int count = ordered.Count;
            var prefixMax = new int[count];
            var suffixMin = new int[count];
            for (int index = 0; index < count; index++)
            {
                if ((index & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                prefixMax[index] = index == 0 ? ordered[index].After : Math.Max(prefixMax[index - 1], ordered[index].After);
            }
            for (int index = count - 1; index >= 0; index--)
            {
                if ((index & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                suffixMin[index] = index == count - 1
                    ? ordered[index].After
                    : Math.Min(suffixMin[index + 1], ordered[index].After);
            }

            var links = new List<AxisAlignmentLink>(count);
            for (int index = 0; index < count; index++)
            {
                if ((index & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                IndexPair pair = ordered[index];
                bool reordered = index > 0 && pair.After < prefixMax[index - 1] ||
                                 index < count - 1 && pair.After > suffixMin[index + 1];
                links.Add(new AxisAlignmentLink(
                    pair.Before,
                    pair.After,
                    AxisAlignmentKind.Matched,
                    pair.Confidence,
                    pair.Strategy,
                    pair.Reasons,
                    reordered));
            }
            return links;
        }

        private static OccurrenceSet BuildOccurrences(
            IReadOnlyDictionary<int, string> beforeTokens,
            IReadOnlyDictionary<int, string> afterTokens,
            CancellationToken cancellationToken)
        {
            var before = ToOccurrences(beforeTokens, cancellationToken);
            var after = ToOccurrences(afterTokens, cancellationToken);
            var tokens = new SortedSet<string>(before.Keys, StringComparer.Ordinal);
            tokens.UnionWith(after.Keys);
            foreach (string token in tokens)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!before.ContainsKey(token)) before.Add(token, new List<int>());
                if (!after.ContainsKey(token)) after.Add(token, new List<int>());
            }
            return new OccurrenceSet(before, after, new ReadOnlyCollection<string>(tokens.ToList()));
        }

        private static Dictionary<string, List<int>> ToOccurrences(
            IReadOnlyDictionary<int, string> tokens,
            CancellationToken cancellationToken)
        {
            var result = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            foreach (KeyValuePair<int, string> pair in tokens)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string token = pair.Value ?? "";
                if (!result.TryGetValue(token, out List<int> indices))
                {
                    indices = new List<int>();
                    result.Add(token, indices);
                }
                indices.Add(pair.Key);
            }
            foreach (List<int> indices in result.Values) indices.Sort();
            return result;
        }

        private static bool HasCompleteKey(AlignmentTable table, int row, IReadOnlyList<int> columns)
        {
            foreach (int column in columns)
                if (!table.TryGetCell(row, column, out string value) || AlignmentText.IsBlank(value)) return false;
            return true;
        }

        private static AxisAlignmentLink OnlyBefore(
            int index,
            AxisAlignmentStrategy strategy,
            string code,
            string message) =>
            new AxisAlignmentLink(
                index,
                null,
                AxisAlignmentKind.OnlyBefore,
                AlignmentConfidence.None,
                strategy,
                new[] { new AlignmentReason(code, message) });

        private static AxisAlignmentLink OnlyAfter(
            int index,
            AxisAlignmentStrategy strategy,
            string code,
            string message) =>
            new AxisAlignmentLink(
                null,
                index,
                AxisAlignmentKind.OnlyAfter,
                AlignmentConfidence.None,
                strategy,
                new[] { new AlignmentReason(code, message) });

        private static void SortLinks(List<AxisAlignmentLink> links)
        {
            links.Sort((left, right) =>
            {
                int leftIndex = left.BeforeIndex ?? int.MaxValue;
                int rightIndex = right.BeforeIndex ?? int.MaxValue;
                int before = leftIndex.CompareTo(rightIndex);
                if (before != 0) return before;
                return (left.AfterIndex ?? int.MaxValue).CompareTo(right.AfterIndex ?? int.MaxValue);
            });
        }

        private static string EvidenceToken(string token) =>
            StableId.Compute("alignment-evidence-token-v1", token ?? "");

        private static void ValidateOptions(TabularAlignmentOptions options)
        {
            if (options.MaxNeighborBoundedGap <= 0 || options.MaxNeighborBoundedGap > 10000)
                throw new ArgumentOutOfRangeException(nameof(options.MaxNeighborBoundedGap));
            if (options.MinimumPositionalSimilarity < 0D || options.MinimumPositionalSimilarity > 1D)
                throw new ArgumentOutOfRangeException(nameof(options.MinimumPositionalSimilarity));
        }

        private sealed class SequenceItem
        {
            internal SequenceItem(int index, string token)
            {
                Index = index;
                Token = token ?? "";
            }

            internal int Index { get; }
            internal string Token { get; }
        }

        private sealed class SequenceMatch
        {
            internal SequenceMatch(int before, int after, string token)
            {
                Before = before;
                After = after;
                Token = token ?? "";
            }

            internal int Before { get; }
            internal int After { get; }
            internal string Token { get; }
        }

        private sealed class BlankColumnCandidate
        {
            internal BlankColumnCandidate(int before, int after, double similarity)
            {
                Before = before;
                After = after;
                Similarity = similarity;
            }

            internal int Before { get; }
            internal int After { get; }
            internal double Similarity { get; }
        }

        private sealed class IndexPair
        {
            internal IndexPair(
                int before,
                int after,
                AlignmentConfidence confidence,
                AxisAlignmentStrategy strategy,
                IEnumerable<AlignmentReason> reasons)
            {
                Before = before;
                After = after;
                Confidence = confidence;
                Strategy = strategy;
                Reasons = new ReadOnlyCollection<AlignmentReason>(
                    new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
            }

            internal int Before { get; }
            internal int After { get; }
            internal AlignmentConfidence Confidence { get; }
            internal AxisAlignmentStrategy Strategy { get; }
            internal IReadOnlyList<AlignmentReason> Reasons { get; }
        }

        private sealed class OccurrenceSet
        {
            internal OccurrenceSet(
                Dictionary<string, List<int>> before,
                Dictionary<string, List<int>> after,
                IReadOnlyList<string> allTokens)
            {
                Before = before;
                After = after;
                AllTokens = allTokens;
            }

            internal Dictionary<string, List<int>> Before { get; }
            internal Dictionary<string, List<int>> After { get; }
            internal IReadOnlyList<string> AllTokens { get; }
        }
    }
}
