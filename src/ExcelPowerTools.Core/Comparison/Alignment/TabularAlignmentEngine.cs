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
            return AlignExactTokens(
                beforeTokens,
                afterTokens,
                AxisAlignmentStrategy.UniqueContentAnchors,
                AlignmentConfidence.Strong,
                "column.exact-content-anchor",
                "A unique exact column-content fingerprint matched across revisions.",
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
                        AlignmentConfidence.Strong,
                        AxisAlignmentStrategy.UniqueContentAnchors,
                        new[]
                        {
                            new AlignmentReason(
                                "row.unique-content-anchor",
                                "A row fingerprint occurs exactly once in each revision."),
                        }));
                    consumedBefore.Add(beforeIndices[0]);
                    consumedAfter.Add(afterIndices[0]);
                }
                else if (beforeIndices.Count > 0 && afterIndices.Count > 0 &&
                         (beforeIndices.Count > 1 || afterIndices.Count > 1))
                {
                    ambiguities.Add(new AxisAlignmentAmbiguity(
                        beforeIndices,
                        afterIndices,
                        EvidenceToken(token),
                        new[]
                        {
                            new AlignmentReason(
                                "row.duplicate-content-ambiguity",
                                "The same exact row content occurs multiple times; occurrence identity is unresolved."),
                        }));
                    foreach (int index in beforeIndices) consumedBefore.Add(index);
                    foreach (int index in afterIndices) consumedAfter.Add(index);
                }
            }

            // Unique exact pairs may include moves. Only a monotonic subset is safe as two-sided sequence bounds.
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
            for (int column = 0; column < before.ColumnCount; column++)
            {
                if ((column & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                before.TryGetCell(beforeRow, column, out string value);
                string normalized = AlignmentText.NormalizeLabel(value);
                beforeTokens.Add(column, normalized.Length == 0 ? "\u0001blank-before:" + column : normalized);
            }
            for (int column = 0; column < after.ColumnCount; column++)
            {
                if ((column & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                after.TryGetCell(afterRow, column, out string value);
                string normalized = AlignmentText.NormalizeLabel(value);
                afterTokens.Add(column, normalized.Length == 0 ? "\u0001blank-after:" + column : normalized);
            }

            return AlignExactTokens(
                beforeTokens,
                afterTokens,
                AxisAlignmentStrategy.HeaderLabels,
                AlignmentConfidence.Strong,
                "column.unique-header-label",
                "A normalized header label occurs exactly once in each revision.",
                cancellationToken);
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
