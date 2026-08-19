using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    public enum HeaderDecision
    {
        Data,
        Ambiguous,
        StrongHeader,
        ConfirmedHeader,
        ConfirmedNoHeader,
    }

    public sealed class HeaderDetectionOptions
    {
        public int MaxScanRows { get; set; } = 25;
        public int MaxEvidenceRows { get; set; } = 8;
        public string CommentPrefix { get; set; }
        public int? ConfirmedHeaderRowIndex { get; set; }
        public bool ConfirmedNoHeader { get; set; }
        /// <summary>
        /// Optional zero-based first keyed-data row for an explicitly confirmed header. This permits retained
        /// subheader/unit rows between the header and data. It may equal RowCount when no stored data follows.
        /// </summary>
        public int? ConfirmedDataStartRowIndex { get; set; }
        public double StrongScoreThreshold { get; set; } = 0.70D;
        public double StrongRunnerUpMargin { get; set; } = 0.12D;
        public double AmbiguousScoreThreshold { get; set; } = 0.48D;
    }

    public sealed class HeaderDetectionResult
    {
        internal HeaderDetectionResult(
            int? candidateRowIndex,
            int? acceptedHeaderRowIndex,
            HeaderDecision decision,
            double confidence,
            double runnerUpMargin,
            IEnumerable<AlignmentReason> evidence,
            int? confirmedDataStartRowIndex = null)
        {
            CandidateRowIndex = candidateRowIndex;
            AcceptedHeaderRowIndex = acceptedHeaderRowIndex;
            Decision = decision;
            Confidence = confidence;
            RunnerUpMargin = runnerUpMargin;
            Evidence = new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(evidence ?? Array.Empty<AlignmentReason>()));
            ConfirmedDataStartRowIndex = confirmedDataStartRowIndex;
        }

        /// <summary>The best raw candidate, retained even when it is too weak to remove from ordinary data.</summary>
        public int? CandidateRowIndex { get; }

        /// <summary>Non-null only for a strong or explicitly confirmed header.</summary>
        public int? AcceptedHeaderRowIndex { get; }
        public HeaderDecision Decision { get; }
        public double Confidence { get; }
        public double RunnerUpMargin { get; }
        public IReadOnlyList<AlignmentReason> Evidence { get; }
        public int? ConfirmedDataStartRowIndex { get; }
        public bool HasAcceptedHeader => AcceptedHeaderRowIndex.HasValue;

        /// <summary>
        /// Weak and ambiguous candidates deliberately return zero, keeping every physical row as data.
        /// </summary>
        public int DataStartRowIndex => AcceptedHeaderRowIndex.HasValue
            ? ConfirmedDataStartRowIndex ?? AcceptedHeaderRowIndex.Value + 1
            : 0;
    }

    /// <summary>
    /// Conservative, scored header inference. A legacy wordy-row candidate is evidence only: acceptance also
    /// requires a strong score, separation from the runner-up, and at least two independent discriminator signals.
    /// </summary>
    public static class HeaderDetector
    {
        private static readonly HashSet<string> HeaderAliases = new HashSet<string>(StringComparer.Ordinal)
        {
            "id", "identifier", "key", "code", "number", "no", "ref", "reference", "uuid", "guid", "tag",
            "name", "description", "type", "status", "date", "time", "timestamp", "amount", "qty", "quantity",
            "price", "total", "value", "category", "group", "owner", "email", "address", "comment", "comments",
            "note", "notes", "version", "revision", "active", "enabled", "source", "target", "path", "file",
            "filename", "sheet", "row", "column", "country", "city"
        };

        public static HeaderDetectionResult Detect(
            AlignmentTable table,
            AlignmentTable peerRevision = null,
            HeaderDetectionOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            options = options ?? new HeaderDetectionOptions();
            ValidateOptions(options);
            cancellationToken.ThrowIfCancellationRequested();

            if (options.ConfirmedNoHeader)
            {
                if (options.ConfirmedHeaderRowIndex.HasValue || options.ConfirmedDataStartRowIndex.HasValue)
                    throw new ArgumentException(
                        "A no-header confirmation cannot also specify header or data row indices.",
                        nameof(options));
                return new HeaderDetectionResult(
                    null,
                    null,
                    HeaderDecision.ConfirmedNoHeader,
                    1D,
                    1D,
                    new[]
                    {
                        new AlignmentReason(
                            "header.user-confirmed-none",
                            "The user explicitly confirmed that every row is ordinary data.",
                            1D),
                    });
            }

            if (options.ConfirmedHeaderRowIndex.HasValue)
            {
                int confirmed = options.ConfirmedHeaderRowIndex.Value;
                if (confirmed < 0 || confirmed >= table.RowCount)
                    throw new ArgumentOutOfRangeException(
                        nameof(options), "The confirmed header row is outside the table.");
                int dataStart = options.ConfirmedDataStartRowIndex ?? confirmed + 1;
                if (dataStart <= confirmed || dataStart > table.RowCount)
                    throw new ArgumentOutOfRangeException(
                        nameof(options),
                        "The confirmed first data row must follow the header and remain within or just after the table.");
                return new HeaderDetectionResult(
                    confirmed,
                    confirmed,
                    HeaderDecision.ConfirmedHeader,
                    1D,
                    1D,
                    new[]
                    {
                        new AlignmentReason(
                            "header.user-confirmed",
                            "The header row was explicitly confirmed by the user.",
                            1D),
                    },
                    dataStart);
            }

            if (options.ConfirmedDataStartRowIndex.HasValue)
                throw new ArgumentException(
                    "A confirmed first data row requires a confirmed header row.",
                    nameof(options));

            if (table.RowCount == 0 || table.ColumnCount == 0)
                return DataResult(null, 0D, 0D, "header.no-data", "No populated table evidence is available.");

            int modalWidth = ModalWidth(table, options.CommentPrefix, cancellationToken);
            if (modalWidth <= 0)
                return DataResult(null, 0D, 0D, "header.no-modal-width", "No stable row width could be established.");

            int firstEligible = FirstEligibleRow(table, modalWidth, cancellationToken);
            int limit = Math.Min(table.RowCount, options.MaxScanRows);
            var candidates = new List<CandidateScore>();
            for (int row = 0; row < limit; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                AlignmentRow candidate = table.Rows[row];
                if (candidate.Count != modalWidth) continue;
                int nonblank = CountNonblank(candidate);
                if (nonblank < Math.Max(1, (modalWidth + 1) / 2)) continue;
                int textual = CountTextual(candidate);
                if (textual < Math.Max(1, (int)Math.Ceiling(nonblank * 0.60D))) continue;

                candidates.Add(ScoreCandidate(
                    table,
                    peerRevision,
                    row,
                    modalWidth,
                    firstEligible,
                    options,
                    cancellationToken));
            }

            if (candidates.Count == 0)
                return DataResult(
                    null,
                    0D,
                    0D,
                    "header.no-wordy-candidate",
                    "No modal-width, mostly populated, predominantly textual row qualified even as a weak candidate.");

            candidates.Sort((left, right) =>
            {
                int score = right.Score.CompareTo(left.Score);
                return score != 0 ? score : left.RowIndex.CompareTo(right.RowIndex);
            });
            CandidateScore best = candidates[0];
            double runnerUp = candidates.Count > 1 ? candidates[1].Score : 0D;
            double margin = Math.Max(0D, best.Score - runnerUp);
            var evidence = new List<AlignmentReason>(best.Reasons)
            {
                new AlignmentReason(
                    "header.runner-up-margin",
                    "The best candidate is separated from the runner-up by " +
                    margin.ToString("0.000", CultureInfo.InvariantCulture) + ".",
                    margin),
            };

            bool strong = best.Score >= options.StrongScoreThreshold &&
                          margin >= options.StrongRunnerUpMargin &&
                          best.IndependentStrongSignals >= 2;
            if (strong)
            {
                evidence.Add(new AlignmentReason(
                    "header.strong",
                    "The candidate passed the score, margin, and independent-evidence gates."));
                return new HeaderDetectionResult(
                    best.RowIndex,
                    best.RowIndex,
                    HeaderDecision.StrongHeader,
                    best.Score,
                    margin,
                    evidence);
            }

            if (best.Score >= options.AmbiguousScoreThreshold)
            {
                evidence.Add(new AlignmentReason(
                    "header.ambiguous-treated-as-data",
                    "Header evidence is not strong and separated enough; the candidate remains an ordinary data row."));
                return new HeaderDetectionResult(
                    best.RowIndex,
                    null,
                    HeaderDecision.Ambiguous,
                    best.Score,
                    margin,
                    evidence);
            }

            evidence.Add(new AlignmentReason(
                "header.weak-treated-as-data",
                "The best raw candidate is weak; every row remains ordinary data."));
            return new HeaderDetectionResult(
                best.RowIndex,
                null,
                HeaderDecision.Data,
                best.Score,
                margin,
                evidence);
        }

        internal static bool IsHeaderAlias(string normalizedLabel) =>
            HeaderAliases.Contains(normalizedLabel ?? "");

        internal static double KeyAliasStrength(string normalizedLabel)
        {
            switch (normalizedLabel ?? "")
            {
                case "id":
                case "identifier":
                case "key":
                case "uuid":
                case "guid":
                    return 1D;
                case "code":
                case "number":
                case "no":
                case "ref":
                case "reference":
                case "tag":
                    return 0.85D;
                case "email":
                case "path":
                case "filename":
                    return 0.60D;
                case "name":
                    return 0.35D;
                default:
                    return 0D;
            }
        }

        private static CandidateScore ScoreCandidate(
            AlignmentTable table,
            AlignmentTable peer,
            int rowIndex,
            int modalWidth,
            int firstEligible,
            HeaderDetectionOptions options,
            CancellationToken cancellationToken)
        {
            AlignmentRow row = table.Rows[rowIndex];
            int nonblank = CountNonblank(row);
            double coverage = (double)nonblank / modalWidth;
            double textual = nonblank == 0 ? 0D : (double)CountTextual(row) / nonblank;
            double unique = UniqueNormalizedRatio(row, options.CommentPrefix);
            double alias = AliasRatio(row, options.CommentPrefix);
            double transition = TypeAndProfileTransition(
                table,
                rowIndex,
                modalWidth,
                options.MaxEvidenceRows,
                options.CommentPrefix,
                cancellationToken);
            double position = rowIndex == firstEligible
                ? 1D
                : Math.Max(0D, 1D - (rowIndex - firstEligible) / 5D);
            double peerAgreement = PeerAgreement(row, peer, modalWidth, options, cancellationToken);

            double score = Clamp(
                0.12D * coverage +
                0.10D * textual +
                0.08D * unique +
                0.20D * alias +
                0.25D * transition +
                0.10D * position +
                0.15D * peerAgreement);

            int signals = 0;
            if (alias >= 0.50D) signals++;
            if (transition >= 0.50D) signals++;
            if (peerAgreement >= 0.80D) signals++;

            var reasons = new List<AlignmentReason>
            {
                Metric("header.coverage", "Nonblank header coverage", coverage, 0.12D),
                Metric("header.textual", "Predominantly textual candidate", textual, 0.10D),
                Metric("header.unique-labels", "Unique normalized labels", unique, 0.08D),
                Metric("header.aliases", "Recognized header labels", alias, 0.20D),
                Metric("header.type-profile-transition", "Transition into later-row value profiles", transition, 0.25D),
                Metric("header.position", "Early-row position", position, 0.10D),
                Metric("header.peer-agreement", "Repeated normalized labels in the peer revision", peerAgreement, 0.15D),
                new AlignmentReason(
                    "header.score",
                    "Candidate row " + rowIndex.ToString(CultureInfo.InvariantCulture) + " scored " +
                    score.ToString("0.000", CultureInfo.InvariantCulture) + ".",
                    score),
            };
            return new CandidateScore(rowIndex, score, signals, reasons);
        }

        private static double TypeAndProfileTransition(
            AlignmentTable table,
            int candidateRow,
            int width,
            int maxEvidenceRows,
            string commentPrefix,
            CancellationToken cancellationToken)
        {
            int discriminating = 0;
            int eligible = 0;
            for (int column = 0; column < width; column++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                table.TryGetCell(candidateRow, column, out string candidate);
                if (AlignmentText.IsBlank(candidate)) continue;
                eligible++;
                ValueProfile candidateProfile = Profile(candidate);
                string normalizedCandidate = AlignmentText.NormalizeLabel(
                    column == 0 ? candidate : candidate,
                    column == 0 ? commentPrefix : null);
                bool alias = IsHeaderAlias(normalizedCandidate);
                int observed = 0;
                int differentProfile = 0;
                int nonAliasValues = 0;
                int end = Math.Min(table.RowCount, candidateRow + 1 + maxEvidenceRows);
                for (int row = candidateRow + 1; row < end; row++)
                {
                    if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (IsCommentRow(table.Rows[row], commentPrefix)) continue;
                    if (!table.TryGetCell(row, column, out string value) || AlignmentText.IsBlank(value)) continue;
                    observed++;
                    if (Profile(value) != candidateProfile) differentProfile++;
                    if (!IsHeaderAlias(AlignmentText.NormalizeLabel(value))) nonAliasValues++;
                }

                if (observed == 0) continue;
                bool transitioned = differentProfile >= Math.Max(1, (int)Math.Ceiling(observed * 0.60D));
                bool aliasSeparated = alias && nonAliasValues >= Math.Max(1, (int)Math.Ceiling(observed * 0.75D));
                if (transitioned || aliasSeparated) discriminating++;
            }
            return eligible == 0 ? 0D : (double)discriminating / eligible;
        }

        private static double PeerAgreement(
            AlignmentRow candidate,
            AlignmentTable peer,
            int modalWidth,
            HeaderDetectionOptions options,
            CancellationToken cancellationToken)
        {
            if (peer == null || peer.RowCount == 0) return 0D;
            int limit = Math.Min(peer.RowCount, options.MaxScanRows);
            double best = 0D;
            for (int row = 0; row < limit; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (peer.Rows[row].Count != modalWidth) continue;
                int compared = 0;
                int equal = 0;
                for (int column = 0; column < modalWidth; column++)
                {
                    string left = candidate.Values[column];
                    string right = peer.Rows[row].Values[column];
                    string leftNormalized = AlignmentText.NormalizeLabel(
                        left, column == 0 ? options.CommentPrefix : null);
                    string rightNormalized = AlignmentText.NormalizeLabel(
                        right, column == 0 ? options.CommentPrefix : null);
                    if (leftNormalized.Length == 0 && rightNormalized.Length == 0) continue;
                    compared++;
                    if (string.Equals(leftNormalized, rightNormalized, StringComparison.Ordinal)) equal++;
                }
                if (compared > 0) best = Math.Max(best, (double)equal / compared);
            }
            return best;
        }

        private static int ModalWidth(
            AlignmentTable table,
            string commentPrefix,
            CancellationToken cancellationToken)
        {
            var counts = new Dictionary<int, int>();
            var first = new Dictionary<int, int>();
            int order = 0;
            for (int row = 0; row < table.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                AlignmentRow values = table.Rows[row];
                if (IsCommentRow(values, commentPrefix) || CountNonblank(values) == 0) continue;
                if (!first.ContainsKey(values.Count)) first.Add(values.Count, order++);
                counts.TryGetValue(values.Count, out int count);
                counts[values.Count] = count + 1;
            }
            if (counts.Count == 0)
            {
                for (int row = 0; row < table.RowCount; row++)
                {
                    if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                    int width = table.Rows[row].Count;
                    if (!first.ContainsKey(width)) first.Add(width, order++);
                    counts.TryGetValue(width, out int count);
                    counts[width] = count + 1;
                }
            }

            int bestWidth = 0;
            int bestCount = -1;
            int bestOrder = int.MaxValue;
            foreach (KeyValuePair<int, int> pair in counts)
            {
                int seen = first[pair.Key];
                if (pair.Value > bestCount || pair.Value == bestCount && seen < bestOrder)
                {
                    bestWidth = pair.Key;
                    bestCount = pair.Value;
                    bestOrder = seen;
                }
            }
            return bestWidth;
        }

        private static int FirstEligibleRow(
            AlignmentTable table,
            int modalWidth,
            CancellationToken cancellationToken)
        {
            for (int row = 0; row < table.RowCount; row++)
            {
                if ((row & 0x3F) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (table.Rows[row].Count == modalWidth && CountNonblank(table.Rows[row]) > 0) return row;
            }
            return 0;
        }

        private static int CountNonblank(AlignmentRow row) =>
            row.Values.Count(value => !AlignmentText.IsBlank(value));

        private static int CountTextual(AlignmentRow row) =>
            row.Values.Count(value => !AlignmentText.IsBlank(value) && Profile(value) == ValueProfile.Text);

        private static double UniqueNormalizedRatio(AlignmentRow row, string commentPrefix)
        {
            var labels = new HashSet<string>(StringComparer.Ordinal);
            int nonblank = 0;
            for (int column = 0; column < row.Count; column++)
            {
                string normalized = AlignmentText.NormalizeLabel(
                    row.Values[column], column == 0 ? commentPrefix : null);
                if (normalized.Length == 0) continue;
                nonblank++;
                labels.Add(normalized);
            }
            return nonblank == 0 ? 0D : (double)labels.Count / nonblank;
        }

        private static double AliasRatio(AlignmentRow row, string commentPrefix)
        {
            int nonblank = 0;
            int aliases = 0;
            for (int column = 0; column < row.Count; column++)
            {
                string normalized = AlignmentText.NormalizeLabel(
                    row.Values[column], column == 0 ? commentPrefix : null);
                if (normalized.Length == 0) continue;
                nonblank++;
                if (IsHeaderAlias(normalized)) aliases++;
            }
            return nonblank == 0 ? 0D : (double)aliases / nonblank;
        }

        private static bool IsCommentRow(AlignmentRow row, string commentPrefix)
        {
            if (string.IsNullOrEmpty(commentPrefix) || row.Count == 0 || row.Values[0] == null) return false;
            return row.Values[0].StartsWith(commentPrefix, StringComparison.Ordinal);
        }

        private static ValueProfile Profile(string value)
        {
            if (AlignmentText.IsBlank(value)) return ValueProfile.Blank;
            string trimmed = value.Trim();
            if (double.TryParse(
                trimmed,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture,
                out _)) return ValueProfile.Number;
            if (bool.TryParse(trimmed, out _)) return ValueProfile.Boolean;
            if (DateTimeOffset.TryParseExact(
                trimmed,
                new[] { "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssK", "o" },
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out _)) return ValueProfile.Date;
            if (trimmed.StartsWith("=", StringComparison.Ordinal)) return ValueProfile.Formula;
            if (trimmed.Any(char.IsDigit) && trimmed.Any(character => !char.IsLetterOrDigit(character)))
                return ValueProfile.StructuredText;
            if (trimmed.Any(char.IsDigit) && trimmed.Any(char.IsLetter)) return ValueProfile.StructuredText;
            return ValueProfile.Text;
        }

        private static HeaderDetectionResult DataResult(
            int? candidate,
            double confidence,
            double margin,
            string code,
            string message) =>
            new HeaderDetectionResult(
                candidate,
                null,
                HeaderDecision.Data,
                confidence,
                margin,
                new[] { new AlignmentReason(code, message) });

        private static AlignmentReason Metric(string code, string label, double value, double weight) =>
            new AlignmentReason(
                code,
                label + ": " + value.ToString("0.000", CultureInfo.InvariantCulture) + ".",
                value * weight);

        private static double Clamp(double value) => Math.Max(0D, Math.Min(1D, value));

        private static void ValidateOptions(HeaderDetectionOptions options)
        {
            if (options.MaxScanRows <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxScanRows));
            if (options.MaxEvidenceRows <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEvidenceRows));
            if (options.StrongScoreThreshold < 0D || options.StrongScoreThreshold > 1D)
                throw new ArgumentOutOfRangeException(nameof(options.StrongScoreThreshold));
            if (options.StrongRunnerUpMargin < 0D || options.StrongRunnerUpMargin > 1D)
                throw new ArgumentOutOfRangeException(nameof(options.StrongRunnerUpMargin));
            if (options.AmbiguousScoreThreshold < 0D || options.AmbiguousScoreThreshold > 1D)
                throw new ArgumentOutOfRangeException(nameof(options.AmbiguousScoreThreshold));
        }

        private enum ValueProfile
        {
            Blank,
            Text,
            StructuredText,
            Number,
            Boolean,
            Date,
            Formula,
        }

        private sealed class CandidateScore
        {
            internal CandidateScore(
                int rowIndex,
                double score,
                int independentStrongSignals,
                IReadOnlyList<AlignmentReason> reasons)
            {
                RowIndex = rowIndex;
                Score = score;
                IndependentStrongSignals = independentStrongSignals;
                Reasons = reasons;
            }

            internal int RowIndex { get; }
            internal double Score { get; }
            internal int IndependentStrongSignals { get; }
            internal IReadOnlyList<AlignmentReason> Reasons { get; }
        }
    }
}
