using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using DiffPlex;
using DiffPlex.Chunkers;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public enum RowMatchBasis { UniqueConstantSequence, ExplicitPosition }
    public sealed class ReviewRowMatch
    {
        internal ReviewRowMatch(int before, int after, RowMatchBasis basis) { BeforeRow = before; AfterRow = after; Basis = basis; }
        public int BeforeRow { get; }
        public int AfterRow { get; }
        public RowMatchBasis Basis { get; }
    }
    public sealed class ReviewRowSuggestion
    {
        internal ReviewRowSuggestion(int before, int after, int shared, int compared)
        { BeforeRow = before; AfterRow = after; SharedInformativeCells = shared; ComparedCells = compared; }
        public int BeforeRow { get; }
        public int AfterRow { get; }
        public int SharedInformativeCells { get; }
        public int ComparedCells { get; }
    }
    public sealed class ReviewRowRegion
    {
        internal ReviewRowRegion(int beforeStart, int beforeCount, int afterStart, int afterCount,
            IEnumerable<ReviewRowSuggestion> suggestions, string reason)
        {
            BeforeStart = beforeStart; BeforeCount = beforeCount; AfterStart = afterStart; AfterCount = afterCount;
            Suggestions = Array.AsReadOnly(suggestions.ToArray()); Reason = reason;
        }
        public int BeforeStart { get; }
        public int BeforeCount { get; }
        public int AfterStart { get; }
        public int AfterCount { get; }
        public ReadOnlyCollection<ReviewRowSuggestion> Suggestions { get; }
        public string Reason { get; }
        public bool IsOneSided => BeforeCount == 0 || AfterCount == 0;
    }
    public sealed class ReviewRowAlignment
    {
        internal ReviewRowAlignment(IEnumerable<ReviewRowMatch> matches, IEnumerable<ReviewRowRegion> regions,
            IEnumerable<ReviewIssue> issues, int beforeRows, int afterRows)
        {
            Matches = Array.AsReadOnly(matches.ToArray()); Regions = Array.AsReadOnly(regions.ToArray());
            Issues = Array.AsReadOnly(issues.ToArray()); BeforeRows = beforeRows; AfterRows = afterRows;
        }
        public ReadOnlyCollection<ReviewRowMatch> Matches { get; }
        public ReadOnlyCollection<ReviewRowRegion> Regions { get; }
        public ReadOnlyCollection<ReviewIssue> Issues { get; }
        public int BeforeRows { get; }
        public int AfterRows { get; }
    }

    /// <summary>First engine candidate: exact unique constant anchors, ordered by DiffPlex; partial candidates never silently bind.</summary>
    public sealed class OrderedRowAlignment
    {
        public ReviewRowAlignment Align(ReviewSheet before, ReviewSheet after, ResolvedReviewScope scope,
            bool explicitFixedPosition = false, int maximumRows = 20000, int candidateBudget = 250000,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null || after == null || scope == null) throw new ArgumentNullException();
            if (maximumRows < 1 || candidateBudget < 0) throw new ArgumentOutOfRangeException();
            cancellationToken.ThrowIfCancellationRequested();
            int n = before.LastContentRow(scope.BeforeCap), m = after.LastContentRow(scope.AfterCap);
            if (n > maximumRows || m > maximumRows)
                throw new InvalidOperationException("The alignment trial row limit is " + maximumRows +
                    "; observed scoped rows are " + n + " and " + m + ". No rows were truncated. Reduce scope or use the larger-file reader when available.");
            var matches = new List<ReviewRowMatch>(); var regions = new List<ReviewRowRegion>(); var issues = new List<ReviewIssue>();
            if (explicitFixedPosition)
            {
                for (int r = 1; r <= Math.Min(n, m); r++) matches.Add(new ReviewRowMatch(r, r, RowMatchBasis.ExplicitPosition));
            }
            else
            {
                var a = Enumerable.Range(1, n).Select(r => scope.AnchorToken(before, r, true)).ToArray();
                var b = Enumerable.Range(1, m).Select(r => scope.AnchorToken(after, r, false)).ToArray();
                var ac = Counts(a); var bc = Counts(b); var ids = new Dictionary<string, int>(StringComparer.Ordinal);
                string[] at = Tokens(a, ac, bc, ids, "A"), bt = Tokens(b, bc, ac, ids, "B");
                var diff = Differ.Instance.CreateDiffs(string.Join("\n", at), string.Join("\n", bt), false, false, LineChunker.Instance);
                int ai = 0, bi = 0;
                foreach (var block in diff.DiffBlocks)
                {
                    while (ai < block.DeleteStartA && bi < block.InsertStartB)
                    { AddAnchor(ai++, bi++); }
                    ai = block.DeleteStartA + block.DeleteCountA;
                    bi = block.InsertStartB + block.InsertCountB;
                }
                while (ai < n && bi < m) { AddAnchor(ai++, bi++); }

                // Independent column evidence can contradict a whole-row hypothesis. Keep the sheet unresolved then.
                var conflicts = ConflictingColumnShifts(before, after, scope, n, m, cancellationToken);
                foreach (var conflict in conflicts)
                {
                    matches.RemoveAll(match => (match.BeforeRow >= conflict.BeforeStart && match.BeforeRow <= conflict.BeforeEnd) ||
                        (match.AfterRow >= conflict.AfterStart && match.AfterRow <= conflict.AfterEnd));
                    issues.Add(ReviewIssue.Create("SR.ALIGN.PARTIAL_SHIFT", "Official columns support incompatible offsets over the same consecutive rows. Review this area, or explicitly compare its fixed positions as data changes.",
                        new Dictionary<string, string> { ["beforeArtifact"] = before.ArtifactId, ["afterArtifact"] = after.ArtifactId,
                            ["beforeSheet"] = before.Name, ["afterSheet"] = after.Name,
                            ["beforeRange"] = conflict.BeforeStart + ":" + conflict.BeforeEnd,
                            ["afterRange"] = conflict.AfterStart + ":" + conflict.AfterEnd,
                            ["columnOffsets"] = conflict.Evidence }));
                }

                void AddAnchor(int x, int y)
                {
                    if (x < n && y < m && a[x] != null && a[x] == b[y] && ac[a[x]] == 1 && bc[a[x]] == 1)
                        matches.Add(new ReviewRowMatch(x + 1, y + 1, RowMatchBasis.UniqueConstantSequence));
                }
            }
            int nextA = 1, nextB = 1, remaining = candidateBudget; bool budgetIssueReported = false;
            foreach (var match in matches.Concat(new[] { new ReviewRowMatch(n + 1, m + 1, RowMatchBasis.UniqueConstantSequence) }))
            {
                int na = match.BeforeRow - nextA, nb = match.AfterRow - nextB;
                if (na != 0 || nb != 0)
                {
                    var candidates = new List<ReviewRowSuggestion>(); bool exhausted = false;
                    for (int x = nextA; x < nextA + na && nb != 0; x++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var rowCandidates = new List<ReviewRowSuggestion>();
                        for (int y = nextB; y < nextB + nb; y++)
                        {
                            if (remaining-- <= 0) { exhausted = true; break; }
                            int shared = 0, compared = 0;
                            for (int s = 0; s < scope.Columns.Count; s++)
                            {
                                if (!scope.Columns[s].BeforeColumn.HasValue || !scope.Columns[s].AfterColumn.HasValue) continue;
                                var av = scope.At(before, x, s, true); var bv = scope.At(after, y, s, false);
                                if (!av.IsObserved || !bv.IsObserved) continue;
                                if (av.HasContent || bv.HasContent) compared++;
                                if (av.Content.HasInformativeConstant && av.Content.Equals(bv.Content)) shared++;
                            }
                            if (shared > 0) rowCandidates.Add(new ReviewRowSuggestion(x, y, shared, compared));
                        }
                        // Limit presentation, not the raw unresolved region. Equal-score peers remain explicit alternatives.
                        candidates.AddRange(rowCandidates.OrderByDescending(c => c.SharedInformativeCells)
                            .ThenBy(c => Math.Abs(c.AfterRow - c.BeforeRow)).Take(5));
                        if (exhausted) break;
                    }
                    string reason = na == 0 ? "Rows occur only on the after side between established anchors." :
                        nb == 0 ? "Rows occur only on the before side between established anchors." :
                        "Row correspondence requires review. Suggestions are not accepted matches.";
                    if (exhausted)
                    {
                        reason += " Candidate search reached its budget; all raw rows remain available.";
                        if (!budgetIssueReported) issues.Add(ReviewIssue.Create("SR.ALIGN.CANDIDATE_BUDGET", reason + " Subsequent unsolved regions share this exhausted search budget.", new Dictionary<string, string>
                        { ["beforeArtifact"] = before.ArtifactId, ["afterArtifact"] = after.ArtifactId,
                            ["beforeSheet"] = before.Name, ["afterSheet"] = after.Name,
                            ["beforeRange"] = nextA + ":" + (nextA + na - 1), ["afterRange"] = nextB + ":" + (nextB + nb - 1) }));
                        budgetIssueReported = true;
                    }
                    regions.Add(new ReviewRowRegion(nextA, na, nextB, nb, candidates, reason));
                }
                nextA = match.BeforeRow + 1; nextB = match.AfterRow + 1;
            }
            return new ReviewRowAlignment(matches, regions, issues, n, m);
        }

        private static Dictionary<string, int> Counts(string[] values) => values.Where(v => v != null)
            .GroupBy(v => v, StringComparer.Ordinal).ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
        private static string[] Tokens(string[] values, Dictionary<string, int> own, Dictionary<string, int> other,
            Dictionary<string, int> ids, string side) => values.Select((v, i) =>
            {
                if (v == null || own[v] != 1 || !other.TryGetValue(v, out int count) || count != 1) return side + i;
                if (!ids.TryGetValue(v, out int id)) ids.Add(v, id = ids.Count);
                return "K" + id.ToString(CultureInfo.InvariantCulture);
            }).ToArray();

        private sealed class ShiftConflict
        {
            public int BeforeStart, BeforeEnd, AfterStart, AfterEnd;
            public string Evidence;
        }

        private static List<ShiftConflict> ConflictingColumnShifts(ReviewSheet a, ReviewSheet b,
            ResolvedReviewScope scope, int n, int m, CancellationToken cancellationToken)
        {
            var evidence = new Dictionary<int, List<KeyValuePair<int, int>>>();
            for (int s = 0; s < scope.Columns.Count; s++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!scope.Columns[s].BeforeColumn.HasValue || !scope.Columns[s].AfterColumn.HasValue) continue;
                var ar = Unique(a, n, s, true); var br = Unique(b, m, s, false);
                foreach (var pair in ar.Where(p => br.ContainsKey(p.Key)))
                {
                    if (!evidence.TryGetValue(pair.Value, out var peers)) evidence.Add(pair.Value, peers = new List<KeyValuePair<int, int>>());
                    peers.Add(new KeyValuePair<int, int>(s + 1, br[pair.Key] - pair.Value));
                }
            }
            // Compare witnesses on the same rows, not the dominant offset of unrelated regions of a sheet.
            var runs = new Dictionary<string, ShiftConflict>(StringComparer.Ordinal); var result = new List<ShiftConflict>();
            foreach (var row in evidence.OrderBy(e => e.Key))
                for (int x = 0; x < row.Value.Count; x++)
                    for (int y = x + 1; y < row.Value.Count; y++)
                    {
                        var left = row.Value[x]; var right = row.Value[y];
                        if (left.Value == right.Value) continue;
                        string key = left.Key + ":" + left.Value + "," + right.Key + ":" + right.Value;
                        if (!runs.TryGetValue(key, out var run) || run.BeforeEnd + 1 != row.Key)
                        {
                            run = new ShiftConflict { BeforeStart = row.Key, BeforeEnd = row.Key,
                                AfterStart = row.Key + Math.Min(left.Value, right.Value), AfterEnd = row.Key + Math.Max(left.Value, right.Value), Evidence = key };
                            runs[key] = run;
                        }
                        else
                        {
                            run.BeforeEnd = row.Key; run.AfterEnd = row.Key + Math.Max(left.Value, right.Value);
                            if (run.BeforeEnd - run.BeforeStart == 2) result.Add(run);
                        }
                    }
            return result;
            Dictionary<string, int> Unique(ReviewSheet sheet, int count, int slot, bool before) => Enumerable.Range(1, count)
                .Select(r => new { Row = r, Cell = scope.At(sheet, r, slot, before) })
                .Where(v => v.Cell.IsObserved && v.Cell.Content.HasInformativeConstant)
                .GroupBy(v => v.Cell.Content.IdentityToken, StringComparer.Ordinal).Where(g => g.Count() == 1)
                .ToDictionary(g => g.Key, g => g.Single().Row, StringComparer.Ordinal);
        }
    }
}
