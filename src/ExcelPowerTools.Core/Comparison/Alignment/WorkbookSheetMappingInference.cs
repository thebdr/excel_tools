using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    /// <summary>
    /// Conservative pairwise sheet mapping. Exact case-insensitive names are accepted when unique. A rename is
    /// auto-resolved only when structural and content evidence both exist, the score reaches the configured
    /// threshold, and the pair is a clear reciprocal one-to-one winner. Everything else remains visible for
    /// user confirmation.
    /// </summary>
    public static class WorkbookSheetMappingInference
    {
        private const double AddInIdentityWeight = 0.25;
        private const double CodeNameWeight = 0.20;
        private const double StructureFingerprintWeight = 0.30;
        private const double ContentFingerprintWeight = 0.35;
        private const double HeaderFingerprintWeight = 0.12;
        private const double FormulaFingerprintWeight = 0.12;
        private const double DimensionsWeight = 0.08;
        private const double StoredCellCountWeight = 0.05;
        private const double ComparisonEpsilon = 0.000000001;

        public static WorkbookSheetMappingResult Infer(
            IEnumerable<WorkbookSheetFingerprint> beforeSheets,
            IEnumerable<WorkbookSheetFingerprint> afterSheets,
            WorkbookSheetMappingOptions options = null)
        {
            List<WorkbookSheetFingerprint> before = CopyAndValidate(beforeSheets, nameof(beforeSheets));
            List<WorkbookSheetFingerprint> after = CopyAndValidate(afterSheets, nameof(afterSheets));
            options = options ?? WorkbookSheetMappingOptions.Default;

            before.Sort(CompareSheets);
            after.Sort(CompareSheets);

            var mappings = new List<ResolvedSheetMapping>();
            var mappedBefore = new HashSet<string>(StringComparer.Ordinal);
            var mappedAfter = new HashSet<string>(StringComparer.Ordinal);
            var nonUniqueExactNameBefore = new HashSet<string>(StringComparer.Ordinal);

            ResolveUniqueExactNames(
                before,
                after,
                mappings,
                mappedBefore,
                mappedAfter,
                nonUniqueExactNameBefore);

            List<CandidatePair> pairs = BuildRenameCandidates(before, after, mappedBefore, mappedAfter);
            Dictionary<string, List<CandidatePair>> byBefore = GroupByBefore(pairs, before, mappedBefore);
            Dictionary<string, List<CandidatePair>> byAfter = GroupByAfter(pairs, after, mappedAfter);

            foreach (WorkbookSheetFingerprint source in before)
            {
                if (mappedBefore.Contains(source.SheetKey)) continue;
                List<CandidatePair> sourceCandidates = byBefore[source.SheetKey];
                if (sourceCandidates.Count == 0) continue;

                CandidatePair winner = sourceCandidates[0];
                if (!MeetsThreshold(winner.Confidence, options.MinimumRenameConfidence) ||
                    !HasWinnerMargin(sourceCandidates, options.MinimumWinnerMargin))
                    continue;

                List<CandidatePair> targetCandidates = byAfter[winner.After.SheetKey];
                if (targetCandidates.Count == 0 || !ReferenceEquals(targetCandidates[0], winner) ||
                    !HasWinnerMargin(targetCandidates, options.MinimumWinnerMargin))
                    continue;

                mappings.Add(new ResolvedSheetMapping(
                    winner.Before,
                    winner.After,
                    SheetMappingKind.InferredRename,
                    winner.Confidence,
                    winner.ReasonCodes));
                mappedBefore.Add(winner.Before.SheetKey);
                mappedAfter.Add(winner.After.SheetKey);
            }

            mappings.Sort(CompareMappings);
            var unresolved = new List<UnresolvedSheetMapping>();
            foreach (WorkbookSheetFingerprint source in before)
            {
                if (mappedBefore.Contains(source.SheetKey)) continue;
                List<CandidatePair> sourceCandidates = byBefore[source.SheetKey];
                var reasons = ExplainUnresolved(
                    source,
                    sourceCandidates,
                    byAfter,
                    nonUniqueExactNameBefore,
                    options);
                unresolved.Add(new UnresolvedSheetMapping(
                    source,
                    sourceCandidates.Select(ToPublicCandidate).ToArray(),
                    reasons));
            }

            var unmatchedAfter = after
                .Where(sheet => !mappedAfter.Contains(sheet.SheetKey))
                .ToArray();
            return new WorkbookSheetMappingResult(
                mappings,
                unresolved,
                Array.Empty<WorkbookSheetFingerprint>(),
                unmatchedAfter);
        }

        private static void ResolveUniqueExactNames(
            IReadOnlyList<WorkbookSheetFingerprint> before,
            IReadOnlyList<WorkbookSheetFingerprint> after,
            ICollection<ResolvedSheetMapping> mappings,
            ISet<string> mappedBefore,
            ISet<string> mappedAfter,
            ISet<string> nonUniqueExactNameBefore)
        {
            Dictionary<string, List<WorkbookSheetFingerprint>> beforeByName = GroupByName(before);
            Dictionary<string, List<WorkbookSheetFingerprint>> afterByName = GroupByName(after);
            var names = new SortedSet<string>(beforeByName.Keys, StringComparer.OrdinalIgnoreCase);
            names.UnionWith(afterByName.Keys);

            foreach (string name in names)
            {
                beforeByName.TryGetValue(name, out List<WorkbookSheetFingerprint> left);
                afterByName.TryGetValue(name, out List<WorkbookSheetFingerprint> right);
                left = left ?? new List<WorkbookSheetFingerprint>();
                right = right ?? new List<WorkbookSheetFingerprint>();

                if (left.Count == 1 && right.Count == 1)
                {
                    WorkbookSheetFingerprint beforeSheet = left[0];
                    WorkbookSheetFingerprint afterSheet = right[0];
                    mappings.Add(new ResolvedSheetMapping(
                        beforeSheet,
                        afterSheet,
                        SheetMappingKind.ExactName,
                        1,
                        new[] { SheetMappingReasonCodes.ExactName }));
                    mappedBefore.Add(beforeSheet.SheetKey);
                    mappedAfter.Add(afterSheet.SheetKey);
                }
                else if (left.Count > 0 && right.Count > 0)
                {
                    foreach (WorkbookSheetFingerprint sheet in left)
                        nonUniqueExactNameBefore.Add(sheet.SheetKey);
                }
            }
        }

        private static List<CandidatePair> BuildRenameCandidates(
            IReadOnlyList<WorkbookSheetFingerprint> before,
            IReadOnlyList<WorkbookSheetFingerprint> after,
            ISet<string> mappedBefore,
            ISet<string> mappedAfter)
        {
            var pairs = new List<CandidatePair>();
            foreach (WorkbookSheetFingerprint source in before)
            {
                if (mappedBefore.Contains(source.SheetKey)) continue;
                foreach (WorkbookSheetFingerprint target in after)
                {
                    if (mappedAfter.Contains(target.SheetKey)) continue;
                    CandidatePair candidate = EvaluateRename(source, target);
                    if (candidate != null) pairs.Add(candidate);
                }
            }
            return pairs;
        }

        private static CandidatePair EvaluateRename(
            WorkbookSheetFingerprint before,
            WorkbookSheetFingerprint after)
        {
            if (before.Kind != WorkbookSheetKind.Unknown &&
                after.Kind != WorkbookSheetKind.Unknown &&
                before.Kind != after.Kind)
                return null;

            bool sameStructure = SameToken(before.StructureFingerprint, after.StructureFingerprint);
            bool sameDimensions = before.UsedRowCount > 0 && before.UsedColumnCount > 0 &&
                                  before.UsedRowCount == after.UsedRowCount &&
                                  before.UsedColumnCount == after.UsedColumnCount;
            bool sameContent = SameToken(before.ContentFingerprint, after.ContentFingerprint);
            bool sameHeader = SameToken(before.HeaderFingerprint, after.HeaderFingerprint);
            bool sameFormula = SameToken(before.FormulaFingerprint, after.FormulaFingerprint);

            // Workbook-local IDs and code names are supporting evidence, never sufficient rename proof.
            if ((!sameStructure && !sameDimensions) || (!sameContent && !sameHeader && !sameFormula))
                return null;

            var reasons = new List<string>();
            double score = 0;
            if (SameIdentity(before.AddInIdentity, after.AddInIdentity))
            {
                score += AddInIdentityWeight;
                reasons.Add(SheetMappingReasonCodes.AddInIdentity);
            }
            if (SameIdentity(before.CodeName, after.CodeName))
            {
                score += CodeNameWeight;
                reasons.Add(SheetMappingReasonCodes.CodeName);
            }
            if (sameStructure)
            {
                score += StructureFingerprintWeight;
                reasons.Add(SheetMappingReasonCodes.StructureFingerprint);
            }
            if (sameContent)
            {
                score += ContentFingerprintWeight;
                reasons.Add(SheetMappingReasonCodes.ContentFingerprint);
            }
            if (sameHeader)
            {
                score += HeaderFingerprintWeight;
                reasons.Add(SheetMappingReasonCodes.HeaderFingerprint);
            }
            if (sameFormula)
            {
                score += FormulaFingerprintWeight;
                reasons.Add(SheetMappingReasonCodes.FormulaFingerprint);
            }
            if (sameDimensions)
            {
                score += DimensionsWeight;
                reasons.Add(SheetMappingReasonCodes.Dimensions);
            }
            if (HasSimilarStoredCellCount(before.StoredCellCount, after.StoredCellCount))
            {
                score += StoredCellCountWeight;
                reasons.Add(SheetMappingReasonCodes.StoredCellCount);
            }

            return new CandidatePair(before, after, Math.Min(1, Math.Round(score, 6)), reasons);
        }

        private static IReadOnlyList<string> ExplainUnresolved(
            WorkbookSheetFingerprint source,
            IReadOnlyList<CandidatePair> sourceCandidates,
            IReadOnlyDictionary<string, List<CandidatePair>> byAfter,
            ISet<string> nonUniqueExactNameBefore,
            WorkbookSheetMappingOptions options)
        {
            var reasons = new List<string>();
            if (nonUniqueExactNameBefore.Contains(source.SheetKey))
                reasons.Add(SheetMappingReasonCodes.ExactNameNotUnique);

            if (sourceCandidates.Count == 0)
            {
                reasons.Add(SheetMappingReasonCodes.NoSupportedRenameEvidence);
                return new ReadOnlyCollection<string>(reasons);
            }

            CandidatePair winner = sourceCandidates[0];
            if (!MeetsThreshold(winner.Confidence, options.MinimumRenameConfidence))
                reasons.Add(SheetMappingReasonCodes.BelowRenameThreshold);
            if (!HasWinnerMargin(sourceCandidates, options.MinimumWinnerMargin))
                reasons.Add(SheetMappingReasonCodes.CompetingTargets);

            List<CandidatePair> targetCandidates = byAfter[winner.After.SheetKey];
            if (targetCandidates.Count == 0 || !ReferenceEquals(targetCandidates[0], winner) ||
                !HasWinnerMargin(targetCandidates, options.MinimumWinnerMargin))
                reasons.Add(SheetMappingReasonCodes.CompetingSources);
            if (reasons.Count == 0) reasons.Add(SheetMappingReasonCodes.NotAutoResolved);
            return new ReadOnlyCollection<string>(reasons);
        }

        private static Dictionary<string, List<CandidatePair>> GroupByBefore(
            IReadOnlyList<CandidatePair> pairs,
            IReadOnlyList<WorkbookSheetFingerprint> before,
            ISet<string> mappedBefore)
        {
            var result = new Dictionary<string, List<CandidatePair>>(StringComparer.Ordinal);
            foreach (WorkbookSheetFingerprint sheet in before)
                if (!mappedBefore.Contains(sheet.SheetKey)) result.Add(sheet.SheetKey, new List<CandidatePair>());
            foreach (CandidatePair pair in pairs) result[pair.Before.SheetKey].Add(pair);
            foreach (List<CandidatePair> candidates in result.Values) candidates.Sort(CompareByTarget);
            return result;
        }

        private static Dictionary<string, List<CandidatePair>> GroupByAfter(
            IReadOnlyList<CandidatePair> pairs,
            IReadOnlyList<WorkbookSheetFingerprint> after,
            ISet<string> mappedAfter)
        {
            var result = new Dictionary<string, List<CandidatePair>>(StringComparer.Ordinal);
            foreach (WorkbookSheetFingerprint sheet in after)
                if (!mappedAfter.Contains(sheet.SheetKey)) result.Add(sheet.SheetKey, new List<CandidatePair>());
            foreach (CandidatePair pair in pairs) result[pair.After.SheetKey].Add(pair);
            foreach (List<CandidatePair> candidates in result.Values) candidates.Sort(CompareBySource);
            return result;
        }

        private static Dictionary<string, List<WorkbookSheetFingerprint>> GroupByName(
            IReadOnlyList<WorkbookSheetFingerprint> sheets)
        {
            var result = new Dictionary<string, List<WorkbookSheetFingerprint>>(StringComparer.OrdinalIgnoreCase);
            foreach (WorkbookSheetFingerprint sheet in sheets)
            {
                // A malformed OOXML declaration may have no name. Preserve it for diagnostics, but never turn
                // the shared absence of a name into exact-name identity evidence.
                if (string.IsNullOrEmpty(sheet.Name)) continue;
                if (!result.TryGetValue(sheet.Name, out List<WorkbookSheetFingerprint> group))
                {
                    group = new List<WorkbookSheetFingerprint>();
                    result.Add(sheet.Name, group);
                }
                group.Add(sheet);
            }
            return result;
        }

        private static List<WorkbookSheetFingerprint> CopyAndValidate(
            IEnumerable<WorkbookSheetFingerprint> sheets,
            string parameterName)
        {
            if (sheets == null) throw new ArgumentNullException(parameterName);
            var copy = new List<WorkbookSheetFingerprint>();
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (WorkbookSheetFingerprint sheet in sheets)
            {
                if (sheet == null)
                    throw new ArgumentException("Sheet fingerprint collections cannot contain null entries.", parameterName);
                if (!keys.Add(sheet.SheetKey))
                    throw new ArgumentException("Duplicate workbook-local sheet key: " + sheet.SheetKey, parameterName);
                copy.Add(sheet);
            }
            return copy;
        }

        private static bool HasWinnerMargin(IReadOnlyList<CandidatePair> candidates, double minimumMargin)
        {
            if (candidates.Count <= 1) return true;
            return candidates[0].Confidence - candidates[1].Confidence + ComparisonEpsilon >= minimumMargin;
        }

        private static bool MeetsThreshold(double confidence, double threshold) =>
            confidence + ComparisonEpsilon >= threshold;

        private static bool SameToken(string left, string right) =>
            left != null && right != null && string.Equals(left, right, StringComparison.Ordinal);

        private static bool SameIdentity(string left, string right) =>
            left != null && right != null && string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

        private static bool HasSimilarStoredCellCount(long left, long right)
        {
            if (left <= 0 || right <= 0) return false;
            double smaller = Math.Min(left, right);
            double larger = Math.Max(left, right);
            return smaller / larger >= 0.95;
        }

        private static SheetMappingCandidate ToPublicCandidate(CandidatePair pair) =>
            new SheetMappingCandidate(pair.Before, pair.After, pair.Confidence, pair.ReasonCodes);

        private static int CompareSheets(WorkbookSheetFingerprint left, WorkbookSheetFingerprint right) =>
            string.Compare(left.SheetKey, right.SheetKey, StringComparison.Ordinal);

        private static int CompareMappings(ResolvedSheetMapping left, ResolvedSheetMapping right)
        {
            int before = CompareSheets(left.Before, right.Before);
            return before != 0 ? before : CompareSheets(left.After, right.After);
        }

        private static int CompareByTarget(CandidatePair left, CandidatePair right)
        {
            int confidence = right.Confidence.CompareTo(left.Confidence);
            if (confidence != 0) return confidence;
            int name = StringComparer.OrdinalIgnoreCase.Compare(left.After.Name, right.After.Name);
            return name != 0 ? name : CompareSheets(left.After, right.After);
        }

        private static int CompareBySource(CandidatePair left, CandidatePair right)
        {
            int confidence = right.Confidence.CompareTo(left.Confidence);
            if (confidence != 0) return confidence;
            int name = StringComparer.OrdinalIgnoreCase.Compare(left.Before.Name, right.Before.Name);
            return name != 0 ? name : CompareSheets(left.Before, right.Before);
        }

        private sealed class CandidatePair
        {
            internal CandidatePair(
                WorkbookSheetFingerprint before,
                WorkbookSheetFingerprint after,
                double confidence,
                IReadOnlyList<string> reasonCodes)
            {
                Before = before;
                After = after;
                Confidence = confidence;
                ReasonCodes = new ReadOnlyCollection<string>(new List<string>(reasonCodes));
            }

            internal WorkbookSheetFingerprint Before { get; }
            internal WorkbookSheetFingerprint After { get; }
            internal double Confidence { get; }
            internal IReadOnlyList<string> ReasonCodes { get; }
        }
    }
}
