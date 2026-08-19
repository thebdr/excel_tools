using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Diff
{
    /// <summary>
    /// Deterministic exact comparison over already-aligned logical subjects. Alignment, grouping, equivalence,
    /// and severity are deliberately separate projections and cannot remove these atomic facts.
    /// </summary>
    public static class ExactComparisonEngine
    {
        public static IReadOnlyList<AtomicCellChange> Compare(
            RevisionSnapshot before,
            RevisionSnapshot after,
            ExactComparisonOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (before == null) throw new ArgumentNullException(nameof(before));
            if (after == null) throw new ArgumentNullException(nameof(after));
            options = options ?? new ExactComparisonOptions();

            var subjects = new SortedSet<string>(StringComparer.Ordinal);
            foreach (string key in before.Cells.Keys) subjects.Add(key);
            foreach (string key in after.Cells.Keys) subjects.Add(key);

            var changes = new List<AtomicCellChange>();
            foreach (string subject in subjects)
            {
                cancellationToken.ThrowIfCancellationRequested();

                CanonicalCellState beforeCell;
                CanonicalCellState afterCell;
                bool existedBefore = before.TryGetCell(subject, out beforeCell);
                bool existsAfter = after.TryGetCell(subject, out afterCell);
                CanonicalCellState beforeEffective = existedBefore ? beforeCell : CanonicalCellState.Blank(subject);
                CanonicalCellState afterEffective = existsAfter ? afterCell : CanonicalCellState.Blank(subject);
                int countBefore = changes.Count;

                CompareFormula(subject, beforeEffective, afterEffective, existedBefore, existsAfter, changes);
                CompareFormulaStructure(subject, beforeEffective, afterEffective, existedBefore, existsAfter, changes);
                CompareValue(subject, beforeEffective, afterEffective, existedBefore, existsAfter, changes);

                bool formulaChanged =
                    !string.Equals(
                        beforeEffective.FormulaText,
                        afterEffective.FormulaText,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        beforeEffective.FormulaMetadata,
                        afterEffective.FormulaMetadata,
                        StringComparison.Ordinal);
                if (options.CompareCachedResults && (!formulaChanged || options.CompareCachedResultsWhenFormulaChanges))
                    CompareCachedResult(subject, beforeEffective, afterEffective, existedBefore, existsAfter, changes);

                if (options.ComparePhysicalPresence && existedBefore != existsAfter && changes.Count == countBefore)
                {
                    changes.Add(new AtomicCellChange(
                        subject,
                        CellChangeAspect.Presence,
                        existsAfter ? CellChangeKind.Added : CellChangeKind.Removed,
                        beforeCell,
                        afterCell,
                        existedBefore ? "present" : "missing",
                        existsAfter ? "present" : "missing",
                        null,
                        null));
                }
            }

            return new ReadOnlyCollection<AtomicCellChange>(changes);
        }

        private static void CompareFormulaStructure(
            string subject,
            CanonicalCellState before,
            CanonicalCellState after,
            bool existedBefore,
            bool existsAfter,
            List<AtomicCellChange> changes)
        {
            if (string.Equals(before.FormulaMetadata, after.FormulaMetadata, StringComparison.Ordinal)) return;

            changes.Add(new AtomicCellChange(
                subject,
                CellChangeAspect.FormulaStructure,
                ChangeKind(
                    before.FormulaMetadata != null,
                    after.FormulaMetadata != null,
                    existedBefore,
                    existsAfter),
                existedBefore ? before : null,
                existsAfter ? after : null,
                before.FormulaMetadata,
                after.FormulaMetadata,
                null,
                null));
        }

        private static void CompareFormula(
            string subject,
            CanonicalCellState before,
            CanonicalCellState after,
            bool existedBefore,
            bool existsAfter,
            List<AtomicCellChange> changes)
        {
            if (string.Equals(before.FormulaText, after.FormulaText, StringComparison.Ordinal)) return;

            changes.Add(new AtomicCellChange(
                subject,
                CellChangeAspect.Formula,
                ChangeKind(before.FormulaText != null, after.FormulaText != null, existedBefore, existsAfter),
                existedBefore ? before : null,
                existsAfter ? after : null,
                before.FormulaText,
                after.FormulaText,
                null,
                null));
        }

        private static void CompareValue(
            string subject,
            CanonicalCellState before,
            CanonicalCellState after,
            bool existedBefore,
            bool existsAfter,
            List<AtomicCellChange> changes)
        {
            if (before.ConstantValue.Equals(after.ConstantValue)) return;

            bool beforeHasValue = before.ConstantValue.Kind != CanonicalValueKind.Blank;
            bool afterHasValue = after.ConstantValue.Kind != CanonicalValueKind.Blank;
            changes.Add(new AtomicCellChange(
                subject,
                CellChangeAspect.Value,
                ChangeKind(beforeHasValue, afterHasValue, existedBefore, existsAfter),
                existedBefore ? before : null,
                existsAfter ? after : null,
                before.ConstantValue.ToString(),
                after.ConstantValue.ToString(),
                before.ConstantValue,
                after.ConstantValue));
        }

        private static void CompareCachedResult(
            string subject,
            CanonicalCellState before,
            CanonicalCellState after,
            bool existedBefore,
            bool existsAfter,
            List<AtomicCellChange> changes)
        {
            CanonicalValue beforeCache = before.CachedResult;
            CanonicalValue afterCache = after.CachedResult;
            if (ReferenceEquals(beforeCache, afterCache)) return;
            if (beforeCache != null && beforeCache.Equals(afterCache)) return;

            changes.Add(new AtomicCellChange(
                subject,
                CellChangeAspect.CachedResult,
                ChangeKind(beforeCache != null, afterCache != null, existedBefore, existsAfter),
                existedBefore ? before : null,
                existsAfter ? after : null,
                beforeCache == null ? null : beforeCache.ToString(),
                afterCache == null ? null : afterCache.ToString(),
                beforeCache,
                afterCache));
        }

        private static CellChangeKind ChangeKind(
            bool hadAspectBefore,
            bool hasAspectAfter,
            bool existedBefore,
            bool existsAfter)
        {
            if ((!existedBefore || !hadAspectBefore) && existsAfter && hasAspectAfter) return CellChangeKind.Added;
            if (existedBefore && hadAspectBefore && (!existsAfter || !hasAspectAfter)) return CellChangeKind.Removed;
            return CellChangeKind.Modified;
        }
    }
}
