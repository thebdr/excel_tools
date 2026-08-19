using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Project
{
    public static class AlignmentCorrectionReconciliationCodes
    {
        public const string KeyPairNotAccepted = "ALN-CORRECTION-KEY-PAIR-NOT-ACCEPTED";
    }

    /// <summary>One structured, auditable safety action applied before project analysis.</summary>
    public sealed class AlignmentCorrectionReconciliationAction
    {
        internal AlignmentCorrectionReconciliationAction(
            string reasonCode,
            AlignmentCorrectionEvent correctionEvent)
        {
            if (string.IsNullOrWhiteSpace(reasonCode))
                throw new ArgumentException("A reconciliation reason code is required.", nameof(reasonCode));
            CorrectionEvent = correctionEvent ?? throw new ArgumentNullException(nameof(correctionEvent));
            ReasonCode = reasonCode;
        }

        public string ReasonCode { get; }
        public AlignmentCorrectionEvent CorrectionEvent { get; }
        public AlignmentCorrectionKind Kind => CorrectionEvent.Kind;
        public string EventId => CorrectionEvent.EventId;
        public string RequestId => CorrectionEvent.RequestId;
        public string SupersededConfirmationId => CorrectionEvent.SupersededConfirmationId;
        public AlignmentCorrectionScope Scope => CorrectionEvent.Scope;
    }

    public sealed class AlignmentCorrectionReconciliationResult
    {
        internal AlignmentCorrectionReconciliationResult(
            AlignmentCorrectionGeneration generation,
            IEnumerable<AlignmentCorrectionReconciliationAction> actions)
        {
            Generation = generation ?? throw new ArgumentNullException(nameof(generation));
            Actions = new ReadOnlyCollection<AlignmentCorrectionReconciliationAction>(
                new List<AlignmentCorrectionReconciliationAction>(
                    actions ?? Array.Empty<AlignmentCorrectionReconciliationAction>()));
        }

        public AlignmentCorrectionGeneration Generation { get; }
        public IReadOnlyList<AlignmentCorrectionReconciliationAction> Actions { get; }
    }

    /// <summary>
    /// Reconciles exact pair-scoped key overrides against the final sheet mappings produced by the current
    /// sheet-pair and deleted/new confirmations. It never redirects a key by name/position: a now-unconsumable key is returned
    /// to automatic inference through the normal append-only correction-event path. Header confirmations are
    /// revision/physical-sheet scoped and deliberately survive remaps, including while a sheet is unresolved.
    /// </summary>
    public static class AlignmentCorrectionReconciler
    {
        public static AlignmentCorrectionReconciliationResult Reconcile(
            IEnumerable<OpenXmlProjectRevision> revisions,
            AlignmentCorrectionGeneration corrections,
            WorkbookSheetMappingOptions mappingOptions,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));
            if (corrections == null) throw new ArgumentNullException(nameof(corrections));
            mappingOptions = mappingOptions ?? WorkbookSheetMappingOptions.Default;
            reviewer = reviewer ?? ReviewerSnapshot.Empty;
            if (corrections.SheetPairs.Count == 0 && corrections.UnmatchedSheets.Count == 0)
                return new AlignmentCorrectionReconciliationResult(
                    corrections,
                    Array.Empty<AlignmentCorrectionReconciliationAction>());

            List<OpenXmlProjectRevision> ordered = revisions.ToList();
            if (ordered.Count == 0)
                throw new ArgumentException("At least one project revision is required.", nameof(revisions));
            if (ordered.Any(item => item == null))
                throw new ArgumentException("Project revisions cannot contain null entries.", nameof(revisions));
            ordered.Sort((left, right) => left.Order.CompareTo(right.Order));
            var byId = new Dictionary<string, OpenXmlProjectRevision>(StringComparer.Ordinal);
            foreach (OpenXmlProjectRevision revision in ordered)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (byId.ContainsKey(revision.RevisionId))
                    throw new ArgumentException(
                        "Project revision IDs must be unique for correction reconciliation.",
                        nameof(revisions));
                byId.Add(revision.RevisionId, revision);
            }

            var invalidKeys = new List<AlignmentKeyColumnsConfirmation>();

            foreach (OpenXmlProjectRevision child in ordered)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrEmpty(child.ParentRevisionId)) continue;
                if (!byId.TryGetValue(child.ParentRevisionId, out OpenXmlProjectRevision parent))
                    throw new ArgumentException(
                        "A project revision references an unknown parent during correction reconciliation.",
                        nameof(revisions));

                OpenXmlWorkbookSheetMappingResult inferred = OpenXmlWorkbookSheetMappingInference.Infer(
                    parent.Snapshot,
                    child.Snapshot,
                    mappingOptions,
                    cancellationToken);
                OpenXmlWorkbookSheetMappingResult effective =
                    OpenXmlComparisonProjectEngine.ApplyConfirmedSheetMappings(
                        inferred,
                        corrections.GetSheetPairs(parent.RevisionId, child.RevisionId),
                        corrections.GetUnmatchedSheets(parent.RevisionId, child.RevisionId));

                var effectiveEdgePairs = new HashSet<SheetPairIdentity>(effective.Mapping.Mappings.Select(item =>
                    new SheetPairIdentity(
                        parent.RevisionId,
                        child.RevisionId,
                        item.Before.SheetKey,
                        item.After.SheetKey)));
                invalidKeys.AddRange(corrections.KeyColumns.Where(item =>
                    string.Equals(item.ParentRevisionId, parent.RevisionId, StringComparison.Ordinal) &&
                    string.Equals(item.ChildRevisionId, child.RevisionId, StringComparison.Ordinal) &&
                    !effectiveEdgePairs.Contains(new SheetPairIdentity(
                        item.ParentRevisionId,
                        item.ChildRevisionId,
                        item.BeforeSheetKey,
                        item.AfterSheetKey))));
            }

            cancellationToken.ThrowIfCancellationRequested();
            AlignmentCorrectionGeneration generation = corrections;
            var actions = new List<AlignmentCorrectionReconciliationAction>();
            foreach (AlignmentKeyColumnsConfirmation key in invalidKeys
                         .OrderBy(item => item.ParentRevisionId, StringComparer.Ordinal)
                         .ThenBy(item => item.ChildRevisionId, StringComparer.Ordinal)
                         .ThenBy(item => item.BeforeSheetKey, StringComparer.Ordinal)
                         .ThenBy(item => item.AfterSheetKey, StringComparer.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string note = AlignmentCorrectionReconciliationCodes.KeyPairNotAccepted +
                              ": active key confirmation '" + key.ConfirmationId +
                              "' targeted exact sheet pair '" + key.BeforeSheetKey + "' -> '" +
                              key.AfterSheetKey + "' on edge '" + key.ParentRevisionId + "' -> '" +
                              key.ChildRevisionId +
                              "', which is not accepted after applying the current sheet remaps. The key " +
                              "scope was returned to automatic inference before analysis; no key was redirected.";
                AlignmentCorrectionRequest request = AlignmentCorrectionRequest.ClearKeyColumns(
                    key.ParentRevisionId,
                    key.ChildRevisionId,
                    key.BeforeSheetKey,
                    key.AfterSheetKey,
                    reviewer,
                    createdUtc,
                    note);
                generation = AlignmentCorrectionEngine.Apply(generation, request);
                actions.Add(new AlignmentCorrectionReconciliationAction(
                    AlignmentCorrectionReconciliationCodes.KeyPairNotAccepted,
                    generation.Events[generation.Events.Count - 1]));
            }

            return new AlignmentCorrectionReconciliationResult(generation, actions);
        }

        private sealed class SheetPairIdentity : IEquatable<SheetPairIdentity>
        {
            internal SheetPairIdentity(
                string parentRevisionId,
                string childRevisionId,
                string beforeSheetKey,
                string afterSheetKey)
            {
                ParentRevisionId = parentRevisionId ?? "";
                ChildRevisionId = childRevisionId ?? "";
                BeforeSheetKey = beforeSheetKey ?? "";
                AfterSheetKey = afterSheetKey ?? "";
            }

            internal string ParentRevisionId { get; }
            internal string ChildRevisionId { get; }
            internal string BeforeSheetKey { get; }
            internal string AfterSheetKey { get; }

            public bool Equals(SheetPairIdentity other) =>
                other != null &&
                string.Equals(ParentRevisionId, other.ParentRevisionId, StringComparison.Ordinal) &&
                string.Equals(ChildRevisionId, other.ChildRevisionId, StringComparison.Ordinal) &&
                string.Equals(BeforeSheetKey, other.BeforeSheetKey, StringComparison.Ordinal) &&
                string.Equals(AfterSheetKey, other.AfterSheetKey, StringComparison.Ordinal);

            public override bool Equals(object obj) => Equals(obj as SheetPairIdentity);

            public override int GetHashCode()
            {
                unchecked
                {
                    int hash = StringComparer.Ordinal.GetHashCode(ParentRevisionId);
                    hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(ChildRevisionId);
                    hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(BeforeSheetKey);
                    return (hash * 397) ^ StringComparer.Ordinal.GetHashCode(AfterSheetKey);
                }
            }
        }

    }
}
