using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Lineage
{
    public enum LineageCorrectionKind
    {
        ReplaceOrder = 0,
        SetParent = 1,
        ClearParent = 2,
        SetMaster = 3,
        ClearMaster = 4,
        SetTrustedBaseline = 5,
        ClearTrustedBaseline = 6,
        SetLabel = 7,
    }

    /// <summary>An immutable requested user correction. Applying it produces a new generation and audit event.</summary>
    public sealed class LineageCorrectionRequest
    {
        private LineageCorrectionRequest(
            LineageCorrectionKind kind,
            string targetRevisionId,
            string value,
            IEnumerable<string> orderedRevisionIds,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note)
        {
            Kind = kind;
            TargetRevisionId = targetRevisionId ?? "";
            Value = value;
            OrderedRevisionIds = new ReadOnlyCollection<string>(
                new List<string>(orderedRevisionIds ?? Array.Empty<string>()));
            Reviewer = reviewer ?? ReviewerSnapshot.Empty;
            CreatedUtc = createdUtc;
            Note = note ?? "";
        }

        public LineageCorrectionKind Kind { get; }
        public string TargetRevisionId { get; }
        public string Value { get; }
        public IReadOnlyList<string> OrderedRevisionIds { get; }
        public ReviewerSnapshot Reviewer { get; }
        public DateTimeOffset CreatedUtc { get; }
        public string Note { get; }

        public static LineageCorrectionRequest ReplaceOrder(
            IEnumerable<string> orderedRevisionIds,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            if (orderedRevisionIds == null) throw new ArgumentNullException(nameof(orderedRevisionIds));
            return new LineageCorrectionRequest(
                LineageCorrectionKind.ReplaceOrder,
                "",
                null,
                orderedRevisionIds,
                reviewer,
                createdUtc,
                note);
        }

        public static LineageCorrectionRequest SetParent(
            string childRevisionId,
            string parentRevisionId,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            RequireRevision(childRevisionId, nameof(childRevisionId));
            RequireRevision(parentRevisionId, nameof(parentRevisionId));
            return new LineageCorrectionRequest(
                LineageCorrectionKind.SetParent,
                childRevisionId,
                parentRevisionId,
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static LineageCorrectionRequest ClearParent(
            string childRevisionId,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            RequireRevision(childRevisionId, nameof(childRevisionId));
            return new LineageCorrectionRequest(
                LineageCorrectionKind.ClearParent,
                childRevisionId,
                null,
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static LineageCorrectionRequest SetMaster(
            string revisionId,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            RequireRevision(revisionId, nameof(revisionId));
            return new LineageCorrectionRequest(
                LineageCorrectionKind.SetMaster,
                revisionId,
                revisionId,
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static LineageCorrectionRequest ClearMaster(
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "") =>
            new LineageCorrectionRequest(
                LineageCorrectionKind.ClearMaster,
                "",
                null,
                null,
                reviewer,
                createdUtc,
                note);

        public static LineageCorrectionRequest SetTrustedBaseline(
            string revisionId,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            RequireRevision(revisionId, nameof(revisionId));
            return new LineageCorrectionRequest(
                LineageCorrectionKind.SetTrustedBaseline,
                revisionId,
                revisionId,
                null,
                reviewer,
                createdUtc,
                note);
        }

        public static LineageCorrectionRequest ClearTrustedBaseline(
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "") =>
            new LineageCorrectionRequest(
                LineageCorrectionKind.ClearTrustedBaseline,
                "",
                null,
                null,
                reviewer,
                createdUtc,
                note);

        public static LineageCorrectionRequest SetLabel(
            string revisionId,
            string label,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note = "")
        {
            RequireRevision(revisionId, nameof(revisionId));
            if (string.IsNullOrWhiteSpace(label))
                throw new ArgumentException("A corrected revision label is required.", nameof(label));
            return new LineageCorrectionRequest(
                LineageCorrectionKind.SetLabel,
                revisionId,
                label,
                null,
                reviewer,
                createdUtc,
                note);
        }

        private static void RequireRevision(string revisionId, string name)
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", name);
        }
    }

    /// <summary>One append-only audit event linking two immutable mapping generations.</summary>
    public sealed class LineageCorrectionEvent
    {
        internal LineageCorrectionEvent(
            string eventId,
            string priorGenerationId,
            string newGenerationId,
            LineageCorrectionKind kind,
            string targetRevisionId,
            string previousValue,
            string newValue,
            IEnumerable<string> previousOrder,
            IEnumerable<string> newOrder,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string note)
        {
            EventId = eventId;
            PriorGenerationId = priorGenerationId;
            NewGenerationId = newGenerationId;
            Kind = kind;
            TargetRevisionId = targetRevisionId ?? "";
            PreviousValue = previousValue;
            NewValue = newValue;
            PreviousOrder = new ReadOnlyCollection<string>(new List<string>(previousOrder ?? Array.Empty<string>()));
            NewOrder = new ReadOnlyCollection<string>(new List<string>(newOrder ?? Array.Empty<string>()));
            Reviewer = reviewer ?? ReviewerSnapshot.Empty;
            CreatedUtc = createdUtc;
            Note = note ?? "";
        }

        public string EventId { get; }
        public string PriorGenerationId { get; }
        public string NewGenerationId { get; }
        public LineageCorrectionKind Kind { get; }
        public string TargetRevisionId { get; }
        public string PreviousValue { get; }
        public string NewValue { get; }
        public IReadOnlyList<string> PreviousOrder { get; }
        public IReadOnlyList<string> NewOrder { get; }
        public ReviewerSnapshot Reviewer { get; }
        public DateTimeOffset CreatedUtc { get; }
        public string Note { get; }
    }

    /// <summary>
    /// Immutable effective mapping after zero or more corrections. Events contain the entire append-only chain;
    /// a new correction never mutates a prior generation.
    /// </summary>
    public sealed class LineageMappingGeneration
    {
        internal LineageMappingGeneration(
            string generationId,
            string priorGenerationId,
            string inferenceProposalGenerationId,
            IEnumerable<string> orderedRevisionIds,
            IDictionary<string, string> parentByRevisionId,
            IDictionary<string, string> labelByRevisionId,
            string masterRevisionId,
            string trustedBaselineRevisionId,
            IEnumerable<LineageCorrectionEvent> events)
        {
            GenerationId = generationId;
            PriorGenerationId = priorGenerationId;
            InferenceProposalGenerationId = inferenceProposalGenerationId;
            OrderedRevisionIds = new ReadOnlyCollection<string>(new List<string>(orderedRevisionIds));
            ParentByRevisionId = new ReadOnlyDictionary<string, string>(
                new Dictionary<string, string>(parentByRevisionId, StringComparer.Ordinal));
            LabelByRevisionId = new ReadOnlyDictionary<string, string>(
                new Dictionary<string, string>(labelByRevisionId, StringComparer.Ordinal));
            MasterRevisionId = masterRevisionId;
            TrustedBaselineRevisionId = trustedBaselineRevisionId;
            Events = new ReadOnlyCollection<LineageCorrectionEvent>(new List<LineageCorrectionEvent>(events));
        }

        public string GenerationId { get; }
        public string PriorGenerationId { get; }
        public string InferenceProposalGenerationId { get; }
        public IReadOnlyList<string> OrderedRevisionIds { get; }
        public IReadOnlyDictionary<string, string> ParentByRevisionId { get; }
        public IReadOnlyDictionary<string, string> LabelByRevisionId { get; }
        public string MasterRevisionId { get; }
        public string TrustedBaselineRevisionId { get; }
        public IReadOnlyList<LineageCorrectionEvent> Events { get; }

        public string GetParent(string revisionId)
        {
            if (!ParentByRevisionId.TryGetValue(revisionId, out string parent))
                throw new KeyNotFoundException("Unknown revision occurrence ID: " + revisionId);
            return parent;
        }

        public string GetLabel(string revisionId)
        {
            if (!LabelByRevisionId.TryGetValue(revisionId, out string label))
                throw new KeyNotFoundException("Unknown revision occurrence ID: " + revisionId);
            return label;
        }
    }

    public static class LineageCorrectionEngine
    {
        public static LineageMappingGeneration FromInference(LineageInferenceResult inference)
        {
            if (inference == null) throw new ArgumentNullException(nameof(inference));
            string[] order = inference.Proposals.Select(proposal => proposal.RevisionId).ToArray();
            var parents = inference.Proposals.ToDictionary(
                proposal => proposal.RevisionId,
                proposal => proposal.ProposedParentRevisionId,
                StringComparer.Ordinal);
            var labels = inference.Proposals.ToDictionary(
                proposal => proposal.RevisionId,
                proposal => proposal.LabelInference.BaseName,
                StringComparer.Ordinal);

            var components = new List<string> { inference.ProposalGenerationId };
            AddListComponents(components, order);
            foreach (string revisionId in order)
            {
                components.Add(revisionId);
                components.Add(parents[revisionId]);
                components.Add(labels[revisionId]);
            }
            components.Add(inference.ExplicitMasterRevisionId);
            components.Add(inference.ExplicitTrustedBaselineRevisionId);
            string generationId = StableId.Compute("lineage-mapping-generation-initial-v1", components.ToArray());

            return new LineageMappingGeneration(
                generationId,
                null,
                inference.ProposalGenerationId,
                order,
                parents,
                labels,
                inference.ExplicitMasterRevisionId,
                inference.ExplicitTrustedBaselineRevisionId,
                Array.Empty<LineageCorrectionEvent>());
        }

        public static LineageMappingGeneration Apply(
            LineageMappingGeneration current,
            LineageCorrectionRequest correction)
        {
            if (current == null) throw new ArgumentNullException(nameof(current));
            if (correction == null) throw new ArgumentNullException(nameof(correction));

            var order = new List<string>(current.OrderedRevisionIds);
            var parents = current.ParentByRevisionId.ToDictionary(
                pair => pair.Key,
                pair => pair.Value,
                StringComparer.Ordinal);
            var labels = current.LabelByRevisionId.ToDictionary(
                pair => pair.Key,
                pair => pair.Value,
                StringComparer.Ordinal);
            string master = current.MasterRevisionId;
            string baseline = current.TrustedBaselineRevisionId;
            string previousValue = null;
            string newValue = null;
            IReadOnlyList<string> previousOrder = Array.Empty<string>();
            IReadOnlyList<string> newOrder = Array.Empty<string>();

            switch (correction.Kind)
            {
                case LineageCorrectionKind.ReplaceOrder:
                    ValidateCompleteOrder(correction.OrderedRevisionIds, order);
                    if (order.SequenceEqual(correction.OrderedRevisionIds, StringComparer.Ordinal))
                        throw new InvalidOperationException("The corrected order is unchanged.");
                    ValidateParentsPrecedeChildren(correction.OrderedRevisionIds, parents);
                    previousOrder = new ReadOnlyCollection<string>(new List<string>(order));
                    order = new List<string>(correction.OrderedRevisionIds);
                    newOrder = new ReadOnlyCollection<string>(new List<string>(order));
                    break;

                case LineageCorrectionKind.SetParent:
                    RequireKnown(correction.TargetRevisionId, parents, "child");
                    RequireKnown(correction.Value, parents, "parent");
                    ValidateParentPrecedesChild(correction.TargetRevisionId, correction.Value, order);
                    previousValue = parents[correction.TargetRevisionId];
                    newValue = correction.Value;
                    EnsureChanged(previousValue, newValue, "parent");
                    parents[correction.TargetRevisionId] = newValue;
                    break;

                case LineageCorrectionKind.ClearParent:
                    RequireKnown(correction.TargetRevisionId, parents, "child");
                    previousValue = parents[correction.TargetRevisionId];
                    newValue = null;
                    EnsureChanged(previousValue, newValue, "parent");
                    parents[correction.TargetRevisionId] = null;
                    break;

                case LineageCorrectionKind.SetMaster:
                    RequireKnown(correction.Value, parents, "master");
                    previousValue = master;
                    newValue = correction.Value;
                    EnsureChanged(previousValue, newValue, "master");
                    master = newValue;
                    break;

                case LineageCorrectionKind.ClearMaster:
                    previousValue = master;
                    newValue = null;
                    EnsureChanged(previousValue, newValue, "master");
                    master = null;
                    break;

                case LineageCorrectionKind.SetTrustedBaseline:
                    RequireKnown(correction.Value, parents, "trusted baseline");
                    previousValue = baseline;
                    newValue = correction.Value;
                    EnsureChanged(previousValue, newValue, "trusted baseline");
                    baseline = newValue;
                    break;

                case LineageCorrectionKind.ClearTrustedBaseline:
                    previousValue = baseline;
                    newValue = null;
                    EnsureChanged(previousValue, newValue, "trusted baseline");
                    baseline = null;
                    break;

                case LineageCorrectionKind.SetLabel:
                    RequireKnown(correction.TargetRevisionId, parents, "revision");
                    previousValue = labels[correction.TargetRevisionId];
                    newValue = correction.Value;
                    EnsureChanged(previousValue, newValue, "label");
                    labels[correction.TargetRevisionId] = newValue;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(correction), "Unknown lineage correction kind.");
            }

            string eventId = ComputeEventId(
                current.GenerationId,
                correction,
                previousValue,
                newValue,
                previousOrder,
                newOrder);
            string newGenerationId = StableId.Compute(
                "lineage-mapping-generation-correction-v1",
                current.GenerationId,
                eventId);
            var correctionEvent = new LineageCorrectionEvent(
                eventId,
                current.GenerationId,
                newGenerationId,
                correction.Kind,
                correction.TargetRevisionId,
                previousValue,
                newValue,
                previousOrder,
                newOrder,
                correction.Reviewer,
                correction.CreatedUtc,
                correction.Note);
            var events = new List<LineageCorrectionEvent>(current.Events) { correctionEvent };

            return new LineageMappingGeneration(
                newGenerationId,
                current.GenerationId,
                current.InferenceProposalGenerationId,
                order,
                parents,
                labels,
                master,
                baseline,
                events);
        }

        private static string ComputeEventId(
            string priorGenerationId,
            LineageCorrectionRequest correction,
            string previousValue,
            string newValue,
            IReadOnlyList<string> previousOrder,
            IReadOnlyList<string> newOrder)
        {
            var components = new List<string>
            {
                priorGenerationId,
                Convert.ToInt32(correction.Kind, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                correction.TargetRevisionId,
                previousValue,
                newValue,
            };
            AddListComponents(components, previousOrder);
            AddListComponents(components, newOrder);
            components.Add(correction.Reviewer.ConfirmedDisplayName);
            components.Add(correction.Reviewer.OfficeDisplayName);
            AddListComponents(components, correction.Reviewer.IdentitySourceLabels);
            components.Add(correction.CreatedUtc.UtcDateTime.ToString("O", CultureInfo.InvariantCulture));
            components.Add(correction.Note);
            return StableId.Compute("lineage-correction-event-v1", components.ToArray());
        }

        private static void ValidateCompleteOrder(
            IReadOnlyList<string> proposed,
            IReadOnlyList<string> current)
        {
            if (proposed.Count != current.Count)
                throw new ArgumentException("A corrected order must contain every revision occurrence exactly once.");
            var expected = new HashSet<string>(current, StringComparer.Ordinal);
            var actual = new HashSet<string>(StringComparer.Ordinal);
            foreach (string revisionId in proposed)
            {
                if (string.IsNullOrWhiteSpace(revisionId) || !expected.Contains(revisionId) || !actual.Add(revisionId))
                    throw new ArgumentException("A corrected order must contain every revision occurrence exactly once.");
            }
        }

        private static void ValidateParentsPrecedeChildren(
            IReadOnlyList<string> order,
            IDictionary<string, string> parents)
        {
            var index = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < order.Count; i++) index[order[i]] = i;
            foreach (KeyValuePair<string, string> pair in parents)
            {
                if (pair.Value == null) continue;
                if (index[pair.Value] >= index[pair.Key])
                    throw new InvalidOperationException(
                        "The corrected order would place parent '" + pair.Value + "' at or after child '" +
                        pair.Key + "'. Clear or reconnect that parent first.");
            }
        }

        private static void ValidateParentPrecedesChild(
            string childRevisionId,
            string parentRevisionId,
            IReadOnlyList<string> order)
        {
            int child = IndexOf(order, childRevisionId);
            int parent = IndexOf(order, parentRevisionId);
            if (parent >= child)
                throw new InvalidOperationException("A parent must precede its child in the current mapping order.");
        }

        private static int IndexOf(IReadOnlyList<string> values, string value)
        {
            for (int i = 0; i < values.Count; i++)
                if (string.Equals(values[i], value, StringComparison.Ordinal)) return i;
            return -1;
        }

        private static void RequireKnown(
            string revisionId,
            IDictionary<string, string> revisions,
            string role)
        {
            if (revisionId == null || !revisions.ContainsKey(revisionId))
                throw new ArgumentException("The corrected " + role + " is not an imported revision: " + revisionId);
        }

        private static void EnsureChanged(string before, string after, string field)
        {
            if (string.Equals(before, after, StringComparison.Ordinal))
                throw new InvalidOperationException("The corrected " + field + " is unchanged.");
        }

        private static void AddListComponents(ICollection<string> components, IEnumerable<string> values)
        {
            string[] copy = (values ?? Array.Empty<string>()).ToArray();
            components.Add(copy.Length.ToString(CultureInfo.InvariantCulture));
            foreach (string value in copy) components.Add(value);
        }
    }
}
