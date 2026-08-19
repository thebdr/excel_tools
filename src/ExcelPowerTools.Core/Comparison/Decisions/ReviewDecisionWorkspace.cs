using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>
    /// Bounded, UI-safe rendering of one exact observation. Exact identities remain available for decision
    /// requests; display strings are deliberately capped so a virtualized review list cannot duplicate very large
    /// formula or cell payloads.
    /// </summary>
    public sealed class ReviewDecisionObservationChoice
    {
        internal ReviewDecisionObservationChoice(
            Observation observation,
            bool isMaster,
            int revisionOrder)
        {
            if (observation == null) throw new ArgumentNullException(nameof(observation));
            ObservationId = observation.ObservationId;
            RevisionId = observation.RevisionId;
            RevisionOrder = revisionOrder;
            IsMaster = isMaster;
            ExactType = observation.ExactType;
            SourceLabel = observation.Provenance.SourceLabel;
            SheetName = observation.Provenance.SheetName;
            Address = observation.Provenance.Address;
            PhysicalRow = observation.Provenance.PhysicalRow;
            PhysicalColumn = observation.Provenance.PhysicalColumn;
            Content = DisplayContent(observation);
            Location = DisplayLocation(observation.Provenance);
        }

        public string ObservationId { get; }
        public string RevisionId { get; }
        public int RevisionOrder { get; }
        public bool IsMaster { get; }
        public ObservationExactType ExactType { get; }
        public string SourceLabel { get; }
        public string SheetName { get; }
        public string Address { get; }
        public int PhysicalRow { get; }
        public int PhysicalColumn { get; }
        public string Content { get; }
        public string Location { get; }

        public string RevisionLabel
        {
            get
            {
                string label = string.IsNullOrWhiteSpace(SourceLabel)
                    ? RevisionId
                    : SourceLabel;
                return IsMaster ? label + "  [master]" : label;
            }
        }

        private static string DisplayContent(Observation observation)
        {
            string text;
            switch (observation.ExactType)
            {
                case ObservationExactType.Formula:
                    text = observation.Formula.StartsWith("=", StringComparison.Ordinal)
                        ? observation.Formula
                        : "=" + observation.Formula;
                    break;
                case ObservationExactType.Missing:
                    text = "(missing)";
                    break;
                case ObservationExactType.Blank:
                    text = "(blank)";
                    break;
                default:
                    text = observation.Value?.CanonicalText ?? observation.DisplayText ?? "";
                    if (text.Length == 0 && observation.ExactType == ObservationExactType.Unsupported)
                        text = "(unsupported)";
                    break;
            }
            return CapSingleLine(text, 240);
        }

        private static string DisplayLocation(ObservationProvenance provenance)
        {
            if (provenance == null) return "";
            if (provenance.SheetName.Length != 0 && provenance.Address.Length != 0)
                return provenance.SheetName + "!" + provenance.Address;
            if (provenance.SheetName.Length != 0) return provenance.SheetName;
            return provenance.Address;
        }

        internal static string CapSingleLine(string value, int maxCharacters)
        {
            string normalized = (value ?? "")
                .Replace("\r\n", " ")
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
            return normalized.Length <= maxCharacters
                ? normalized
                : normalized.Substring(0, maxCharacters - 1) + "…";
        }
    }

    /// <summary>One independently reviewable subject plus its exact ordered revision choices.</summary>
    public sealed class ReviewDecisionWorkspaceRow
    {
        internal ReviewDecisionWorkspaceRow(
            ReviewDecisionSubject subject,
            ReviewDecisionState state,
            IEnumerable<ReviewDecisionObservationChoice> observations,
            IEnumerable<string> groupIds)
        {
            if (subject == null) throw new ArgumentNullException(nameof(subject));
            if (state == null) throw new ArgumentNullException(nameof(state));
            SubjectId = subject.SubjectId;
            LogicalSubjectKey = subject.LogicalSubjectKey;
            Aspect = subject.Aspect;
            Policy = subject.Policy;
            Status = state.Status;
            EffectiveEventId = state.EffectiveEventId;
            SelectedObservationId = state.SelectedObservationId;

            var frozenObservations = new List<ReviewDecisionObservationChoice>(
                observations ?? Array.Empty<ReviewDecisionObservationChoice>());
            if (frozenObservations.Any(item => item == null))
                throw new ArgumentException("Workspace observations cannot contain null entries.", nameof(observations));
            Observations = new ReadOnlyCollection<ReviewDecisionObservationChoice>(frozenObservations);
            Master = frozenObservations.SingleOrDefault(item => item.IsMaster);
            Candidates = new ReadOnlyCollection<ReviewDecisionObservationChoice>(
                frozenObservations.Where(item => !item.IsMaster).ToList());

            var frozenGroups = new List<string>(groupIds ?? Array.Empty<string>());
            frozenGroups.Sort(StringComparer.Ordinal);
            GroupIds = new ReadOnlyCollection<string>(frozenGroups);

            ReviewDecisionObservationChoice location =
                Master != null && Master.Location.Length != 0
                    ? Master
                    :
                frozenObservations.LastOrDefault(item => item.Location.Length != 0);
            Location = string.IsNullOrEmpty(location?.Location)
                ? LogicalSubjectKey
                : location.Location;
            MasterContent = Master?.Content ?? "(unavailable)";
            ReviewDecisionObservationChoice latestCandidate = Candidates.LastOrDefault();
            LatestCandidateContent = latestCandidate?.Content ?? "(none)";
            LatestCandidateRevisionLabel = latestCandidate?.RevisionLabel ?? "";
        }

        public string SubjectId { get; }
        public string LogicalSubjectKey { get; }
        public ComparisonAspect Aspect { get; }
        public ReviewDecisionSubjectPolicy Policy { get; }
        public ReviewDecisionStatus Status { get; }
        public string EffectiveEventId { get; }
        public string SelectedObservationId { get; }
        public IReadOnlyList<ReviewDecisionObservationChoice> Observations { get; }
        public ReviewDecisionObservationChoice Master { get; }
        public IReadOnlyList<ReviewDecisionObservationChoice> Candidates { get; }
        public IReadOnlyList<string> GroupIds { get; }
        public string Location { get; }
        public string MasterContent { get; }
        public string LatestCandidateContent { get; }
        public string LatestCandidateRevisionLabel { get; }
        public bool CanSelectObservation =>
            Policy == ReviewDecisionSubjectPolicy.Selectable ||
            Policy == ReviewDecisionSubjectPolicy.SelectableAndMergeable;
    }

    /// <summary>
    /// Stable virtual-list projection of one hydrated decision generation. It retains no paths and grants no new
    /// authority beyond the exact catalog/generation identities supplied to the request builder.
    /// </summary>
    public sealed class ReviewDecisionWorkspaceProjection
    {
        private readonly IReadOnlyDictionary<string, ReviewDecisionWorkspaceRow> _rowsBySubjectId;

        private ReviewDecisionWorkspaceProjection(
            ReviewDecisionGeneration generation,
            IEnumerable<ReviewDecisionWorkspaceRow> rows)
        {
            GenerationId = generation.GenerationId;
            CatalogId = generation.CatalogId;
            ProjectId = generation.ProjectId;
            MasterRevisionId = generation.Catalog.MasterRevisionId;
            var frozen = new List<ReviewDecisionWorkspaceRow>(rows);
            Rows = new ReadOnlyCollection<ReviewDecisionWorkspaceRow>(frozen);
            _rowsBySubjectId = new ReadOnlyDictionary<string, ReviewDecisionWorkspaceRow>(
                frozen.ToDictionary(item => item.SubjectId, StringComparer.Ordinal));
        }

        public string GenerationId { get; }
        public string CatalogId { get; }
        public string ProjectId { get; }
        public string MasterRevisionId { get; }
        public IReadOnlyList<ReviewDecisionWorkspaceRow> Rows { get; }

        public bool TryGetRow(string subjectId, out ReviewDecisionWorkspaceRow row) =>
            _rowsBySubjectId.TryGetValue(subjectId ?? "", out row);

        public static ReviewDecisionWorkspaceProjection Build(ReviewDecisionGeneration generation)
        {
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            ReviewDecisionCatalog catalog = generation.Catalog;
            var revisionOrder = BuildRevisionOrder(catalog);
            var groupsBySubject = BuildGroupsBySubject(catalog);
            IReadOnlyDictionary<string, List<Observation>> observationsBySubject =
                catalog.Observations
                    .GroupBy(item => item.SubjectId, StringComparer.Ordinal)
                    .ToDictionary(
                        group => group.Key,
                        group => group.ToList(),
                        StringComparer.Ordinal);
            var rows = new List<ReviewDecisionWorkspaceRow>(catalog.Subjects.Count);
            foreach (ReviewDecisionSubject subject in catalog.Subjects)
            {
                IEnumerable<Observation> subjectObservations =
                    observationsBySubject.TryGetValue(subject.SubjectId, out List<Observation> retained)
                        ? retained
                        : Array.Empty<Observation>();
                var observations = subjectObservations
                    .Select(item => new ReviewDecisionObservationChoice(
                        item,
                        string.Equals(item.RevisionId, catalog.MasterRevisionId, StringComparison.Ordinal),
                        revisionOrder.TryGetValue(item.RevisionId, out int order) ? order : int.MaxValue))
                    .OrderBy(item => item.RevisionOrder)
                    .ThenBy(item => item.RevisionId, StringComparer.Ordinal)
                    .ThenBy(item => item.ObservationId, StringComparer.Ordinal)
                    .ToList();
                rows.Add(new ReviewDecisionWorkspaceRow(
                    subject,
                    generation.GetSubjectState(subject.SubjectId),
                    observations,
                    groupsBySubject.TryGetValue(subject.SubjectId, out List<string> groupIds)
                        ? groupIds
                        : Array.Empty<string>()));
            }
            rows.Sort(CompareRows);
            return new ReviewDecisionWorkspaceProjection(generation, rows);
        }

        private static Dictionary<string, int> BuildRevisionOrder(ReviewDecisionCatalog catalog)
        {
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            if (catalog.EvidenceGeneration != null)
            {
                foreach (ReviewEvidenceRevisionIdentity revision in catalog.EvidenceGeneration.Revisions)
                    result.Add(revision.RevisionId, revision.Order);
                return result;
            }
            int order = 0;
            foreach (string revisionId in catalog.Observations
                         .Select(item => item.RevisionId)
                         .Distinct(StringComparer.Ordinal)
                         .OrderBy(item => item, StringComparer.Ordinal))
                result.Add(revisionId, order++);
            return result;
        }

        private static Dictionary<string, List<string>> BuildGroupsBySubject(ReviewDecisionCatalog catalog)
        {
            var result = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            foreach (ReviewDecisionGroup group in catalog.Groups)
            foreach (string subjectId in group.SubjectIds)
            {
                if (!result.TryGetValue(subjectId, out List<string> groups))
                {
                    groups = new List<string>();
                    result.Add(subjectId, groups);
                }
                groups.Add(group.GroupId);
            }
            return result;
        }

        private static int CompareRows(ReviewDecisionWorkspaceRow left, ReviewDecisionWorkspaceRow right)
        {
            ReviewDecisionObservationChoice leftLocation = SortLocation(left);
            ReviewDecisionObservationChoice rightLocation = SortLocation(right);
            int value = StringComparer.OrdinalIgnoreCase.Compare(
                leftLocation?.SheetName ?? "",
                rightLocation?.SheetName ?? "");
            if (value != 0) return value;
            value = PositiveOrMaximum(leftLocation?.PhysicalRow ?? 0)
                .CompareTo(PositiveOrMaximum(rightLocation?.PhysicalRow ?? 0));
            if (value != 0) return value;
            value = PositiveOrMaximum(leftLocation?.PhysicalColumn ?? 0)
                .CompareTo(PositiveOrMaximum(rightLocation?.PhysicalColumn ?? 0));
            if (value != 0) return value;
            value = ((int)left.Aspect).CompareTo((int)right.Aspect);
            return value != 0
                ? value
                : StringComparer.Ordinal.Compare(left.SubjectId, right.SubjectId);
        }

        private static ReviewDecisionObservationChoice SortLocation(ReviewDecisionWorkspaceRow row) =>
            row.Master != null && row.Master.Location.Length != 0
                ? row.Master
                : row.Observations.LastOrDefault(item => item.Location.Length != 0);

        private static int PositiveOrMaximum(int value) => value > 0 ? value : int.MaxValue;
    }

    /// <summary>Builds optimistic-concurrency requests from an exact workspace selection.</summary>
    public static class ReviewDecisionWorkspaceRequestBuilder
    {
        public static ReviewDecisionRequest ForSubjects(
            ReviewDecisionGeneration generation,
            IEnumerable<string> subjectIds,
            ReviewDecisionStatus status,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string selectedObservationId = "",
            string runId = "",
            string addInBuild = "",
            string note = "")
        {
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            List<string> frozen = FreezeKnownSubjects(generation, subjectIds);
            if (frozen.Count == 1)
            {
                return ReviewDecisionRequest.ForSubject(
                    generation.ProjectId,
                    generation.BaseGenerationId,
                    generation.MappingGenerationId,
                    generation.GenerationId,
                    frozen[0],
                    status,
                    reviewer,
                    createdUtc,
                    selectedObservationId,
                    runId,
                    addInBuild,
                    note);
            }
            return ReviewDecisionRequest.ForSubjects(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.GenerationId,
                frozen,
                status,
                reviewer,
                createdUtc,
                selectedObservationId,
                runId,
                addInBuild,
                note);
        }

        public static ReviewDecisionRequest ForGroup(
            ReviewDecisionGeneration generation,
            string groupId,
            ReviewDecisionStatus status,
            ReviewerSnapshot reviewer,
            DateTimeOffset createdUtc,
            string runId = "",
            string addInBuild = "",
            string note = "")
        {
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            if (!generation.Catalog.TryGetGroup(groupId, out ReviewDecisionGroup group))
                throw new KeyNotFoundException("The selected review group is not in this decision generation.");
            return ReviewDecisionRequest.ForGroup(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.GenerationId,
                group.GroupId,
                group.SubjectIds,
                status,
                reviewer,
                createdUtc,
                runId: runId,
                addInBuild: addInBuild,
                note: note);
        }

        private static List<string> FreezeKnownSubjects(
            ReviewDecisionGeneration generation,
            IEnumerable<string> subjectIds)
        {
            if (subjectIds == null) throw new ArgumentNullException(nameof(subjectIds));
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string subjectId in subjectIds)
            {
                if (string.IsNullOrWhiteSpace(subjectId) || !generation.Catalog.ContainsSubject(subjectId))
                    throw new KeyNotFoundException("The selected review subject is not in this decision generation.");
                if (!seen.Add(subjectId))
                    throw new ArgumentException("A workspace selection cannot repeat a subject.", nameof(subjectIds));
                result.Add(subjectId);
            }
            if (result.Count == 0)
                throw new ArgumentException("Select at least one review subject.", nameof(subjectIds));
            result.Sort(StringComparer.Ordinal);
            return result;
        }
    }
}
