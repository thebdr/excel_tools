using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    public static class ReviewDecisionLocationRoles
    {
        public const string Selected = "selected";
        public const string Master = "master";
    }

    /// <summary>The effective review state of one subject, including the audit event that established it.</summary>
    public sealed class ReviewDecisionState
    {
        internal ReviewDecisionState(
            string subjectId,
            ReviewDecisionStatus status,
            string selectedObservationId,
            string effectiveEventId)
        {
            SubjectId = subjectId;
            Status = status;
            SelectedObservationId = selectedObservationId ?? "";
            EffectiveEventId = effectiveEventId ?? "";
        }

        public string SubjectId { get; }
        public ReviewDecisionStatus Status { get; }
        public string SelectedObservationId { get; }
        /// <summary>Latest event establishing this state; empty only for untouched root state.</summary>
        public string EffectiveEventId { get; }
        public bool HasReviewedDecision => Status != ReviewDecisionStatus.Unreviewed;
    }

    /// <summary>
    /// Immutable decision-ledger generation. Events are the complete append-only history; SubjectStates contains
    /// every catalog subject, while ActiveDecisions contains only currently reviewed (non-Unreviewed) subjects.
    /// </summary>
    public sealed class ReviewDecisionGeneration
    {
        private readonly IReadOnlyDictionary<string, DecisionEvent> _eventsById;

        internal ReviewDecisionGeneration(
            ReviewDecisionCatalog catalog,
            string generationId,
            string priorGenerationId,
            IEnumerable<ReviewDecisionState> subjectStates,
            IEnumerable<DecisionEvent> events)
        {
            Catalog = catalog ?? throw new ArgumentNullException(nameof(catalog));
            GenerationId = Require(generationId, nameof(generationId));
            PriorGenerationId = priorGenerationId ?? "";

            var frozenEvents = new List<DecisionEvent>(events ?? Array.Empty<DecisionEvent>());
            if (frozenEvents.Any(item => item == null))
                throw new ArgumentException("A decision generation cannot contain null events.", nameof(events));
            if (frozenEvents.Select(item => item.EventId).Distinct(StringComparer.Ordinal).Count() != frozenEvents.Count)
                throw new ArgumentException("A decision generation cannot repeat an event ID.", nameof(events));
            Events = new ReadOnlyCollection<DecisionEvent>(frozenEvents);
            _eventsById = new ReadOnlyDictionary<string, DecisionEvent>(
                frozenEvents.ToDictionary(item => item.EventId, StringComparer.Ordinal));

            var stateMap = new Dictionary<string, ReviewDecisionState>(StringComparer.Ordinal);
            foreach (ReviewDecisionState state in subjectStates ?? Array.Empty<ReviewDecisionState>())
            {
                if (state == null)
                    throw new ArgumentException("A decision generation cannot contain a null subject state.", nameof(subjectStates));
                if (!Catalog.ContainsSubject(state.SubjectId))
                    throw new ArgumentException("A decision state must belong to the exact catalog.", nameof(subjectStates));
                if (stateMap.ContainsKey(state.SubjectId))
                    throw new ArgumentException("A decision generation cannot repeat a subject state.", nameof(subjectStates));
                stateMap.Add(state.SubjectId, state);
                if (state.EffectiveEventId.Length != 0 && !_eventsById.ContainsKey(state.EffectiveEventId))
                    throw new ArgumentException("A decision state references an event outside this generation.", nameof(subjectStates));
            }
            if (stateMap.Count != Catalog.SubjectIds.Count ||
                Catalog.SubjectIds.Any(subjectId => !stateMap.ContainsKey(subjectId)))
                throw new ArgumentException("A decision generation requires exactly one state per catalog subject.", nameof(subjectStates));
            SubjectStates = new ReadOnlyDictionary<string, ReviewDecisionState>(stateMap);

            var active = new Dictionary<string, DecisionEvent>(StringComparer.Ordinal);
            foreach (ReviewDecisionState state in stateMap.Values.Where(item => item.HasReviewedDecision))
            {
                if (state.EffectiveEventId.Length == 0 ||
                    !_eventsById.TryGetValue(state.EffectiveEventId, out DecisionEvent activeEvent))
                    throw new ArgumentException("A reviewed subject state requires an effective decision event.", nameof(subjectStates));
                active.Add(state.SubjectId, activeEvent);
            }
            ActiveDecisions = new ReadOnlyDictionary<string, DecisionEvent>(active);
        }

        public ReviewDecisionCatalog Catalog { get; }
        public string CatalogId => Catalog.CatalogId;
        public string ProjectId => Catalog.ProjectId;
        public string BaseGenerationId => Catalog.BaseGenerationId;
        public string MappingGenerationId => Catalog.MappingGenerationId;
        public string GenerationId { get; }
        public string PriorGenerationId { get; }
        public IReadOnlyDictionary<string, ReviewDecisionState> SubjectStates { get; }
        public IReadOnlyDictionary<string, DecisionEvent> ActiveDecisions { get; }
        public IReadOnlyList<DecisionEvent> Events { get; }

        public ReviewDecisionState GetSubjectState(string subjectId)
        {
            if (!SubjectStates.TryGetValue(subjectId ?? "", out ReviewDecisionState state))
                throw new KeyNotFoundException("The subject is not in this decision generation.");
            return state;
        }

        public bool TryGetActiveDecision(string subjectId, out DecisionEvent decisionEvent) =>
            ActiveDecisions.TryGetValue(subjectId ?? "", out decisionEvent);

        /// <summary>
        /// Returns the latest event establishing a subject state, including an Unreviewed clear event. Untouched
        /// root subjects return false.
        /// </summary>
        public bool TryGetEffectiveEvent(string subjectId, out DecisionEvent decisionEvent)
        {
            decisionEvent = null;
            if (!SubjectStates.TryGetValue(subjectId ?? "", out ReviewDecisionState state) ||
                state.EffectiveEventId.Length == 0)
                return false;
            return _eventsById.TryGetValue(state.EffectiveEventId, out decisionEvent);
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable identity is required.", name);
            return value;
        }
    }

    /// <summary>
    /// Canonical result of validating one request against one exact generation state. It contains only the delta
    /// needed by either immutable Apply or the linear replay accumulator.
    /// </summary>
    internal sealed class ReviewDecisionTransition
    {
        internal ReviewDecisionTransition(
            IReadOnlyList<string> affectedSubjectIds,
            DecisionEvent decisionEvent,
            string newGenerationId,
            int indexedMasterObservationLookupCount)
        {
            AffectedSubjectIds = affectedSubjectIds;
            DecisionEvent = decisionEvent;
            NewGenerationId = newGenerationId;
            IndexedMasterObservationLookupCount = indexedMasterObservationLookupCount;
        }

        internal IReadOnlyList<string> AffectedSubjectIds { get; }
        internal DecisionEvent DecisionEvent { get; }
        internal string NewGenerationId { get; }
        internal int IndexedMasterObservationLookupCount { get; }
    }

    /// <summary>Creates and advances immutable append-only reviewer-decision generations.</summary>
    public static class ReviewDecisionEngine
    {
        public static ReviewDecisionGeneration Create(ReviewDecisionCatalog catalog)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            string generationId = ComputeRootGenerationId(catalog);
            return new ReviewDecisionGeneration(
                catalog,
                generationId,
                "",
                catalog.SubjectIds.Select(subjectId => new ReviewDecisionState(
                    subjectId,
                    ReviewDecisionStatus.Unreviewed,
                    "",
                    "")),
                Array.Empty<DecisionEvent>());
        }

        public static ReviewDecisionGeneration Apply(
            ReviewDecisionGeneration current,
            ReviewDecisionRequest request)
        {
            if (current == null) throw new ArgumentNullException(nameof(current));
            if (request == null) throw new ArgumentNullException(nameof(request));
            ReviewDecisionTransition transition = ComputeTransition(
                current.Catalog,
                current.GenerationId,
                current.GetSubjectState,
                request,
                CancellationToken.None);
            DecisionEvent decisionEvent = transition.DecisionEvent;
            var affected = new HashSet<string>(
                transition.AffectedSubjectIds,
                StringComparer.Ordinal);
            var states = current.SubjectStates.Values
                .Select(state => affected.Contains(state.SubjectId)
                    ? new ReviewDecisionState(
                        state.SubjectId,
                        request.NewDecision,
                        decisionEvent.SelectedObservationId,
                        decisionEvent.EventId)
                    : state)
                .OrderBy(state => state.SubjectId, StringComparer.Ordinal)
                .ToList();
            var events = new List<DecisionEvent>(current.Events) { decisionEvent };
            return new ReviewDecisionGeneration(
                current.Catalog,
                transition.NewGenerationId,
                current.GenerationId,
                states,
                events);
        }

        internal static string ComputeRootGenerationId(ReviewDecisionCatalog catalog)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            return StableId.Compute(
                "review-decision-generation-root-v1",
                catalog.CatalogId);
        }

        /// <summary>
        /// Validates and computes one canonical transition without copying the complete state or event history.
        /// </summary>
        internal static ReviewDecisionTransition ComputeTransition(
            ReviewDecisionCatalog catalog,
            string currentGenerationId,
            Func<string, ReviewDecisionState> getSubjectState,
            ReviewDecisionRequest request,
            CancellationToken cancellationToken)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (string.IsNullOrWhiteSpace(currentGenerationId))
                throw new ArgumentException(
                    "A current decision generation ID is required.",
                    nameof(currentGenerationId));
            if (getSubjectState == null) throw new ArgumentNullException(nameof(getSubjectState));
            if (request == null) throw new ArgumentNullException(nameof(request));
            cancellationToken.ThrowIfCancellationRequested();
            ValidateRequestContext(catalog, currentGenerationId, request);

            IReadOnlyList<string> affectedSubjectIds = ResolveAffectedSubjects(
                catalog,
                request,
                cancellationToken);
            Observation selectedObservation = ValidateObservationSelection(
                catalog,
                request,
                affectedSubjectIds);
            var priorStates = new List<ReviewDecisionState>(affectedSubjectIds.Count);
            foreach (string subjectId in affectedSubjectIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                priorStates.Add(getSubjectState(subjectId));
            }
            EnsureBatchHasOnePriorState(priorStates);

            ReviewDecisionState prior = priorStates[0];
            string nextSelectedObservationId = selectedObservation?.ObservationId ?? "";
            if (priorStates.All(state =>
                    state.Status == request.NewDecision &&
                    string.Equals(
                        state.SelectedObservationId,
                        nextSelectedObservationId,
                        StringComparison.Ordinal)))
                throw new InvalidOperationException("The requested decision already is the effective state.");

            IReadOnlyList<string> supportingRevisionIds = selectedObservation == null
                ? Array.Empty<string>()
                : new[] { selectedObservation.RevisionId };
            IReadOnlyList<string> supportingContentHashes = selectedObservation == null
                ? Array.Empty<string>()
                : new[] { selectedObservation.ContentHash };
            IReadOnlyList<DecisionLocationContext> locations = BuildLocations(
                catalog,
                selectedObservation,
                out int indexedMasterObservationLookupCount);

            string eventId = ComputeEventId(
                catalog.CatalogId,
                currentGenerationId,
                request,
                affectedSubjectIds,
                prior,
                supportingRevisionIds,
                supportingContentHashes,
                locations);
            var decisionEvent = new DecisionEvent(
                eventId,
                catalog.ProjectId,
                catalog.BaseGenerationId,
                catalog.MappingGenerationId,
                prior.Status,
                request.NewDecision,
                request.CreatedUtc,
                request.SelectionScope,
                request.RunId,
                request.SubjectId,
                request.GroupId,
                affectedSubjectIds,
                nextSelectedObservationId,
                supportingRevisionIds,
                supportingContentHashes,
                request.Reviewer,
                request.AddInBuild,
                request.Note,
                prior.EffectiveEventId,
                locations,
                request.RequestId);
            string newGenerationId = StableId.Compute(
                "review-decision-generation-v1",
                currentGenerationId,
                decisionEvent.EventId);
            return new ReviewDecisionTransition(
                affectedSubjectIds,
                decisionEvent,
                newGenerationId,
                indexedMasterObservationLookupCount);
        }

        private static void ValidateRequestContext(
            ReviewDecisionCatalog catalog,
            string currentGenerationId,
            ReviewDecisionRequest request)
        {
            if (!string.Equals(catalog.ProjectId, request.ProjectId, StringComparison.Ordinal))
                throw new InvalidOperationException("The decision request belongs to another project.");
            if (!string.Equals(catalog.BaseGenerationId, request.BaseGenerationId, StringComparison.Ordinal))
                throw new InvalidOperationException("The decision request targets a stale evidence generation.");
            if (!string.Equals(catalog.MappingGenerationId, request.MappingGenerationId, StringComparison.Ordinal))
                throw new InvalidOperationException("The decision request targets a stale mapping generation.");
            if (!string.Equals(
                    currentGenerationId,
                    request.ExpectedDecisionGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("The decision request targets a stale decision generation.");
        }

        private static IReadOnlyList<string> ResolveAffectedSubjects(
            ReviewDecisionCatalog catalog,
            ReviewDecisionRequest request,
            CancellationToken cancellationToken)
        {
            foreach (string subjectId in request.AffectedSubjectIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!catalog.ContainsSubject(subjectId))
                    throw new InvalidOperationException(
                        "The decision request references an unknown subject: " + subjectId);
            }

            if (request.SelectionScope == DecisionSelectionScope.Group)
            {
                if (!catalog.TryGetGroup(request.GroupId, out ReviewDecisionGroup group))
                    throw new InvalidOperationException(
                        "The decision request references an unknown group: " + request.GroupId);
                if (!group.SubjectIds.SequenceEqual(request.AffectedSubjectIds, StringComparer.Ordinal))
                    throw new InvalidOperationException(
                        "The frozen affected-subject set does not exactly match the review group.");
            }

            // ReviewDecisionRequest already owns an immutable, ordinally sorted copy. DecisionEvent will copy it
            // again so the audit record remains independent of both request and catalog lifetime.
            return request.AffectedSubjectIds;
        }

        private static Observation ValidateObservationSelection(
            ReviewDecisionCatalog catalog,
            ReviewDecisionRequest request,
            IReadOnlyList<string> affectedSubjectIds)
        {
            bool selectsObservation = request.NewDecision == ReviewDecisionStatus.UseMaster ||
                                      request.NewDecision == ReviewDecisionStatus.UseCandidate;
            if (!selectsObservation)
                return null;
            if (affectedSubjectIds.Count != 1)
                throw new InvalidOperationException(
                    "A selected observation can establish only one subject decision.");
            if (!catalog.TryGetSubject(affectedSubjectIds[0], out ReviewDecisionSubject subject))
                throw new InvalidOperationException(
                    "The selected subject is not in the exact review catalog.");
            if (!subject.CanSelectObservation)
                throw new InvalidOperationException(
                    subject.HasExplicitMetadata
                        ? "This review subject is nonselectable under its explicit decision policy."
                        : "Legacy subject metadata does not authorize observation selection; rebuild the catalog with an explicit formula/value policy.");
            if (!catalog.TryGetObservation(request.SelectedObservationId, out Observation observation))
                throw new InvalidOperationException(
                    "The selected observation is not in the exact review catalog.");
            if (!string.Equals(observation.SubjectId, affectedSubjectIds[0], StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The selected observation belongs to another logical subject.");

            bool isMaster = string.Equals(
                observation.RevisionId,
                catalog.MasterRevisionId,
                StringComparison.Ordinal);
            if (request.NewDecision == ReviewDecisionStatus.UseMaster && !isMaster)
                throw new InvalidOperationException(
                    "A use-master decision must select this catalog's master-revision observation.");
            if (request.NewDecision == ReviewDecisionStatus.UseCandidate && isMaster)
                throw new InvalidOperationException(
                    "A use-candidate decision cannot select the master-revision observation.");
            return observation;
        }

        private static IReadOnlyList<DecisionLocationContext> BuildLocations(
            ReviewDecisionCatalog catalog,
            Observation selectedObservation,
            out int indexedMasterObservationLookupCount)
        {
            indexedMasterObservationLookupCount = 0;
            if (selectedObservation == null)
                return Array.Empty<DecisionLocationContext>();

            var result = new List<DecisionLocationContext>();
            AddLocation(result, ReviewDecisionLocationRoles.Selected, selectedObservation);
            indexedMasterObservationLookupCount = 1;
            if (catalog.TryGetObservation(
                    selectedObservation.SubjectId,
                    catalog.MasterRevisionId,
                    out Observation master))
                AddLocation(result, ReviewDecisionLocationRoles.Master, master);
            return new ReadOnlyCollection<DecisionLocationContext>(result);
        }

        private static void AddLocation(
            ICollection<DecisionLocationContext> target,
            string role,
            Observation observation)
        {
            ObservationProvenance provenance = observation.Provenance;
            if (provenance == null ||
                (provenance.LogicalSheetId.Length == 0 &&
                 provenance.SourceLabel.Length == 0 &&
                 provenance.SheetName.Length == 0 &&
                 provenance.Address.Length == 0 &&
                 provenance.PhysicalRow == 0 &&
                 provenance.PhysicalColumn == 0))
                return;
            target.Add(new DecisionLocationContext(
                role,
                observation.RevisionId,
                logicalSheetId: provenance.LogicalSheetId,
                sheetName: provenance.SheetName,
                address: provenance.Address,
                sourceLabel: provenance.SourceLabel,
                physicalRow: provenance.PhysicalRow,
                physicalColumn: provenance.PhysicalColumn));
        }

        private static void EnsureBatchHasOnePriorState(IReadOnlyList<ReviewDecisionState> priorStates)
        {
            if (priorStates.Count <= 1) return;
            ReviewDecisionState first = priorStates[0];
            bool ambiguous = priorStates.Skip(1).Any(state =>
                state.Status != first.Status ||
                !string.Equals(
                    state.SelectedObservationId,
                    first.SelectedObservationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    state.EffectiveEventId,
                    first.EffectiveEventId,
                    StringComparison.Ordinal));
            if (ambiguous)
                throw new InvalidOperationException(
                    "This batch has mixed prior decision provenance and cannot name one superseded event. " +
                    "Review homogeneous subsets or individual subjects instead.");
        }

        private static string ComputeEventId(
            string catalogId,
            string currentGenerationId,
            ReviewDecisionRequest request,
            IReadOnlyList<string> affectedSubjectIds,
            ReviewDecisionState prior,
            IEnumerable<string> supportingRevisionIds,
            IEnumerable<string> supportingContentHashes,
            IEnumerable<DecisionLocationContext> locations)
        {
            var components = new List<string>
            {
                catalogId,
                currentGenerationId,
                request.RequestId,
                ((int)prior.Status).ToString(CultureInfo.InvariantCulture),
                prior.SelectedObservationId,
                prior.EffectiveEventId,
                ((int)request.NewDecision).ToString(CultureInfo.InvariantCulture),
                request.SelectedObservationId,
                ((int)request.SelectionScope).ToString(CultureInfo.InvariantCulture),
                request.SubjectId,
                request.GroupId,
            };
            components.AddRange(affectedSubjectIds.Select(item => "affected:" + item));
            components.AddRange(supportingRevisionIds.Select(item => "supporting-revision:" + item));
            components.AddRange(supportingContentHashes.Select(item => "supporting-content:" + item));
            foreach (DecisionLocationContext location in locations)
            {
                components.Add("location-role:" + location.Role);
                components.Add("location-revision:" + location.RevisionId);
                components.Add("location-logical-sheet:" + location.LogicalSheetId);
                components.Add("location-sheet:" + location.SheetName);
                components.Add("location-address:" + location.Address);
                components.Add("location-source:" + location.SourceLabel);
                components.Add("location-row:" +
                               location.PhysicalRow.ToString(CultureInfo.InvariantCulture));
                components.Add("location-column:" +
                               location.PhysicalColumn.ToString(CultureInfo.InvariantCulture));
            }
            return StableId.Compute("review-decision-event-v1", components.ToArray());
        }
    }
}
