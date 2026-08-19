using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>Durable decision-event payload versions understood by the replay engine.</summary>
    public static class ReviewDecisionEventContractVersions
    {
        public const int V1 = 1;
    }

    /// <summary>Stable diagnostic codes for verbose decision-ledger replay failures.</summary>
    public static class ReviewDecisionReplayCodes
    {
        public const string RecordInvalid = "CMP-DECISION-REPLAY-RECORD-INVALID";
        public const string UnsupportedContractVersion = "CMP-DECISION-REPLAY-VERSION-UNSUPPORTED";
        public const string SequenceMismatch = "CMP-DECISION-REPLAY-SEQUENCE-INVALID";
        public const string GenerationChainMismatch = "CMP-DECISION-REPLAY-CHAIN-INVALID";
        public const string EventRejected = "CMP-DECISION-REPLAY-EVENT-REJECTED";
        public const string EventPayloadMismatch = "CMP-DECISION-REPLAY-EVENT-PAYLOAD-MISMATCH";
        public const string FinalGenerationMismatch = "CMP-DECISION-REPLAY-FINAL-GENERATION-INVALID";
    }

    /// <summary>
    /// One immutable, versioned event plus its explicit position in the decision-generation hash chain.
    /// Sequence is zero based. Prior/New generation IDs are redundant by design so persistence corruption can be
    /// diagnosed at the exact record without trusting a materialized effective-state snapshot.
    /// </summary>
    public sealed class ReviewDecisionReplayRecord
    {
        public ReviewDecisionReplayRecord(
            int contractVersion,
            int sequence,
            string priorGenerationId,
            string newGenerationId,
            DecisionEvent decisionEvent)
        {
            if (contractVersion <= 0) throw new ArgumentOutOfRangeException(nameof(contractVersion));
            if (sequence < 0) throw new ArgumentOutOfRangeException(nameof(sequence));
            ContractVersion = contractVersion;
            Sequence = sequence;
            PriorGenerationId = Require(priorGenerationId, nameof(priorGenerationId));
            NewGenerationId = Require(newGenerationId, nameof(newGenerationId));
            DecisionEvent = decisionEvent ?? throw new ArgumentNullException(nameof(decisionEvent));
        }

        public int ContractVersion { get; }
        public int Sequence { get; }
        public string PriorGenerationId { get; }
        public string NewGenerationId { get; }
        public DecisionEvent DecisionEvent { get; }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable replay identity is required.", name);
            return value;
        }
    }

    /// <summary>A replay failure with a stable code and exact ledger location suitable for verbose diagnostics.</summary>
    public sealed class ReviewDecisionReplayException : InvalidOperationException
    {
        internal ReviewDecisionReplayException(
            string code,
            int sequence,
            string eventId,
            string detail,
            Exception innerException = null)
            : base(FormatMessage(code, sequence, eventId, detail), innerException)
        {
            Code = code;
            Sequence = sequence;
            EventId = eventId ?? "";
        }

        public string Code { get; }
        /// <summary>Zero-based record sequence; for final-envelope failures this is the record count.</summary>
        public int Sequence { get; }
        public string EventId { get; }

        private static string FormatMessage(string code, int sequence, string eventId, string detail)
        {
            if (string.IsNullOrWhiteSpace(code))
                throw new ArgumentException("A replay diagnostic code is required.", nameof(code));
            return code + ": decision replay failed at sequence " + sequence +
                   (string.IsNullOrEmpty(eventId) ? "" : ", event " + eventId) +
                   ". " + (detail ?? "");
        }
    }

    /// <summary>
    /// Optional deterministic work counters for scale regression and operational diagnostics. Counts describe
    /// semantic replay work, not elapsed time, and are reset at the start of each Replay or Capture call.
    /// </summary>
    public sealed class ReviewDecisionReplayMetrics
    {
        public int InitializedSubjectCount { get; internal set; }
        public int ProcessedEventCount { get; internal set; }
        public long AffectedSubjectUpdateCount { get; internal set; }
        public long IndexedMasterObservationLookupCount { get; internal set; }
        public int FullGenerationMaterializationCount { get; internal set; }

        internal void Reset()
        {
            InitializedSubjectCount = 0;
            ProcessedEventCount = 0;
            AffectedSubjectUpdateCount = 0;
            IndexedMasterObservationLookupCount = 0;
            FullGenerationMaterializationCount = 0;
        }
    }

    /// <summary>
    /// Reconstructs immutable decision generations exclusively by applying their ordered events to the exact
    /// catalog root. Persisted subject-state or active-decision snapshots are never accepted as authority.
    /// </summary>
    public static class ReviewDecisionReplayEngine
    {
        public static ReviewDecisionGeneration Replay(
            ReviewDecisionCatalog catalog,
            IEnumerable<ReviewDecisionReplayRecord> records,
            string expectedGenerationId,
            string expectedPriorGenerationId = "") =>
            Replay(
                catalog,
                records,
                expectedGenerationId,
                expectedPriorGenerationId,
                CancellationToken.None,
                null);

        public static ReviewDecisionGeneration Replay(
            ReviewDecisionCatalog catalog,
            IEnumerable<ReviewDecisionReplayRecord> records,
            string expectedGenerationId,
            string expectedPriorGenerationId,
            CancellationToken cancellationToken,
            ReviewDecisionReplayMetrics metrics = null)
        {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (records == null) throw new ArgumentNullException(nameof(records));
            Require(expectedGenerationId, nameof(expectedGenerationId));
            expectedPriorGenerationId = expectedPriorGenerationId ?? "";
            ValidateOptionalIdentity(expectedPriorGenerationId, nameof(expectedPriorGenerationId));
            metrics?.Reset();
            cancellationToken.ThrowIfCancellationRequested();

            var accumulator = new ReplayAccumulator(
                catalog,
                cancellationToken,
                metrics);
            int expectedSequence = 0;
            foreach (ReviewDecisionReplayRecord record in records)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (record == null)
                    throw Failure(
                        ReviewDecisionReplayCodes.RecordInvalid,
                        expectedSequence,
                        "",
                        "The ledger contains a null replay record.");
                DecisionEvent stored = record.DecisionEvent;
                if (record.ContractVersion != ReviewDecisionEventContractVersions.V1)
                    throw Failure(
                        ReviewDecisionReplayCodes.UnsupportedContractVersion,
                        record.Sequence,
                        stored.EventId,
                        "Unsupported decision-event contract version " + record.ContractVersion + ".");
                if (record.Sequence != expectedSequence)
                    throw Failure(
                        ReviewDecisionReplayCodes.SequenceMismatch,
                        record.Sequence,
                        stored.EventId,
                        "Expected zero-based sequence " + expectedSequence + ".");
                if (!string.Equals(
                        record.PriorGenerationId,
                        accumulator.CurrentGenerationId,
                        StringComparison.Ordinal))
                    throw Failure(
                        ReviewDecisionReplayCodes.GenerationChainMismatch,
                        record.Sequence,
                        stored.EventId,
                        "The record prior generation does not match the replayed generation.");

                AppliedReplayTransition applied = accumulator.Apply(stored, record.Sequence);
                if (!string.Equals(
                        record.NewGenerationId,
                        applied.NewGenerationId,
                        StringComparison.Ordinal))
                    throw Failure(
                        ReviewDecisionReplayCodes.GenerationChainMismatch,
                        record.Sequence,
                        stored.EventId,
                        "The record new generation does not match the replayed generation.");
                expectedSequence++;
            }

            RequireFinalGeneration(
                accumulator.CurrentGenerationId,
                accumulator.PriorGenerationId,
                accumulator.LastEventId,
                expectedGenerationId,
                expectedPriorGenerationId,
                expectedSequence);
            return accumulator.Freeze();
        }

        /// <summary>
        /// Produces canonical replay records for an existing immutable generation. The generation is independently
        /// rebuilt and compared first, so callers never need to duplicate the generation-ID hashing recipe.
        /// </summary>
        public static IReadOnlyList<ReviewDecisionReplayRecord> Capture(
            ReviewDecisionGeneration generation) =>
            Capture(generation, CancellationToken.None, null);

        public static IReadOnlyList<ReviewDecisionReplayRecord> Capture(
            ReviewDecisionGeneration generation,
            CancellationToken cancellationToken,
            ReviewDecisionReplayMetrics metrics = null)
        {
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            metrics?.Reset();
            cancellationToken.ThrowIfCancellationRequested();
            var accumulator = new ReplayAccumulator(
                generation.Catalog,
                cancellationToken,
                metrics);
            var records = new List<ReviewDecisionReplayRecord>(generation.Events.Count);
            for (int sequence = 0; sequence < generation.Events.Count; sequence++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DecisionEvent stored = generation.Events[sequence];
                string priorGenerationId = accumulator.CurrentGenerationId;
                AppliedReplayTransition applied = accumulator.Apply(stored, sequence);
                records.Add(new ReviewDecisionReplayRecord(
                    ReviewDecisionEventContractVersions.V1,
                    sequence,
                    priorGenerationId,
                    applied.NewGenerationId,
                    applied.CanonicalEvent));
            }

            RequireFinalGeneration(
                accumulator.CurrentGenerationId,
                accumulator.PriorGenerationId,
                accumulator.LastEventId,
                generation.GenerationId,
                generation.PriorGenerationId,
                generation.Events.Count);
            ReviewDecisionGeneration rebuilt = accumulator.Freeze();
            RequireExactGenerationState(
                generation,
                rebuilt,
                cancellationToken);
            return new ReadOnlyCollection<ReviewDecisionReplayRecord>(records);
        }

        private static ReviewDecisionTransition ComputeStoredTransition(
            ReviewDecisionCatalog catalog,
            string currentGenerationId,
            Func<string, ReviewDecisionState> getSubjectState,
            DecisionEvent stored,
            int sequence,
            CancellationToken cancellationToken)
        {
            try
            {
                ReviewDecisionRequest request = ReconstructRequest(
                    catalog,
                    currentGenerationId,
                    stored);
                ReviewDecisionTransition transition = ReviewDecisionEngine.ComputeTransition(
                    catalog,
                    currentGenerationId,
                    getSubjectState,
                    request,
                    cancellationToken);
                RequireExactEvent(
                    transition.DecisionEvent,
                    stored,
                    sequence,
                    cancellationToken);
                return transition;
            }
            catch (ReviewDecisionReplayException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is InvalidOperationException ||
                exception is KeyNotFoundException)
            {
                throw Failure(
                    ReviewDecisionReplayCodes.EventRejected,
                    sequence,
                    stored?.EventId ?? "",
                    "The stored event cannot be applied to the exact catalog and prior state: " +
                    exception.Message,
                    exception);
            }
        }

        private static ReviewDecisionRequest ReconstructRequest(
            ReviewDecisionCatalog catalog,
            string currentGenerationId,
            DecisionEvent stored)
        {
            if (stored == null) throw new ArgumentNullException(nameof(stored));
            if (!string.Equals(stored.ProjectId, catalog.ProjectId, StringComparison.Ordinal))
                throw new InvalidOperationException("The event belongs to another project.");
            if (!string.Equals(
                    stored.EvidenceGenerationId,
                    catalog.BaseGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("The event belongs to another evidence generation.");
            if (!string.Equals(
                    stored.MappingGenerationId,
                    catalog.MappingGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("The event belongs to another mapping generation.");
            if (string.IsNullOrWhiteSpace(stored.RequestId))
                throw new InvalidOperationException("A durable v1 event has no request ID.");

            ReviewDecisionRequest request;
            switch (stored.SelectionScope)
            {
                case DecisionSelectionScope.SingleSubject:
                    request = ReviewDecisionRequest.ForSubject(
                        stored.ProjectId,
                        stored.EvidenceGenerationId,
                        stored.MappingGenerationId,
                        currentGenerationId,
                        stored.SubjectId,
                        stored.NewDecision,
                        stored.ReviewerSnapshot,
                        stored.CreatedUtc,
                        stored.SelectedObservationId,
                        stored.RunId,
                        stored.AddInBuild,
                        stored.Note);
                    break;
                case DecisionSelectionScope.MultipleSubjects:
                    request = ReviewDecisionRequest.ForSubjects(
                        stored.ProjectId,
                        stored.EvidenceGenerationId,
                        stored.MappingGenerationId,
                        currentGenerationId,
                        stored.AffectedSubjectIds,
                        stored.NewDecision,
                        stored.ReviewerSnapshot,
                        stored.CreatedUtc,
                        stored.SelectedObservationId,
                        stored.RunId,
                        stored.AddInBuild,
                        stored.Note);
                    break;
                case DecisionSelectionScope.Group:
                    request = ReviewDecisionRequest.ForGroup(
                        stored.ProjectId,
                        stored.EvidenceGenerationId,
                        stored.MappingGenerationId,
                        currentGenerationId,
                        stored.GroupId,
                        stored.AffectedSubjectIds,
                        stored.NewDecision,
                        stored.ReviewerSnapshot,
                        stored.CreatedUtc,
                        stored.SelectedObservationId,
                        stored.RunId,
                        stored.AddInBuild,
                        stored.Note);
                    break;
                default:
                    throw new InvalidOperationException("The event selection scope is unsupported.");
            }

            if (!string.Equals(request.RequestId, stored.RequestId, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The stored request ID does not match the reconstructed immutable request.");
            return request;
        }

        private static void RequireExactGenerationState(
            ReviewDecisionGeneration expected,
            ReviewDecisionGeneration actual,
            CancellationToken cancellationToken)
        {
            if (!string.Equals(expected.CatalogId, actual.CatalogId, StringComparison.Ordinal) ||
                expected.SubjectStates.Count != actual.SubjectStates.Count ||
                expected.Events.Count != actual.Events.Count)
                throw Failure(
                    ReviewDecisionReplayCodes.FinalGenerationMismatch,
                    actual.Events.Count,
                    actual.Events.Count == 0 ? "" : actual.Events[actual.Events.Count - 1].EventId,
                    "The captured generation does not match the canonically rebuilt state.");
            foreach (string subjectId in expected.Catalog.SubjectIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ReviewDecisionState left = expected.GetSubjectState(subjectId);
                ReviewDecisionState right = actual.GetSubjectState(subjectId);
                if (left.Status != right.Status ||
                    !string.Equals(
                        left.SelectedObservationId,
                        right.SelectedObservationId,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        left.EffectiveEventId,
                        right.EffectiveEventId,
                        StringComparison.Ordinal))
                    throw Failure(
                        ReviewDecisionReplayCodes.FinalGenerationMismatch,
                        actual.Events.Count,
                        right.EffectiveEventId,
                        "The captured generation contains an effective subject state not derived from its events.");
            }
        }

        private sealed class AppliedReplayTransition
        {
            internal AppliedReplayTransition(
                string newGenerationId,
                DecisionEvent canonicalEvent)
            {
                NewGenerationId = newGenerationId;
                CanonicalEvent = canonicalEvent;
            }

            internal string NewGenerationId { get; }
            internal DecisionEvent CanonicalEvent { get; }
        }

        /// <summary>
        /// Mutable only within one replay call. It holds the current delta-applied state and freezes a single
        /// immutable ReviewDecisionGeneration after all records have been validated.
        /// </summary>
        private sealed class ReplayAccumulator
        {
            private readonly ReviewDecisionCatalog _catalog;
            private readonly Dictionary<string, ReviewDecisionState> _states;
            private readonly List<DecisionEvent> _events;
            private readonly CancellationToken _cancellationToken;
            private readonly ReviewDecisionReplayMetrics _metrics;
            private bool _frozen;

            internal ReplayAccumulator(
                ReviewDecisionCatalog catalog,
                CancellationToken cancellationToken,
                ReviewDecisionReplayMetrics metrics)
            {
                _catalog = catalog;
                _cancellationToken = cancellationToken;
                _metrics = metrics;
                _states = new Dictionary<string, ReviewDecisionState>(
                    catalog.SubjectIds.Count,
                    StringComparer.Ordinal);
                _events = new List<DecisionEvent>();
                foreach (string subjectId in catalog.SubjectIds)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    _states.Add(
                        subjectId,
                        new ReviewDecisionState(
                            subjectId,
                            ReviewDecisionStatus.Unreviewed,
                            "",
                            ""));
                    if (_metrics != null) _metrics.InitializedSubjectCount++;
                }
                CurrentGenerationId = ReviewDecisionEngine.ComputeRootGenerationId(catalog);
                PriorGenerationId = "";
            }

            internal string CurrentGenerationId { get; private set; }
            internal string PriorGenerationId { get; private set; }
            internal string LastEventId => _events.Count == 0 ? "" : _events[_events.Count - 1].EventId;

            internal AppliedReplayTransition Apply(DecisionEvent stored, int sequence)
            {
                if (_frozen)
                    throw new InvalidOperationException("A frozen replay accumulator cannot accept events.");
                _cancellationToken.ThrowIfCancellationRequested();
                ReviewDecisionTransition transition = ComputeStoredTransition(
                    _catalog,
                    CurrentGenerationId,
                    GetSubjectState,
                    stored,
                    sequence,
                    _cancellationToken);
                string priorGenerationId = CurrentGenerationId;
                foreach (string subjectId in transition.AffectedSubjectIds)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    _states[subjectId] = new ReviewDecisionState(
                        subjectId,
                        transition.DecisionEvent.NewDecision,
                        transition.DecisionEvent.SelectedObservationId,
                        transition.DecisionEvent.EventId);
                    if (_metrics != null) _metrics.AffectedSubjectUpdateCount++;
                }
                _events.Add(transition.DecisionEvent);
                PriorGenerationId = priorGenerationId;
                CurrentGenerationId = transition.NewGenerationId;
                if (_metrics != null)
                {
                    _metrics.ProcessedEventCount++;
                    _metrics.IndexedMasterObservationLookupCount +=
                        transition.IndexedMasterObservationLookupCount;
                }
                return new AppliedReplayTransition(
                    transition.NewGenerationId,
                    transition.DecisionEvent);
            }

            internal ReviewDecisionGeneration Freeze()
            {
                if (_frozen)
                    throw new InvalidOperationException("A replay accumulator can be frozen only once.");
                _cancellationToken.ThrowIfCancellationRequested();
                var orderedStates = new List<ReviewDecisionState>(_states.Count);
                foreach (string subjectId in _catalog.SubjectIds)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    orderedStates.Add(_states[subjectId]);
                }
                _frozen = true;
                if (_metrics != null) _metrics.FullGenerationMaterializationCount++;
                ReviewDecisionGeneration result = new ReviewDecisionGeneration(
                    _catalog,
                    CurrentGenerationId,
                    PriorGenerationId,
                    orderedStates,
                    _events);
                _cancellationToken.ThrowIfCancellationRequested();
                return result;
            }

            private ReviewDecisionState GetSubjectState(string subjectId)
            {
                if (!_states.TryGetValue(subjectId ?? "", out ReviewDecisionState state))
                    throw new KeyNotFoundException("The subject is not in this replay accumulator.");
                return state;
            }
        }

        private static void RequireExactEvent(
            DecisionEvent expected,
            DecisionEvent actual,
            int sequence,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RequireEqual(expected.EventId, actual.EventId, "EventId", sequence, actual.EventId);
            RequireEqual(expected.RequestId, actual.RequestId, "RequestId", sequence, actual.EventId);
            RequireEqual(expected.ProjectId, actual.ProjectId, "ProjectId", sequence, actual.EventId);
            RequireEqual(expected.RunId, actual.RunId, "RunId", sequence, actual.EventId);
            RequireEqual(
                expected.EvidenceGenerationId,
                actual.EvidenceGenerationId,
                "EvidenceGenerationId",
                sequence,
                actual.EventId);
            RequireEqual(
                expected.MappingGenerationId,
                actual.MappingGenerationId,
                "MappingGenerationId",
                sequence,
                actual.EventId);
            RequireEqual(expected.SubjectId, actual.SubjectId, "SubjectId", sequence, actual.EventId);
            RequireEqual(expected.GroupId, actual.GroupId, "GroupId", sequence, actual.EventId);
            RequireSequenceEqual(
                expected.AffectedSubjectIds,
                actual.AffectedSubjectIds,
                "AffectedSubjectIds",
                sequence,
                actual.EventId,
                cancellationToken);
            RequireEqual(expected.PriorDecision, actual.PriorDecision, "PriorDecision", sequence, actual.EventId);
            RequireEqual(expected.NewDecision, actual.NewDecision, "NewDecision", sequence, actual.EventId);
            RequireEqual(
                expected.SelectedObservationId,
                actual.SelectedObservationId,
                "SelectedObservationId",
                sequence,
                actual.EventId);
            RequireSequenceEqual(
                expected.SupportingRevisionIds,
                actual.SupportingRevisionIds,
                "SupportingRevisionIds",
                sequence,
                actual.EventId,
                cancellationToken);
            RequireSequenceEqual(
                expected.SupportingContentHashes,
                actual.SupportingContentHashes,
                "SupportingContentHashes",
                sequence,
                actual.EventId,
                cancellationToken);
            RequireReviewerEqual(
                expected.ReviewerSnapshot,
                actual.ReviewerSnapshot,
                sequence,
                actual.EventId,
                cancellationToken);
            if (expected.CreatedUtc.Ticks != actual.CreatedUtc.Ticks ||
                expected.CreatedUtc.Offset != actual.CreatedUtc.Offset)
                PayloadMismatch("CreatedUtc", sequence, actual.EventId);
            RequireEqual(expected.AddInBuild, actual.AddInBuild, "AddInBuild", sequence, actual.EventId);
            RequireEqual(expected.Note, actual.Note, "Note", sequence, actual.EventId);
            RequireEqual(
                expected.SelectionScope,
                actual.SelectionScope,
                "SelectionScope",
                sequence,
                actual.EventId);
            RequireEqual(
                expected.SupersededEventId,
                actual.SupersededEventId,
                "SupersededEventId",
                sequence,
                actual.EventId);
            RequireLocationsEqual(
                expected.Locations,
                actual.Locations,
                sequence,
                actual.EventId,
                cancellationToken);
        }

        private static void RequireReviewerEqual(
            ReviewerSnapshot expected,
            ReviewerSnapshot actual,
            int sequence,
            string eventId,
            CancellationToken cancellationToken)
        {
            if (expected == null || actual == null)
            {
                if (!ReferenceEquals(expected, actual)) PayloadMismatch("ReviewerSnapshot", sequence, eventId);
                return;
            }
            RequireEqual(
                expected.ConfirmedDisplayName,
                actual.ConfirmedDisplayName,
                "ReviewerSnapshot.ConfirmedDisplayName",
                sequence,
                eventId);
            RequireEqual(
                expected.OfficeDisplayName,
                actual.OfficeDisplayName,
                "ReviewerSnapshot.OfficeDisplayName",
                sequence,
                eventId);
            RequireSequenceEqual(
                expected.IdentitySourceLabels,
                actual.IdentitySourceLabels,
                "ReviewerSnapshot.IdentitySourceLabels",
                sequence,
                eventId,
                cancellationToken);
        }

        private static void RequireLocationsEqual(
            IReadOnlyList<DecisionLocationContext> expected,
            IReadOnlyList<DecisionLocationContext> actual,
            int sequence,
            string eventId,
            CancellationToken cancellationToken)
        {
            if (expected.Count != actual.Count) PayloadMismatch("Locations.Count", sequence, eventId);
            for (int index = 0; index < expected.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DecisionLocationContext left = expected[index];
                DecisionLocationContext right = actual[index];
                string prefix = "Locations[" + index + "].";
                RequireEqual(left.Role, right.Role, prefix + "Role", sequence, eventId);
                RequireEqual(left.RevisionId, right.RevisionId, prefix + "RevisionId", sequence, eventId);
                RequireEqual(
                    left.LogicalSheetId,
                    right.LogicalSheetId,
                    prefix + "LogicalSheetId",
                    sequence,
                    eventId);
                RequireEqual(left.SheetName, right.SheetName, prefix + "SheetName", sequence, eventId);
                RequireEqual(left.Address, right.Address, prefix + "Address", sequence, eventId);
                RequireEqual(left.SourceLabel, right.SourceLabel, prefix + "SourceLabel", sequence, eventId);
                if (left.PhysicalRow != right.PhysicalRow)
                    PayloadMismatch(prefix + "PhysicalRow", sequence, eventId);
                if (left.PhysicalColumn != right.PhysicalColumn)
                    PayloadMismatch(prefix + "PhysicalColumn", sequence, eventId);
            }
        }

        private static void RequireSequenceEqual(
            IReadOnlyList<string> expected,
            IReadOnlyList<string> actual,
            string field,
            int sequence,
            string eventId,
            CancellationToken cancellationToken)
        {
            if (expected.Count != actual.Count) PayloadMismatch(field + ".Count", sequence, eventId);
            for (int index = 0; index < expected.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.Equals(expected[index], actual[index], StringComparison.Ordinal))
                    PayloadMismatch(field + "[" + index + "]", sequence, eventId);
            }
        }

        private static void RequireEqual(
            string expected,
            string actual,
            string field,
            int sequence,
            string eventId)
        {
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
                PayloadMismatch(field, sequence, eventId);
        }

        private static void RequireEqual<T>(
            T expected,
            T actual,
            string field,
            int sequence,
            string eventId)
            where T : struct
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
                PayloadMismatch(field, sequence, eventId);
        }

        private static void PayloadMismatch(string field, int sequence, string eventId)
        {
            throw Failure(
                ReviewDecisionReplayCodes.EventPayloadMismatch,
                sequence,
                eventId,
                "The stored event differs from the canonically regenerated field " + field + ".");
        }

        private static void RequireFinalGeneration(
            string actualGenerationId,
            string actualPriorGenerationId,
            string lastEventId,
            string expectedGenerationId,
            string expectedPriorGenerationId,
            int recordCount)
        {
            if (!string.Equals(actualGenerationId, expectedGenerationId, StringComparison.Ordinal) ||
                !string.Equals(
                    actualPriorGenerationId,
                    expectedPriorGenerationId,
                    StringComparison.Ordinal))
                throw Failure(
                    ReviewDecisionReplayCodes.FinalGenerationMismatch,
                    recordCount,
                    lastEventId,
                    "The replayed final or prior decision-generation ID disagrees with the envelope.");
        }

        private static ReviewDecisionReplayException Failure(
            string code,
            int sequence,
            string eventId,
            string detail,
            Exception innerException = null) =>
            new ReviewDecisionReplayException(code, sequence, eventId, detail, innerException);

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable replay identity is required.", name);
            return value;
        }

        private static void ValidateOptionalIdentity(string value, string name)
        {
            if (value.Length != 0 && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An optional replay identity must be empty or non-blank.", name);
        }
    }
}
