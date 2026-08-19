using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewDecisionReplayEngineTests
    {
        private static readonly DateTimeOffset When = new DateTimeOffset(
            2026, 7, 17, 10, 11, 12, TimeSpan.Zero);
        private static readonly ReviewerSnapshot Reviewer = new ReviewerSnapshot(
            "Confirmed Reviewer",
            "Office Reviewer",
            new[] { "confirmed-dialog", "office-profile" });

        [Fact]
        public void Empty_root_capture_and_replay_are_exact_and_derive_all_subject_states()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());

            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(root);
            ReviewDecisionGeneration replayed = ReviewDecisionReplayEngine.Replay(
                root.Catalog,
                records,
                root.GenerationId,
                root.PriorGenerationId);

            Assert.Empty(records);
            Assert.Equal(root.GenerationId, replayed.GenerationId);
            Assert.Empty(replayed.PriorGenerationId);
            Assert.Empty(replayed.Events);
            Assert.Empty(replayed.ActiveDecisions);
            Assert.Equal(root.Catalog.SubjectIds.Count, replayed.SubjectStates.Count);
            Assert.All(
                replayed.SubjectStates.Values,
                state =>
                {
                    Assert.Equal(ReviewDecisionStatus.Unreviewed, state.Status);
                    Assert.Empty(state.SelectedObservationId);
                    Assert.Empty(state.EffectiveEventId);
                });
        }

        [Fact]
        public void Capture_and_replay_rebuild_selection_group_and_clear_history_with_exact_chain_ids()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionGeneration candidate = ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(
                    root,
                    "subject-1",
                    ReviewDecisionStatus.UseCandidate,
                    "observation-1-candidate",
                    When,
                    "Candidate wins"));
            ReviewDecisionGeneration grouped = ReviewDecisionEngine.Apply(
                candidate,
                ReviewDecisionRequest.ForGroup(
                    candidate.ProjectId,
                    candidate.BaseGenerationId,
                    candidate.MappingGenerationId,
                    candidate.GenerationId,
                    "group-23",
                    new[] { "subject-3", "subject-2" },
                    ReviewDecisionStatus.Equivalent,
                    Reviewer,
                    When.AddMinutes(1),
                    runId: "run-2",
                    addInBuild: "2.0",
                    note: "Group agrees"));
            ReviewDecisionGeneration cleared = ReviewDecisionEngine.Apply(
                grouped,
                ReviewDecisionRequest.ClearSubject(
                    grouped.ProjectId,
                    grouped.BaseGenerationId,
                    grouped.MappingGenerationId,
                    grouped.GenerationId,
                    "subject-1",
                    Reviewer,
                    When.AddMinutes(2),
                    runId: "run-3",
                    addInBuild: "2.0",
                    note: "Clear selection"));

            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(cleared);
            ReviewDecisionGeneration replayed = ReviewDecisionReplayEngine.Replay(
                cleared.Catalog,
                records,
                cleared.GenerationId,
                cleared.PriorGenerationId);

            Assert.Equal(3, records.Count);
            for (int index = 0; index < records.Count; index++)
            {
                Assert.Equal(ReviewDecisionEventContractVersions.V1, records[index].ContractVersion);
                Assert.Equal(index, records[index].Sequence);
                Assert.NotEmpty(records[index].DecisionEvent.RequestId);
                Assert.Equal(
                    records[index].DecisionEvent.GenerationId,
                    records[index].DecisionEvent.EvidenceGenerationId);
                Assert.Equal(
                    index == 0 ? root.GenerationId : records[index - 1].NewGenerationId,
                    records[index].PriorGenerationId);
                Assert.Equal(
                    index == records.Count - 1
                        ? cleared.GenerationId
                        : records[index + 1].PriorGenerationId,
                    records[index].NewGenerationId);
            }
            Assert.Equal(cleared.GenerationId, replayed.GenerationId);
            Assert.Equal(cleared.PriorGenerationId, replayed.PriorGenerationId);
            Assert.Equal(
                cleared.Events.Select(item => item.EventId),
                replayed.Events.Select(item => item.EventId));
            Assert.All(
                replayed.Events.Zip(cleared.Events),
                pair => Assert.NotSame(pair.Second, pair.First));
            Assert.Equal(ReviewDecisionStatus.Unreviewed, replayed.GetSubjectState("subject-1").Status);
            Assert.True(replayed.TryGetEffectiveEvent("subject-1", out DecisionEvent clearEvent));
            Assert.Equal(ReviewDecisionStatus.Unreviewed, clearEvent.NewDecision);
            Assert.Equal(ReviewDecisionStatus.Equivalent, replayed.GetSubjectState("subject-2").Status);
            Assert.Equal(ReviewDecisionStatus.Equivalent, replayed.GetSubjectState("subject-3").Status);
            Assert.Equal(2, replayed.ActiveDecisions.Count);

            Assert.Empty(root.Events);
            Assert.Single(candidate.Events);
            Assert.Equal(2, grouped.Events.Count);
            Assert.Equal(3, cleared.Events.Count);
            Assert.Equal(ReviewDecisionStatus.UseCandidate, candidate.GetSubjectState("subject-1").Status);
        }

        [Fact]
        public void Replay_rejects_unknown_version_reorder_omission_duplication_and_generation_chain_tampering()
        {
            ReviewDecisionGeneration history = ThreeEventHistory();
            List<ReviewDecisionReplayRecord> captured =
                ReviewDecisionReplayEngine.Capture(history).ToList();

            ReviewDecisionReplayException version = Assert.Throws<ReviewDecisionReplayException>(() =>
                Replay(
                    history,
                    Replace(
                        captured,
                        0,
                        Record(captured[0], contractVersion: 99))));
            AssertReplayFailure(
                version,
                ReviewDecisionReplayCodes.UnsupportedContractVersion,
                0,
                captured[0].DecisionEvent.EventId);

            ReviewDecisionReplayException reordered = Assert.Throws<ReviewDecisionReplayException>(() =>
                Replay(history, new[] { captured[1], captured[0], captured[2] }));
            AssertReplayFailure(
                reordered,
                ReviewDecisionReplayCodes.SequenceMismatch,
                1,
                captured[1].DecisionEvent.EventId);

            ReviewDecisionReplayException omitted = Assert.Throws<ReviewDecisionReplayException>(() =>
                Replay(history, new[] { captured[0], captured[2] }));
            AssertReplayFailure(
                omitted,
                ReviewDecisionReplayCodes.SequenceMismatch,
                2,
                captured[2].DecisionEvent.EventId);

            ReviewDecisionReplayException duplicated = Assert.Throws<ReviewDecisionReplayException>(() =>
                Replay(history, new[] { captured[0], captured[0], captured[1], captured[2] }));
            AssertReplayFailure(
                duplicated,
                ReviewDecisionReplayCodes.SequenceMismatch,
                0,
                captured[0].DecisionEvent.EventId);

            ReviewDecisionReplayException prior = Assert.Throws<ReviewDecisionReplayException>(() =>
                Replay(
                    history,
                    Replace(
                        captured,
                        1,
                        Record(captured[1], priorGenerationId: StableId.Compute("wrong-prior")))));
            AssertReplayFailure(
                prior,
                ReviewDecisionReplayCodes.GenerationChainMismatch,
                1,
                captured[1].DecisionEvent.EventId);

            ReviewDecisionReplayException next = Assert.Throws<ReviewDecisionReplayException>(() =>
                Replay(
                    history,
                    Replace(
                        captured,
                        1,
                        Record(captured[1], newGenerationId: StableId.Compute("wrong-next")))));
            AssertReplayFailure(
                next,
                ReviewDecisionReplayCodes.GenerationChainMismatch,
                1,
                captured[1].DecisionEvent.EventId);
        }

        [Fact]
        public void Replay_rejects_wrong_final_generation_envelope_with_stable_diagnostic_context()
        {
            ReviewDecisionGeneration history = ThreeEventHistory();
            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(history);

            ReviewDecisionReplayException generation = Assert.Throws<ReviewDecisionReplayException>(() =>
                ReviewDecisionReplayEngine.Replay(
                    history.Catalog,
                    records,
                    StableId.Compute("wrong-final"),
                    history.PriorGenerationId));
            AssertReplayFailure(
                generation,
                ReviewDecisionReplayCodes.FinalGenerationMismatch,
                records.Count,
                records[records.Count - 1].DecisionEvent.EventId);

            ReviewDecisionReplayException prior = Assert.Throws<ReviewDecisionReplayException>(() =>
                ReviewDecisionReplayEngine.Replay(
                    history.Catalog,
                    records,
                    history.GenerationId,
                    StableId.Compute("wrong-final-prior")));
            Assert.Equal(ReviewDecisionReplayCodes.FinalGenerationMismatch, prior.Code);
        }

        [Fact]
        public void Replay_rejects_tampering_across_request_derived_and_canonical_event_fields()
        {
            ReviewDecisionGeneration selected = ReviewDecisionEngine.Apply(
                ReviewDecisionEngine.Create(Catalog()),
                SubjectRequest(
                    ReviewDecisionEngine.Create(Catalog()),
                    "subject-1",
                    ReviewDecisionStatus.UseCandidate,
                    "observation-1-candidate",
                    When,
                    "Original note"));
            ReviewDecisionReplayRecord original =
                Assert.Single(ReviewDecisionReplayEngine.Capture(selected));
            DecisionEvent source = original.DecisionEvent;
            DecisionEvent[] tampered =
            {
                Clone(source, eventId: StableId.Compute("forged-event")),
                Clone(source, requestId: StableId.Compute("forged-request")),
                Clone(source, projectId: "another-project"),
                Clone(source, note: "Changed note"),
                Clone(source, reviewer: new ReviewerSnapshot("Someone else")),
                Clone(source, createdUtc: source.CreatedUtc.AddSeconds(1)),
                Clone(source, selectedObservationId: "observation-1-master"),
                Clone(source, supportingContentHashes: new[] { StableId.Compute("forged-content") }),
                Clone(
                    source,
                    locations: new[]
                    {
                        new DecisionLocationContext("selected", "candidate", address: "Z99"),
                    }),
                Clone(source, priorDecision: ReviewDecisionStatus.Equivalent),
                Clone(source, supersededEventId: StableId.Compute("forged-superseded")),
            };

            foreach (DecisionEvent changed in tampered)
            {
                ReviewDecisionReplayRecord record = Record(original, decisionEvent: changed);
                ReviewDecisionReplayException exception = Assert.Throws<ReviewDecisionReplayException>(() =>
                    ReviewDecisionReplayEngine.Replay(
                        selected.Catalog,
                        new[] { record },
                        selected.GenerationId,
                        selected.PriorGenerationId));
                Assert.Contains(
                    exception.Code,
                    new[]
                    {
                        ReviewDecisionReplayCodes.EventRejected,
                        ReviewDecisionReplayCodes.EventPayloadMismatch,
                    });
                Assert.Equal(0, exception.Sequence);
                Assert.Equal(changed.EventId, exception.EventId);
            }
        }

        [Fact]
        public void Replay_reapplies_selection_group_and_batch_policy_instead_of_trusting_event_payloads()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionRequest reviewOnlySelection = ReviewDecisionRequest.ForSubject(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "subject-3",
                ReviewDecisionStatus.UseCandidate,
                Reviewer,
                When,
                selectedObservationId: "observation-3-candidate");
            ReviewDecisionRequest wrongMaster = ReviewDecisionRequest.ForSubject(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "subject-1",
                ReviewDecisionStatus.UseMaster,
                Reviewer,
                When,
                selectedObservationId: "observation-1-candidate");
            ReviewDecisionRequest changedGroup = ReviewDecisionRequest.ForGroup(
                root.ProjectId,
                root.BaseGenerationId,
                root.MappingGenerationId,
                root.GenerationId,
                "group-23",
                new[] { "subject-1", "subject-2" },
                ReviewDecisionStatus.Equivalent,
                Reviewer,
                When);

            foreach (ReviewDecisionRequest request in new[]
                     {
                         reviewOnlySelection,
                         wrongMaster,
                         changedGroup,
                     })
            {
                DecisionEvent forged = EventFromRequest(request);
                var record = new ReviewDecisionReplayRecord(
                    ReviewDecisionEventContractVersions.V1,
                    0,
                    root.GenerationId,
                    StableId.Compute("forged-new", request.RequestId),
                    forged);
                ReviewDecisionReplayException exception = Assert.Throws<ReviewDecisionReplayException>(() =>
                    ReviewDecisionReplayEngine.Replay(
                        root.Catalog,
                        new[] { record },
                        record.NewGenerationId,
                        root.GenerationId));
                AssertReplayFailure(
                    exception,
                    ReviewDecisionReplayCodes.EventRejected,
                    0,
                    forged.EventId);
            }
        }

        [Fact]
        public void Replay_rejects_a_batch_with_mixed_effective_event_provenance()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionGeneration first = ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(root, "subject-1", ReviewDecisionStatus.Equivalent, "", When));
            ReviewDecisionGeneration second = ReviewDecisionEngine.Apply(
                first,
                SubjectRequest(
                    first,
                    "subject-2",
                    ReviewDecisionStatus.Equivalent,
                    "",
                    When.AddMinutes(1)));
            List<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(second).ToList();
            ReviewDecisionRequest mixed = ReviewDecisionRequest.ForSubjects(
                second.ProjectId,
                second.BaseGenerationId,
                second.MappingGenerationId,
                second.GenerationId,
                new[] { "subject-1", "subject-2" },
                ReviewDecisionStatus.NeedsInvestigation,
                Reviewer,
                When.AddMinutes(2));
            DecisionEvent forged = EventFromRequest(
                mixed,
                ReviewDecisionStatus.Equivalent,
                first.Events[0].EventId);
            records.Add(new ReviewDecisionReplayRecord(
                ReviewDecisionEventContractVersions.V1,
                2,
                second.GenerationId,
                StableId.Compute("forged-mixed-new"),
                forged));

            ReviewDecisionReplayException exception = Assert.Throws<ReviewDecisionReplayException>(() =>
                ReviewDecisionReplayEngine.Replay(
                    second.Catalog,
                    records,
                    records[2].NewGenerationId,
                    second.GenerationId));
            AssertReplayFailure(
                exception,
                ReviewDecisionReplayCodes.EventRejected,
                2,
                forged.EventId);
            Assert.Contains("mixed prior decision provenance", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Capture_returns_an_immutable_record_collection_independent_of_the_source_generation()
        {
            ReviewDecisionGeneration history = ThreeEventHistory();
            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(history);

            Assert.Throws<NotSupportedException>(() =>
                ((IList<ReviewDecisionReplayRecord>)records).Clear());
            Assert.All(
                records,
                record => Assert.NotSame(
                    history.Events[record.Sequence],
                    record.DecisionEvent));
            Assert.Equal(3, history.Events.Count);
        }

        [Fact]
        public void High_cardinality_capture_and_replay_materialize_once_and_update_only_affected_subjects()
        {
            const int subjectCount = 384;
            const int eventCount = 384;
            ReviewDecisionGeneration generation = ReviewDecisionEngine.Create(
                LargeCatalog(subjectCount));
            for (int index = 0; index < eventCount; index++)
            {
                string subjectId = LargeSubjectId(index);
                generation = ReviewDecisionEngine.Apply(
                    generation,
                    ReviewDecisionRequest.ForSubject(
                        generation.ProjectId,
                        generation.BaseGenerationId,
                        generation.MappingGenerationId,
                        generation.GenerationId,
                        subjectId,
                        ReviewDecisionStatus.Equivalent,
                        ReviewerSnapshot.Empty,
                        When.AddTicks(index + 1)));
            }

            var captureMetrics = new ReviewDecisionReplayMetrics();
            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(
                    generation,
                    CancellationToken.None,
                    captureMetrics);
            AssertLinearMetrics(captureMetrics, subjectCount, eventCount);

            var replayMetrics = new ReviewDecisionReplayMetrics();
            ReviewDecisionGeneration replayed = ReviewDecisionReplayEngine.Replay(
                generation.Catalog,
                records,
                generation.GenerationId,
                generation.PriorGenerationId,
                CancellationToken.None,
                replayMetrics);
            AssertLinearMetrics(replayMetrics, subjectCount, eventCount);
            Assert.Equal(generation.GenerationId, replayed.GenerationId);
            Assert.Equal(subjectCount, replayed.ActiveDecisions.Count);
        }

        [Fact]
        public void Selection_heavy_replay_uses_one_indexed_master_lookup_per_event()
        {
            const int subjectCount = 384;
            const int eventCount = 256;
            ReviewDecisionCatalog catalog = LargeSelectionCatalog(subjectCount);
            Assert.True(catalog.TryGetObservation(
                LargeSubjectId(0),
                "master",
                out Observation indexedMaster));
            Assert.Equal(
                "observation-" + LargeSubjectId(0) + "-master",
                indexedMaster.ObservationId);

            ReviewDecisionGeneration generation = ReviewDecisionEngine.Create(catalog);
            for (int index = 0; index < eventCount; index++)
            {
                bool candidate = index % 2 == 0;
                generation = ReviewDecisionEngine.Apply(
                    generation,
                    ReviewDecisionRequest.ForSubject(
                        generation.ProjectId,
                        generation.BaseGenerationId,
                        generation.MappingGenerationId,
                        generation.GenerationId,
                        LargeSubjectId(0),
                        candidate
                            ? ReviewDecisionStatus.UseCandidate
                            : ReviewDecisionStatus.UseMaster,
                        ReviewerSnapshot.Empty,
                        When.AddTicks(index + 1),
                        selectedObservationId:
                            "observation-" + LargeSubjectId(0) +
                            (candidate ? "-candidate" : "-master")));
            }

            var captureMetrics = new ReviewDecisionReplayMetrics();
            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(
                    generation,
                    CancellationToken.None,
                    captureMetrics);
            AssertSelectionMetrics(captureMetrics, subjectCount, eventCount);

            var replayMetrics = new ReviewDecisionReplayMetrics();
            ReviewDecisionGeneration replayed = ReviewDecisionReplayEngine.Replay(
                catalog,
                records,
                generation.GenerationId,
                generation.PriorGenerationId,
                CancellationToken.None,
                replayMetrics);
            AssertSelectionMetrics(replayMetrics, subjectCount, eventCount);
            Assert.Equal(generation.GenerationId, replayed.GenerationId);
            Assert.Single(replayed.ActiveDecisions);
        }

        [Fact]
        public void Replay_and_capture_honor_cancellation_without_materializing_a_partial_generation()
        {
            ReviewDecisionGeneration generation = ThreeEventHistory();
            IReadOnlyList<ReviewDecisionReplayRecord> records =
                ReviewDecisionReplayEngine.Capture(generation);
            using var replayCancellation = new CancellationTokenSource();
            var replayMetrics = new ReviewDecisionReplayMetrics();
            IEnumerable<ReviewDecisionReplayRecord> cancelDuringEnumeration()
            {
                for (int index = 0; index < records.Count; index++)
                {
                    if (index == 1) replayCancellation.Cancel();
                    yield return records[index];
                }
            }

            Assert.Throws<OperationCanceledException>(() =>
                ReviewDecisionReplayEngine.Replay(
                    generation.Catalog,
                    cancelDuringEnumeration(),
                    generation.GenerationId,
                    generation.PriorGenerationId,
                    replayCancellation.Token,
                    replayMetrics));
            Assert.Equal(1, replayMetrics.ProcessedEventCount);
            Assert.Equal(0, replayMetrics.FullGenerationMaterializationCount);

            using var captureCancellation = new CancellationTokenSource();
            captureCancellation.Cancel();
            var captureMetrics = new ReviewDecisionReplayMetrics();
            Assert.Throws<OperationCanceledException>(() =>
                ReviewDecisionReplayEngine.Capture(
                    generation,
                    captureCancellation.Token,
                    captureMetrics));
            Assert.Equal(0, captureMetrics.ProcessedEventCount);
            Assert.Equal(0, captureMetrics.FullGenerationMaterializationCount);
        }

        private static ReviewDecisionGeneration ThreeEventHistory()
        {
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(Catalog());
            ReviewDecisionGeneration first = ReviewDecisionEngine.Apply(
                root,
                SubjectRequest(
                    root,
                    "subject-1",
                    ReviewDecisionStatus.UseCandidate,
                    "observation-1-candidate",
                    When));
            ReviewDecisionGeneration second = ReviewDecisionEngine.Apply(
                first,
                ReviewDecisionRequest.ForGroup(
                    first.ProjectId,
                    first.BaseGenerationId,
                    first.MappingGenerationId,
                    first.GenerationId,
                    "group-23",
                    new[] { "subject-2", "subject-3" },
                    ReviewDecisionStatus.Equivalent,
                    Reviewer,
                    When.AddMinutes(1)));
            return ReviewDecisionEngine.Apply(
                second,
                ReviewDecisionRequest.ClearSubject(
                    second.ProjectId,
                    second.BaseGenerationId,
                    second.MappingGenerationId,
                    second.GenerationId,
                    "subject-1",
                    Reviewer,
                    When.AddMinutes(2)));
        }

        private static void AssertLinearMetrics(
            ReviewDecisionReplayMetrics metrics,
            int subjectCount,
            int eventCount)
        {
            Assert.Equal(subjectCount, metrics.InitializedSubjectCount);
            Assert.Equal(eventCount, metrics.ProcessedEventCount);
            Assert.Equal(eventCount, metrics.AffectedSubjectUpdateCount);
            Assert.Equal(0, metrics.IndexedMasterObservationLookupCount);
            Assert.Equal(1, metrics.FullGenerationMaterializationCount);
        }

        private static void AssertSelectionMetrics(
            ReviewDecisionReplayMetrics metrics,
            int subjectCount,
            int eventCount)
        {
            Assert.Equal(subjectCount, metrics.InitializedSubjectCount);
            Assert.Equal(eventCount, metrics.ProcessedEventCount);
            Assert.Equal(eventCount, metrics.AffectedSubjectUpdateCount);
            Assert.Equal(eventCount, metrics.IndexedMasterObservationLookupCount);
            Assert.Equal(1, metrics.FullGenerationMaterializationCount);
        }

        private static ReviewDecisionCatalog LargeCatalog(int subjectCount)
        {
            var subjects = new List<ReviewDecisionSubject>(subjectCount);
            var observations = new List<Observation>(subjectCount * 2);
            for (int index = 0; index < subjectCount; index++)
            {
                string subjectId = LargeSubjectId(index);
                subjects.Add(new ReviewDecisionSubject(
                    subjectId,
                    "logical-" + subjectId,
                    ComparisonAspect.Structure,
                    ReviewDecisionSubjectPolicy.ReviewOnly));
                observations.Add(Observation(
                    "observation-" + subjectId + "-master",
                    subjectId,
                    "master",
                    subjectId + "-master"));
                observations.Add(Observation(
                    "observation-" + subjectId + "-candidate",
                    subjectId,
                    "candidate",
                    subjectId + "-candidate"));
            }
            return new ReviewDecisionCatalog(
                "large-project",
                "large-evidence",
                "large-mapping",
                "master",
                subjects,
                observations);
        }

        private static ReviewDecisionCatalog LargeSelectionCatalog(int subjectCount)
        {
            var subjects = new List<ReviewDecisionSubject>(subjectCount);
            var observations = new List<Observation>(subjectCount * 2);
            for (int index = 0; index < subjectCount; index++)
            {
                string subjectId = LargeSubjectId(index);
                subjects.Add(index == 0
                    ? new ReviewDecisionSubject(
                        subjectId,
                        "logical-" + subjectId,
                        ComparisonAspect.Value,
                        ReviewDecisionSubjectPolicy.SelectableAndMergeable)
                    : new ReviewDecisionSubject(
                        subjectId,
                        "logical-" + subjectId,
                        ComparisonAspect.Structure,
                        ReviewDecisionSubjectPolicy.ReviewOnly));
                observations.Add(Observation(
                    "observation-" + subjectId + "-master",
                    subjectId,
                    "master",
                    subjectId + "-master"));
                observations.Add(Observation(
                    "observation-" + subjectId + "-candidate",
                    subjectId,
                    "candidate",
                    subjectId + "-candidate"));
            }
            return new ReviewDecisionCatalog(
                "large-selection-project",
                "large-selection-evidence",
                "large-selection-mapping",
                "master",
                subjects,
                observations);
        }

        private static string LargeSubjectId(int index) =>
            "subject-" + index.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);

        private static ReviewDecisionGeneration Replay(
            ReviewDecisionGeneration expected,
            IEnumerable<ReviewDecisionReplayRecord> records) =>
            ReviewDecisionReplayEngine.Replay(
                expected.Catalog,
                records,
                expected.GenerationId,
                expected.PriorGenerationId);

        private static IEnumerable<ReviewDecisionReplayRecord> Replace(
            IReadOnlyList<ReviewDecisionReplayRecord> records,
            int index,
            ReviewDecisionReplayRecord replacement) =>
            records.Select((item, itemIndex) => itemIndex == index ? replacement : item);

        private static ReviewDecisionReplayRecord Record(
            ReviewDecisionReplayRecord source,
            int? contractVersion = null,
            string? priorGenerationId = null,
            string? newGenerationId = null,
            DecisionEvent? decisionEvent = null) =>
            new ReviewDecisionReplayRecord(
                contractVersion ?? source.ContractVersion,
                source.Sequence,
                priorGenerationId ?? source.PriorGenerationId,
                newGenerationId ?? source.NewGenerationId,
                decisionEvent ?? source.DecisionEvent);

        private static void AssertReplayFailure(
            ReviewDecisionReplayException exception,
            string code,
            int sequence,
            string eventId)
        {
            Assert.Equal(code, exception.Code);
            Assert.Equal(sequence, exception.Sequence);
            Assert.Equal(eventId, exception.EventId);
            Assert.Contains(code, exception.Message, StringComparison.Ordinal);
            Assert.Contains("sequence " + sequence, exception.Message, StringComparison.Ordinal);
        }

        private static DecisionEvent EventFromRequest(
            ReviewDecisionRequest request,
            ReviewDecisionStatus priorDecision = ReviewDecisionStatus.Unreviewed,
            string supersededEventId = "") =>
            new DecisionEvent(
                StableId.Compute("forged-replay-event", request.RequestId),
                request.ProjectId,
                request.BaseGenerationId,
                request.MappingGenerationId,
                priorDecision,
                request.NewDecision,
                request.CreatedUtc,
                request.SelectionScope,
                request.RunId,
                request.SubjectId,
                request.GroupId,
                request.AffectedSubjectIds,
                request.SelectedObservationId,
                request.SelectedObservationId.Length == 0 ? Array.Empty<string>() : new[] { "candidate" },
                request.SelectedObservationId.Length == 0
                    ? Array.Empty<string>()
                    : new[] { StableId.Compute("forged-support") },
                request.Reviewer,
                request.AddInBuild,
                request.Note,
                supersededEventId,
                Array.Empty<DecisionLocationContext>(),
                request.RequestId);

        private static DecisionEvent Clone(
            DecisionEvent source,
            string? eventId = null,
            string? requestId = null,
            string? projectId = null,
            ReviewDecisionStatus? priorDecision = null,
            string? selectedObservationId = null,
            IReadOnlyList<string>? supportingContentHashes = null,
            ReviewerSnapshot? reviewer = null,
            DateTimeOffset? createdUtc = null,
            string? note = null,
            string? supersededEventId = null,
            IReadOnlyList<DecisionLocationContext>? locations = null) =>
            new DecisionEvent(
                eventId ?? source.EventId,
                projectId ?? source.ProjectId,
                source.EvidenceGenerationId,
                source.MappingGenerationId,
                priorDecision ?? source.PriorDecision,
                source.NewDecision,
                createdUtc ?? source.CreatedUtc,
                source.SelectionScope,
                source.RunId,
                source.SubjectId,
                source.GroupId,
                source.AffectedSubjectIds,
                selectedObservationId ?? source.SelectedObservationId,
                source.SupportingRevisionIds,
                supportingContentHashes ?? source.SupportingContentHashes,
                reviewer ?? source.ReviewerSnapshot,
                source.AddInBuild,
                note ?? source.Note,
                supersededEventId ?? source.SupersededEventId,
                locations ?? source.Locations,
                requestId ?? source.RequestId);

        private static ReviewDecisionRequest SubjectRequest(
            ReviewDecisionGeneration generation,
            string subjectId,
            ReviewDecisionStatus status,
            string selectedObservationId,
            DateTimeOffset createdUtc,
            string note = "") =>
            ReviewDecisionRequest.ForSubject(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.GenerationId,
                subjectId,
                status,
                Reviewer,
                createdUtc,
                selectedObservationId,
                "run-1",
                "2.0",
                note);

        private static ReviewDecisionCatalog Catalog() =>
            new ReviewDecisionCatalog(
                "project-1",
                "evidence-1",
                "mapping-1",
                "master",
                new[]
                {
                    new ReviewDecisionSubject(
                        "subject-1",
                        "logical-subject-1",
                        ComparisonAspect.Formula,
                        ReviewDecisionSubjectPolicy.SelectableAndMergeable),
                    new ReviewDecisionSubject(
                        "subject-2",
                        "logical-subject-2",
                        ComparisonAspect.Value,
                        ReviewDecisionSubjectPolicy.SelectableAndMergeable),
                    new ReviewDecisionSubject(
                        "subject-3",
                        "logical-subject-3",
                        ComparisonAspect.Structure,
                        ReviewDecisionSubjectPolicy.ReviewOnly),
                },
                new[]
                {
                    Observation("observation-1-master", "subject-1", "master", "1-master"),
                    Observation("observation-1-candidate", "subject-1", "candidate", "1-candidate"),
                    Observation("observation-2-master", "subject-2", "master", "2-master"),
                    Observation("observation-2-candidate", "subject-2", "candidate", "2-candidate"),
                    Observation("observation-3-master", "subject-3", "master", "3-master"),
                    Observation("observation-3-candidate", "subject-3", "candidate", "3-candidate"),
                },
                new[]
                {
                    new ReviewDecisionGroup("group-23", new[] { "subject-3", "subject-2" }),
                });

        private static Observation Observation(
            string observationId,
            string subjectId,
            string revisionId,
            string text) =>
            new Observation(
                observationId,
                subjectId,
                revisionId,
                ObservationExactType.String,
                "hash-" + text,
                value: new CanonicalScalar(ObservationExactType.String, text),
                provenance: new ObservationProvenance(
                    sourceLabel: revisionId,
                    sheetName: "Sheet-" + subjectId,
                    address: revisionId == "master" ? "C10" : "C11",
                    physicalRow: revisionId == "master" ? 10 : 11,
                    physicalColumn: 3,
                    logicalSheetId: "logical-sheet-" + subjectId));
    }
}
