using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Comparison.Merge;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergeRecoveryReconcilerTests
    {
        private const long ExpectedLength = 654321;
        private static readonly DateTimeOffset ObservedUtc =
            new DateTimeOffset(2026, 7, 23, 10, 11, 12, TimeSpan.Zero);

        [Fact]
        public void Prepared_resumes_only_when_destination_is_proven_absent()
        {
            MergeRecoveryJournalReplay prepared = Replay(MergeRecoveryJournalState.Prepared);

            AssertAction(
                MergeRecoveryAction.ResumeValidation,
                MergeRecoveryReconciliationCodes.ResumeValidation,
                prepared,
                Exact(),
                MergeRecoveryArtifactObservation.Absent());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationUnexpected,
                prepared,
                Exact(),
                Exact());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationUnavailable,
                prepared,
                Exact(),
                MergeRecoveryArtifactObservation.Unavailable("access-denied"));
        }

        [Fact]
        public void Validated_publishes_only_an_exact_stage_when_destination_is_absent()
        {
            MergeRecoveryJournalReplay validated = Replay(MergeRecoveryJournalState.Validated);

            AssertAction(
                MergeRecoveryAction.PublishValidatedStage,
                MergeRecoveryReconciliationCodes.PublishValidatedStage,
                validated,
                Exact(),
                MergeRecoveryArtifactObservation.Absent());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.StageMissing,
                validated,
                MergeRecoveryArtifactObservation.Absent(),
                MergeRecoveryArtifactObservation.Absent());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.StageMismatch,
                validated,
                Mismatch(),
                MergeRecoveryArtifactObservation.Absent());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.StageUnavailable,
                validated,
                MergeRecoveryArtifactObservation.Unavailable("locked"),
                MergeRecoveryArtifactObservation.Absent());
        }

        [Fact]
        public void Validated_reconciles_only_an_exact_existing_destination()
        {
            MergeRecoveryJournalReplay validated = Replay(MergeRecoveryJournalState.Validated);

            AssertAction(
                MergeRecoveryAction.AppendRecoveredOutputPublished,
                MergeRecoveryReconciliationCodes.ReconcilePublishedOutput,
                validated,
                Mismatch(),
                Exact());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationMismatch,
                validated,
                Exact(),
                Mismatch());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationUnavailable,
                validated,
                Exact(),
                MergeRecoveryArtifactObservation.Unavailable("sharing-violation"));
        }

        [Fact]
        public void Output_published_commits_audit_only_for_the_exact_destination()
        {
            MergeRecoveryJournalReplay published = Replay(MergeRecoveryJournalState.OutputPublished);

            AssertAction(
                MergeRecoveryAction.CommitAudit,
                MergeRecoveryReconciliationCodes.CommitAudit,
                published,
                MergeRecoveryArtifactObservation.Absent(),
                Exact());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationMismatch,
                published,
                Exact(),
                MergeRecoveryArtifactObservation.Absent());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationMismatch,
                published,
                Exact(),
                Mismatch());
            AssertAction(
                MergeRecoveryAction.Block,
                MergeRecoveryReconciliationCodes.DestinationUnavailable,
                published,
                Exact(),
                MergeRecoveryArtifactObservation.Unavailable("unreadable"));
        }

        [Fact]
        public void Audit_committed_is_complete_only_while_the_exact_output_remains()
        {
            MergeRecoveryJournalReplay committed = Replay(MergeRecoveryJournalState.AuditCommitted);

            AssertAction(
                MergeRecoveryAction.Complete,
                MergeRecoveryReconciliationCodes.Complete,
                committed,
                MergeRecoveryArtifactObservation.Absent(),
                Exact());
            foreach (MergeRecoveryArtifactObservation destination in new[]
                     {
                         MergeRecoveryArtifactObservation.Absent(),
                         Mismatch(),
                         MergeRecoveryArtifactObservation.Unavailable("offline"),
                     })
            {
                AssertAction(
                    MergeRecoveryAction.Block,
                    MergeRecoveryReconciliationCodes.CommittedOutputCorrupt,
                    committed,
                    Exact(),
                    destination);
            }
        }

        [Fact]
        public void Artifact_observations_are_bounded_and_canonical()
        {
            MergeRecoveryArtifactObservation exact =
                MergeRecoveryArtifactObservation.Observed(7, HashUpper('A'));
            Assert.Equal(7, exact.ByteLength);
            Assert.Equal(Hash('a'), exact.Sha256);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                MergeRecoveryArtifactObservation.Observed(-1, Hash('a')));
            Assert.Throws<ArgumentException>(() =>
                MergeRecoveryArtifactObservation.Observed(1, "bad"));
            Assert.Throws<ArgumentException>(() =>
                MergeRecoveryArtifactObservation.Unavailable(""));
            Assert.Throws<ArgumentException>(() =>
                MergeRecoveryArtifactObservation.Unavailable(new string('x', 513)));
        }

        private static void AssertAction(
            MergeRecoveryAction expectedAction,
            string expectedCode,
            MergeRecoveryJournalReplay journal,
            MergeRecoveryArtifactObservation stage,
            MergeRecoveryArtifactObservation destination)
        {
            MergeRecoveryReconciliation result =
                MergeRecoveryReconciler.Reconcile(journal, stage, destination);
            Assert.Equal(expectedAction, result.Action);
            Assert.Equal(expectedCode, result.Code);
            Assert.Equal(expectedAction == MergeRecoveryAction.Block, result.IsBlocked);
            Assert.NotEmpty(result.Message);
            Assert.NotEmpty(result.CorrectiveAction);
        }

        private static MergeRecoveryArtifactObservation Exact() =>
            MergeRecoveryArtifactObservation.Observed(ExpectedLength, Hash('2'));

        private static MergeRecoveryArtifactObservation Mismatch() =>
            MergeRecoveryArtifactObservation.Observed(ExpectedLength, Hash('e'));

        private static MergeRecoveryJournalReplay Replay(MergeRecoveryJournalState state)
        {
            var records = new List<MergeRecoveryJournalRecord>();
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(
                new MergeRecoveryPreparedBinding(
                    "project-17",
                    "merge-42",
                    "revision-master-3",
                    Hash('b'),
                    123456,
                    Hash('a'),
                    "evidence-generation-9",
                    "mapping-generation-4",
                    "decision-generation-11",
                    Hash('c'),
                    Hash('d'),
                    "locator-7f5e8c",
                    "Merged Review.xlsx"));
            records.Add(prepared);
            if (state == MergeRecoveryJournalState.Prepared)
                return MergeRecoveryJournal.Replay(records);

            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                new MergeRecoveryValidationEvidence(
                    "stage-e783.tmp",
                    "merge-validation-v1",
                    Hash('1'),
                    ExpectedLength,
                    Hash('2')));
            records.Add(validated);
            if (state == MergeRecoveryJournalState.Validated)
                return MergeRecoveryJournal.Replay(records);

            MergeRecoveryJournalRecord published = MergeRecoveryJournal.CreateOutputPublished(
                validated,
                new MergeRecoveryOutputObservation(
                    validated.DestinationLocatorId,
                    validated.DestinationLeaf,
                    ExpectedLength,
                    Hash('2'),
                    ObservedUtc));
            records.Add(published);
            if (state == MergeRecoveryJournalState.OutputPublished)
                return MergeRecoveryJournal.Replay(records);

            records.Add(MergeRecoveryJournal.CreateAuditCommitted(
                published,
                new MergeRecoveryAuditReceipt(
                    validated.AuditCommitId,
                    validated.OutputRevisionId,
                    Hash('6'),
                    Hash('7'),
                    ObservedUtc.AddMinutes(1))));
            return MergeRecoveryJournal.Replay(records);
        }

        private static string Hash(char value) => new string(char.ToLowerInvariant(value), 64);
        private static string HashUpper(char value) => new string(char.ToUpperInvariant(value), 64);
    }
}
