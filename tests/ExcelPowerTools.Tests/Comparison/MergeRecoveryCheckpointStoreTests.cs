using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelPowerTools.Core.Comparison.Merge;
using ExcelPowerTools.Core.Comparison.Publication;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergeRecoveryCheckpointStoreTests : IDisposable
    {
        private readonly string _root = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.MergeCheckpoint.Tests",
            Guid.NewGuid().ToString("N"));
        private readonly string _checkpointDirectory;
        private readonly string _outputDirectory;

        public MergeRecoveryCheckpointStoreTests()
        {
            _checkpointDirectory = Path.Combine(_root, "journal");
            _outputDirectory = Path.Combine(_root, "output");
            Directory.CreateDirectory(_checkpointDirectory);
            Directory.CreateDirectory(_outputDirectory);
        }

        [Fact]
        public void Publishes_four_cumulative_create_new_checkpoints_and_loads_exact_tip()
        {
            if (!IsSupportedCheckpointVolume()) return;
            IReadOnlyList<MergeRecoveryJournalRecord> complete = CompleteJournal();
            MergeRecoveryCheckpointStore store = Store();

            for (int count = 1; count <= complete.Count; count++)
            {
                MergeRecoveryCheckpointWriteResult result =
                    store.Append(complete.Take(count));
                Assert.Equal(MergeRecoveryCheckpointWriteStatus.Published, result.Status);
                Assert.True(result.IsSuccess);
                Assert.True(File.Exists(result.CheckpointPath));
                Assert.Equal("", result.RetainedStagePath);
            }

            MergeRecoveryCheckpointSnapshot snapshot = store.Load();
            Assert.Equal(4, snapshot.Records.Count);
            Assert.Equal(4, snapshot.CheckpointPaths.Count);
            Assert.Equal(
                MergeRecoveryJournalState.AuditCommitted,
                snapshot.Current.State);
            Assert.Equal(complete.Select(record => record.RecordId),
                snapshot.Records.Select(record => record.RecordId));

            for (int index = 0; index < complete.Count; index++)
            {
                string path = Path.Combine(
                    _checkpointDirectory,
                    MergeRecoveryCheckpointStoreContract.FileName(
                        (MergeRecoveryJournalState)(index + 1)));
                IReadOnlyList<MergeRecoveryJournalRecord> decoded =
                    MergeRecoveryJournalCodec.Deserialize(File.ReadAllBytes(path));
                Assert.Equal(index + 1, decoded.Count);
                Assert.Equal(complete[index].RecordId, decoded[index].RecordId);
            }

            Assert.Empty(Directory.GetFiles(_checkpointDirectory, "*.stage"));
            Assert.False(File.Exists(LocalBinding().StagePath));
            Assert.False(File.Exists(LocalBinding().DestinationPath));
        }

        [Fact]
        public void Exact_existing_checkpoint_is_idempotent()
        {
            if (!IsSupportedCheckpointVolume()) return;
            IReadOnlyList<MergeRecoveryJournalRecord> complete = CompleteJournal();
            MergeRecoveryCheckpointStore store = Store();
            MergeRecoveryCheckpointWriteResult first = store.Append(complete.Take(1));
            byte[] before = File.ReadAllBytes(first.CheckpointPath);

            MergeRecoveryCheckpointWriteResult second = store.Append(complete.Take(1));

            Assert.Equal(MergeRecoveryCheckpointWriteStatus.AlreadyPresent, second.Status);
            Assert.True(second.IsSuccess);
            Assert.Equal(before, File.ReadAllBytes(first.CheckpointPath));
            Assert.Empty(Directory.GetFiles(_checkpointDirectory, "*.stage"));
        }

        [Fact]
        public void Corrupt_existing_checkpoint_blocks_without_overwrite_or_stage_creation()
        {
            if (!IsSupportedCheckpointVolume()) return;
            IReadOnlyList<MergeRecoveryJournalRecord> complete = CompleteJournal();
            MergeRecoveryCheckpointStore store = Store();
            Assert.True(store.Append(complete.Take(1)).IsSuccess);
            string validatedPath = Path.Combine(
                _checkpointDirectory,
                MergeRecoveryCheckpointStoreContract.FileName(
                    MergeRecoveryJournalState.Validated));
            byte[] conflicting = { 0x01, 0x02, 0x03 };
            File.WriteAllBytes(validatedPath, conflicting);

            MergeRecoveryCheckpointWriteResult result = store.Append(complete.Take(2));

            Assert.Equal(MergeRecoveryCheckpointWriteStatus.Blocked, result.Status);
            Assert.Equal(MergeRecoveryCheckpointStoreCodes.Corrupt, result.Code);
            Assert.Equal(conflicting, File.ReadAllBytes(validatedPath));
            Assert.Empty(Directory.GetFiles(_checkpointDirectory, "*.stage"));
        }

        [Fact]
        public void Gap_and_out_of_sequence_append_are_blocked_without_guessing()
        {
            IReadOnlyList<MergeRecoveryJournalRecord> complete = CompleteJournal();
            string validatedPath = Path.Combine(
                _checkpointDirectory,
                MergeRecoveryCheckpointStoreContract.FileName(
                    MergeRecoveryJournalState.Validated));
            File.WriteAllBytes(
                validatedPath,
                MergeRecoveryJournalCodec.Serialize(complete.Take(2)));
            MergeRecoveryCheckpointStoreException gap = Assert.Throws<MergeRecoveryCheckpointStoreException>(
                () => Store().Load());
            Assert.Equal(MergeRecoveryCheckpointStoreCodes.Gap, gap.Code);

            string secondRoot = Path.Combine(_root, "second-journal");
            Directory.CreateDirectory(secondRoot);
            MergeRecoveryLocalBinding secondBinding = LocalBinding(secondRoot);
            MergeRecoveryCheckpointWriteResult sequence =
                new MergeRecoveryCheckpointStore(secondBinding).Append(complete.Take(2));
            Assert.Equal(MergeRecoveryCheckpointWriteStatus.Blocked, sequence.Status);
            Assert.Equal(MergeRecoveryCheckpointStoreCodes.Sequence, sequence.Code);
            Assert.Empty(Directory.GetFiles(secondRoot));
        }

        [Fact]
        public void Local_binding_must_match_portable_ids_and_exact_stage_destination_leaves()
        {
            MergeRecoveryJournalReplay validated =
                MergeRecoveryJournal.Replay(CompleteJournal().Take(2));
            MergeRecoveryLocalBinding correct = LocalBinding();
            new MergeRecoveryCheckpointStore(correct);

            MergeRecoveryLocalBinding wrongLocator = new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-other",
                _checkpointDirectory,
                correct.StagePath,
                correct.DestinationPath);
            MergeRecoveryCheckpointStoreException locator = Assert.Throws<MergeRecoveryCheckpointStoreException>(
                () => new MergeRecoveryCheckpointStore(wrongLocator).Append(validated.Records));
            Assert.Equal(MergeRecoveryCheckpointStoreCodes.BindingMismatch, locator.Code);

            MergeRecoveryLocalBinding wrongStage = new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-7f5e8c",
                _checkpointDirectory,
                Path.Combine(_outputDirectory, "other.stage"),
                correct.DestinationPath);
            MergeRecoveryCheckpointStoreException stage = Assert.Throws<MergeRecoveryCheckpointStoreException>(
                () => new MergeRecoveryCheckpointStore(wrongStage).Append(validated.Records));
            Assert.Equal(MergeRecoveryCheckpointStoreCodes.BindingMismatch, stage.Code);
            Assert.Empty(Directory.GetFiles(_checkpointDirectory));
        }

        [Fact]
        public void Local_binding_rejects_relative_normalized_alias_network_and_same_leaf_paths()
        {
            string stage = Path.Combine(_outputDirectory, "stage-e783.tmp");
            string destination = Path.Combine(_outputDirectory, "Merged Review.xlsx");
            Assert.Throws<ArgumentException>(() => new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-7f5e8c",
                ".",
                stage,
                destination));
            Assert.Throws<ArgumentException>(() => new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-7f5e8c",
                _checkpointDirectory + Path.DirectorySeparatorChar + ".",
                stage,
                destination));
            Assert.Throws<ArgumentException>(() => new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-7f5e8c",
                _checkpointDirectory,
                @"\\server\share\stage.tmp",
                @"\\server\share\result.xlsx"));
            Assert.Throws<ArgumentException>(() => new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-7f5e8c",
                _checkpointDirectory,
                destination.ToUpperInvariant(),
                destination));
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_root))
                    Directory.Delete(_root, true);
            }
            catch
            {
            }
        }

        private MergeRecoveryCheckpointStore Store() =>
            new MergeRecoveryCheckpointStore(LocalBinding());

        private MergeRecoveryLocalBinding LocalBinding(string? checkpointDirectory = null) =>
            new MergeRecoveryLocalBinding(
                "project-17",
                "merge-42",
                "locator-7f5e8c",
                checkpointDirectory ?? _checkpointDirectory,
                Path.Combine(_outputDirectory, "stage-e783.tmp"),
                Path.Combine(_outputDirectory, "Merged Review.xlsx"));

        private bool IsSupportedCheckpointVolume()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return false;
            string probe = Path.Combine(_checkpointDirectory, "probe.stage");
            File.WriteAllBytes(probe, new byte[] { 1 });
            ValidatedStagePublicationCapability capability =
                ValidatedStageFilePublisher.CheckCapability(
                    probe,
                    Path.Combine(_checkpointDirectory, "probe-result.tmp"));
            File.Delete(probe);
            Assert.True(capability.IsSupported, capability.Code + ": " + capability.Message);
            return true;
        }

        private static IReadOnlyList<MergeRecoveryJournalRecord> CompleteJournal()
        {
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
            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                new MergeRecoveryValidationEvidence(
                    "stage-e783.tmp",
                    "merge-validation-v1",
                    Hash('1'),
                    654321,
                    Hash('2')));
            MergeRecoveryJournalRecord published = MergeRecoveryJournal.CreateOutputPublished(
                validated,
                new MergeRecoveryOutputObservation(
                    "locator-7f5e8c",
                    "Merged Review.xlsx",
                    654321,
                    Hash('2'),
                    new DateTimeOffset(2026, 7, 23, 10, 11, 12, TimeSpan.Zero)));
            MergeRecoveryJournalRecord committed = MergeRecoveryJournal.CreateAuditCommitted(
                published,
                new MergeRecoveryAuditReceipt(
                    validated.AuditCommitId,
                    validated.OutputRevisionId,
                    Hash('6'),
                    Hash('7'),
                    new DateTimeOffset(2026, 7, 23, 10, 12, 12, TimeSpan.Zero)));
            return new[] { prepared, validated, published, committed };
        }

        private static string Hash(char value) => new string(char.ToLowerInvariant(value), 64);
    }
}
