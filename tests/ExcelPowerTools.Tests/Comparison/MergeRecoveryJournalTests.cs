using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelPowerTools.Core.Comparison.Merge;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergeRecoveryJournalTests
    {
        private static readonly DateTimeOffset ObservedOffset = new DateTimeOffset(
            2026, 7, 17, 14, 15, 16, 123, TimeSpan.FromHours(2)).AddTicks(4567);
        private static readonly DateTimeOffset CommittedOffset = ObservedOffset.AddMinutes(1);

        [Fact]
        public void Four_transitions_are_exact_immutable_deterministic_and_replayable()
        {
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(Binding());
            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                Validation());
            MergeRecoveryJournalRecord published = MergeRecoveryJournal.CreateOutputPublished(
                validated,
                Observation());
            MergeRecoveryJournalRecord committed = MergeRecoveryJournal.CreateAuditCommitted(
                published,
                AuditReceipt(validated));

            MergeRecoveryJournalRecord[] records = { prepared, validated, published, committed };
            MergeRecoveryJournalReplay replay = MergeRecoveryJournal.Replay(records);

            Assert.Equal(
                new[]
                {
                    MergeRecoveryJournalState.Prepared,
                    MergeRecoveryJournalState.Validated,
                    MergeRecoveryJournalState.OutputPublished,
                    MergeRecoveryJournalState.AuditCommitted,
                },
                replay.Records.Select(record => record.State));
            Assert.Equal(new[] { 0, 1, 2, 3 }, replay.Records.Select(record => record.Sequence));
            Assert.Equal("", prepared.PriorRecordId);
            for (int index = 1; index < records.Length; index++)
                Assert.Equal(records[index - 1].RecordId, records[index].PriorRecordId);
            Assert.All(records, record => Assert.Equal(64, record.RecordId.Length));
            Assert.All(records, record => Assert.Equal(record.RecordId.ToLowerInvariant(), record.RecordId));
            Assert.Equal(MergePublicationEvidenceKind.DirectPublisher, published.PublicationEvidenceKind);
            Assert.Equal(ObservedOffset.ToUniversalTime(), published.ObservedUtc);
            Assert.Equal(CommittedOffset.ToUniversalTime(), committed.AuditReceipt.CommittedUtc);
            Assert.Same(committed, replay.Current);
            Assert.True(replay.IsAuditCommitted);

            MergeRecoveryJournalRecord repeatedPrepared = MergeRecoveryJournal.CreatePrepared(Binding());
            MergeRecoveryJournalRecord repeatedValidated = MergeRecoveryJournal.CreateValidated(
                repeatedPrepared,
                Validation());
            Assert.Equal(prepared.RecordId, repeatedPrepared.RecordId);
            Assert.Equal(validated.RecordId, repeatedValidated.RecordId);
            Assert.Equal(validated.OutputRevisionId, repeatedValidated.OutputRevisionId);
            Assert.Equal(validated.AuditCommitId, repeatedValidated.AuditCommitId);
        }

        [Fact]
        public void Restart_reconciliation_is_evidence_not_a_recovered_ledger_state()
        {
            MergeRecoveryJournalRecord validated = Validated();
            MergeRecoveryJournalRecord reconciled =
                MergeRecoveryJournal.CreateRecoveredOutputPublished(
                    validated,
                    Observation(),
                    ObservedOffset.ToUniversalTime());

            Assert.Equal(MergeRecoveryJournalState.OutputPublished, reconciled.State);
            Assert.Equal(MergePublicationEvidenceKind.RestartReconciled, reconciled.PublicationEvidenceKind);
            Assert.DoesNotContain("Recovered", Enum.GetNames(typeof(MergeRecoveryJournalState)));
            Assert.DoesNotContain("Failed", Enum.GetNames(typeof(MergeRecoveryJournalState)));
        }

        [Theory]
        [InlineData("Merged Review.xlsx")]
        [InlineData("mErGeD rEvIeW.XLSX")]
        public void Public_validated_creation_rejects_stage_destination_aliases_case_insensitively(
            string stageLeaf)
        {
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(Binding());

            MergeRecoveryJournalException exception = Assert.Throws<MergeRecoveryJournalException>(() =>
                MergeRecoveryJournal.CreateValidated(
                    prepared,
                    Validation(stageLeaf)));

            Assert.Equal(MergeRecoveryJournalCodes.LeafAlias, exception.Code);
            Assert.Equal(1, exception.Sequence);
        }

        [Theory]
        [InlineData(MergeRecoveryJournalState.Validated)]
        [InlineData(MergeRecoveryJournalState.OutputPublished)]
        [InlineData(MergeRecoveryJournalState.AuditCommitted)]
        public void Parsed_validated_and_later_records_cannot_bypass_leaf_alias_invariant(
            MergeRecoveryJournalState state)
        {
            MergeRecoveryPreparedBinding binding = Binding();
            MergeRecoveryValidationEvidence validation = Validation("merged review.XLSX");
            string outputRevisionId = MergeRecoveryJournal.ComputeOutputRevisionId(binding, validation);
            string auditCommitId = MergeRecoveryJournal.ComputeAuditCommitId(
                binding,
                validation,
                outputRevisionId);
            MergeRecoveryAuditReceipt? auditReceipt = state == MergeRecoveryJournalState.AuditCommitted
                ? new MergeRecoveryAuditReceipt(
                    auditCommitId,
                    outputRevisionId,
                    Hash('6'),
                    Hash('7'),
                    CommittedOffset)
                : null;

            MergeRecoveryJournalException exception = Assert.Throws<MergeRecoveryJournalException>(() =>
                MergeRecoveryJournal.CreateParsedRecord(
                    MergeRecoveryJournalContract.Version,
                    (int)state - 1,
                    state,
                    Hash('0'),
                    Hash('9'),
                    binding,
                    validation,
                    outputRevisionId,
                    auditCommitId,
                    state >= MergeRecoveryJournalState.OutputPublished
                        ? MergePublicationEvidenceKind.DirectPublisher
                        : (MergePublicationEvidenceKind?)null,
                    state >= MergeRecoveryJournalState.OutputPublished
                        ? ObservedOffset
                        : (DateTimeOffset?)null,
                    auditReceipt));

            Assert.Equal(MergeRecoveryJournalCodes.LeafAlias, exception.Code);
            Assert.Equal((int)state - 1, exception.Sequence);
        }

        [Fact]
        public void Codec_rejects_case_insensitive_leaf_alias_before_record_id_replay()
        {
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(Binding());
            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                Validation());
            string canonical = StrictUtf8(MergeRecoveryJournalCodec.Serialize(new[] { prepared, validated }));
            string aliased = canonical.Replace(
                "stageLeaf=\"stage-e783.tmp\"",
                "stageLeaf=\"merged review.XLSX\"");
            Assert.NotEqual(canonical, aliased);

            MergeRecoveryJournalException exception = Assert.Throws<MergeRecoveryJournalException>(() =>
                MergeRecoveryJournalCodec.Deserialize(Utf8(aliased)));

            Assert.Equal(MergeRecoveryJournalCodes.LeafAlias, exception.Code);
            Assert.Equal(1, exception.Sequence);
        }

        [Fact]
        public void Output_observation_and_audit_receipt_must_bind_exact_frozen_evidence()
        {
            MergeRecoveryJournalRecord validated = Validated();
            MergeRecoveryOutputObservation wrongHash = new MergeRecoveryOutputObservation(
                validated.DestinationLocatorId,
                validated.DestinationLeaf,
                validated.Validation.ExpectedOutputLength,
                Hash('e'),
                ObservedOffset);
            MergeRecoveryOutputObservation wrongLocator = new MergeRecoveryOutputObservation(
                "locator-other",
                validated.DestinationLeaf,
                validated.Validation.ExpectedOutputLength,
                validated.Validation.ExpectedOutputSha256,
                ObservedOffset);
            Assert.Equal(
                MergeRecoveryJournalCodes.Observation,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.CreateOutputPublished(validated, wrongHash)).Code);
            Assert.Equal(
                MergeRecoveryJournalCodes.Observation,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.CreateRecoveredOutputPublished(validated, wrongLocator)).Code);

            MergeRecoveryJournalRecord published = MergeRecoveryJournal.CreateOutputPublished(
                validated,
                Observation());
            MergeRecoveryAuditReceipt wrongAudit = new MergeRecoveryAuditReceipt(
                Hash('a'),
                validated.OutputRevisionId,
                Hash('6'),
                Hash('7'),
                CommittedOffset);
            Assert.Equal(
                MergeRecoveryJournalCodes.AuditReceipt,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.CreateAuditCommitted(published, wrongAudit)).Code);
        }

        [Fact]
        public void Replay_rejects_frozen_drift_skips_repeats_bad_prior_and_terminal_extension()
        {
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(Binding());
            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                Validation());
            MergeRecoveryJournalRecord published = MergeRecoveryJournal.CreateOutputPublished(
                validated,
                Observation());
            MergeRecoveryJournalRecord committed = MergeRecoveryJournal.CreateAuditCommitted(
                published,
                AuditReceipt(validated));

            MergeRecoveryPreparedBinding driftedBinding = Binding(destinationLeaf: "Another output.xlsx");
            MergeRecoveryValidationEvidence validation = Validation();
            string outputId = MergeRecoveryJournal.ComputeOutputRevisionId(driftedBinding, validation);
            string auditId = MergeRecoveryJournal.ComputeAuditCommitId(
                driftedBinding,
                validation,
                outputId);
            MergeRecoveryJournalRecord drifted = RawRecord(
                1,
                MergeRecoveryJournalState.Validated,
                prepared.RecordId,
                driftedBinding,
                validation,
                outputId,
                auditId);
            Assert.Equal(
                MergeRecoveryJournalCodes.FrozenFieldDrift,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.Replay(new[] { prepared, drifted })).Code);

            Assert.Equal(
                MergeRecoveryJournalCodes.Sequence,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.Replay(new[] { prepared, published })).Code);

            MergeRecoveryJournalRecord repeated = RawRecord(
                2,
                MergeRecoveryJournalState.Validated,
                validated.RecordId,
                validated.Binding,
                validated.Validation,
                validated.OutputRevisionId,
                validated.AuditCommitId);
            Assert.Equal(
                MergeRecoveryJournalCodes.State,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.Replay(new[] { prepared, validated, repeated })).Code);

            MergeRecoveryJournalRecord badPrior = RawRecord(
                1,
                MergeRecoveryJournalState.Validated,
                Hash('f'),
                prepared.Binding,
                validated.Validation,
                validated.OutputRevisionId,
                validated.AuditCommitId);
            Assert.Equal(
                MergeRecoveryJournalCodes.PriorChain,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.Replay(new[] { prepared, badPrior })).Code);

            Assert.Equal(
                MergeRecoveryJournalCodes.Transition,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.CreateAuditCommitted(committed, AuditReceipt(validated))).Code);
            Assert.Equal(
                MergeRecoveryJournalCodes.Count,
                Assert.Throws<MergeRecoveryJournalException>(() =>
                    MergeRecoveryJournal.Replay(
                        new[] { prepared, validated, published, committed, committed })).Code);
        }

        [Fact]
        public void Later_state_fields_are_forbidden_in_early_records()
        {
            MergeRecoveryPreparedBinding binding = Binding();
            Assert.Throws<ArgumentException>(() =>
                MergeRecoveryJournal.CreateParsedRecord(
                    MergeRecoveryJournalContract.Version,
                    0,
                    MergeRecoveryJournalState.Prepared,
                    Hash('0'),
                    "",
                    binding,
                    Validation(),
                    Hash('1'),
                    Hash('2'),
                    null,
                    null,
                    null));
            Assert.Throws<ArgumentException>(() =>
                MergeRecoveryJournal.CreateParsedRecord(
                    MergeRecoveryJournalContract.Version,
                    1,
                    MergeRecoveryJournalState.Validated,
                    Hash('0'),
                    Hash('1'),
                    binding,
                    Validation(),
                    Hash('2'),
                    Hash('3'),
                    MergePublicationEvidenceKind.DirectPublisher,
                    ObservedOffset,
                    null));
        }

        [Fact]
        public void Canonical_xml_round_trip_is_byte_exact_and_golden_deterministic()
        {
            IReadOnlyList<MergeRecoveryJournalRecord> records = CompleteJournal();
            byte[] first = MergeRecoveryJournalCodec.Serialize(records);
            byte[] second = MergeRecoveryJournalCodec.Serialize(CompleteJournal());
            IReadOnlyList<MergeRecoveryJournalRecord> decoded =
                MergeRecoveryJournalCodec.Deserialize(first);
            byte[] roundTrip = MergeRecoveryJournalCodec.Serialize(decoded);

            Assert.Equal(first, second);
            Assert.Equal(first, roundTrip);
            Assert.Equal(4, decoded.Count);
            Assert.Equal(
                "81d43c8f463df5346ea374a65041151a0639d91ba3465c6b0e473add1758ffcc",
                Sha256(first));
            string text = StrictUtf8(first);
            Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\"?>", text);
            Assert.DoesNotContain("\r", text);
            Assert.DoesNotContain("\n", text);
        }

        [Fact]
        public void Codec_rejects_malformed_utf8_dtd_unknown_reordered_oversized_and_deep_content()
        {
            byte[] canonical = MergeRecoveryJournalCodec.Serialize(CompleteJournal());
            string xml = StrictUtf8(canonical);

            AssertCodecRejects(new byte[] { 0xC3, 0x28 });
            int declarationEnd = xml.IndexOf("?>", StringComparison.Ordinal) + 2;
            AssertCodecRejects(Utf8(
                xml.Insert(
                    declarationEnd,
                    "<!DOCTYPE mergeRecoveryJournal [<!ENTITY x 'unsafe'>]>")));
            AssertCodecRejects(Utf8(
                xml.Replace(
                    "</mergeRecoveryJournal>",
                    "<unknown/></mergeRecoveryJournal>")));
            AssertCodecRejects(Utf8(
                xml.Replace(
                    "contractVersion=\"1\" sequence=\"0\"",
                    "sequence=\"0\" contractVersion=\"1\"")));
            AssertCodecRejects(new byte[MergeRecoveryJournalContract.MaxJournalBytes + 1]);

            var deep = new StringBuilder();
            for (int index = 0; index < 1024; index++) deep.Append("<unknown>");
            for (int index = 0; index < 1024; index++) deep.Append("</unknown>");
            AssertCodecRejects(Utf8(
                xml.Replace(
                    "</mergeRecoveryJournal>",
                    deep + "</mergeRecoveryJournal>")));
        }

        [Fact]
        public void Hashes_and_timestamps_are_normalized_to_lowercase_canonical_utc()
        {
            MergeRecoveryJournalRecord[] records = CompleteJournal().ToArray();
            Assert.Equal(Hash('a'), records[0].BaseSha256);
            Assert.Equal(Hash('b'), records[0].SourceProjectStateRoot);
            Assert.Equal(Hash('c'), records[0].DecisionSetHash);
            Assert.Equal(Hash('d'), records[0].MergePlanHash);
            Assert.Equal(Hash('1'), records[1].Validation.ValidationReceiptHash);
            Assert.Equal(Hash('2'), records[1].Validation.ExpectedOutputSha256);
            Assert.Equal(TimeSpan.Zero, records[2].ObservedUtc.GetValueOrDefault().Offset);
            Assert.Equal(TimeSpan.Zero, records[3].AuditReceipt.CommittedUtc.Offset);

            string xml = StrictUtf8(MergeRecoveryJournalCodec.Serialize(records));
            Assert.Contains("observedUtc=\"2026-07-17T12:15:16.1234567Z\"", xml);
            Assert.Contains("committedUtc=\"2026-07-17T12:16:16.1234567Z\"", xml);
            Assert.DoesNotContain(HashUpper('A'), xml);
        }

        [Fact]
        public void Portable_bytes_have_validated_leaves_and_no_absolute_path_or_machine_token()
        {
            string forbiddenPath = @"C:\Users\operator\Desktop\Merged Review.xlsx";
            string forbiddenMachine = "CORP-WORKSTATION-927";
            Assert.Throws<ArgumentException>(() => Binding(projectId: forbiddenPath));
            Assert.Throws<ArgumentException>(() => Binding(destinationLeaf: forbiddenPath));
            Assert.Throws<ArgumentException>(() => Binding(projectId: @"server\share"));
            Assert.Throws<ArgumentException>(() => Binding(destinationLeaf: "CON.xlsx"));
            Assert.Throws<ArgumentException>(() => Binding(destinationLeaf: new string('x', 256)));
            Assert.Throws<ArgumentException>(() => Binding(projectId: new string('x', 257)));

            byte[] bytes = MergeRecoveryJournalCodec.Serialize(CompleteJournal());
            string portable = StrictUtf8(bytes);
            Assert.DoesNotContain(forbiddenPath, portable, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(forbiddenMachine, portable, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("C:\\", portable, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("occurrenceToken", portable, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("machineName", portable, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("volumeSerial", portable, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("destinationLocatorId=\"locator-7f5e8c\"", portable);
            Assert.Contains("destinationLeaf=\"Merged Review.xlsx\"", portable);
        }

        [Fact]
        public void Public_leaf_models_reject_xml_10_invalid_text_before_serialization()
        {
            string[] invalidFragments =
            {
                "\u0000",
                "\u0001",
                "\u0008",
                "\u000b",
                "\u000c",
                "\u000e",
                "\u001f",
                "\ud800",
                "\udfff",
                "\ufffe",
                "\uffff",
            };

            foreach (string fragment in invalidFragments)
            {
                string leaf = "unsafe-" + fragment + ".xlsx";
                ArgumentException destination = Assert.Throws<ArgumentException>(() =>
                    Binding(destinationLeaf: leaf));
                ArgumentException stage = Assert.Throws<ArgumentException>(() =>
                    Validation(leaf));
                ArgumentException observation = Assert.Throws<ArgumentException>(() =>
                    new MergeRecoveryOutputObservation(
                        "locator-7f5e8c",
                        leaf,
                        654321,
                        Hash('2'),
                        ObservedOffset));
                Assert.Equal("destinationLeaf", destination.ParamName);
                Assert.Equal("stageLeaf", stage.ParamName);
                Assert.Equal("destinationLeaf", observation.ParamName);
            }
        }

        [Fact]
        public void Xml_10_valid_supplementary_leaf_text_round_trips_canonically()
        {
            MergeRecoveryPreparedBinding binding = Binding(
                destinationLeaf: "Merged \ufffd \U0001f4c4.xlsx");
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(binding);
            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                Validation("stage-\U0001f4e6.tmp"));

            byte[] bytes = MergeRecoveryJournalCodec.Serialize(new[] { prepared, validated });
            IReadOnlyList<MergeRecoveryJournalRecord> decoded =
                MergeRecoveryJournalCodec.Deserialize(bytes);

            Assert.Equal(bytes, MergeRecoveryJournalCodec.Serialize(decoded));
            Assert.Equal("Merged \ufffd \U0001f4c4.xlsx", decoded[1].DestinationLeaf);
            Assert.Equal("stage-\U0001f4e6.tmp", decoded[1].Validation.StageLeaf);
        }

        public static IEnumerable<object[]> ReservedDeviceStems()
        {
            string[] stems =
            {
                "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$", "CLOCK$",
                "COM1", "COM9", "COM\u00b9", "COM\u00b2", "COM\u00b3",
                "LPT1", "LPT9", "LPT\u00b9", "LPT\u00b2", "LPT\u00b3",
            };
            return stems.Select(stem => new object[] { stem });
        }

        [Theory]
        [MemberData(nameof(ReservedDeviceStems))]
        public void Windows_device_stems_are_rejected_for_destination_and_stage_with_or_without_extensions(
            string stem)
        {
            foreach (string leaf in new[]
                     {
                         stem,
                         stem.ToLowerInvariant() + ".xlsx",
                         stem + " .xlsx",
                     })
            {
                Assert.Equal(
                    "destinationLeaf",
                    Assert.Throws<ArgumentException>(() => Binding(destinationLeaf: leaf)).ParamName);
                Assert.Equal(
                    "stageLeaf",
                    Assert.Throws<ArgumentException>(() => Validation(leaf)).ParamName);
            }
        }

        [Theory]
        [InlineData("COM0.xlsx")]
        [InlineData("COM10.xlsx")]
        [InlineData("LPT0.tmp")]
        [InlineData("LPT10.tmp")]
        [InlineData("CLOCK.xlsx")]
        [InlineData("CONIN.xlsx")]
        public void Nondevice_lookalike_leaves_remain_portable(string leaf)
        {
            Assert.Equal(leaf, Binding(destinationLeaf: leaf).DestinationLeaf);
            Assert.Equal(leaf, Validation(leaf).StageLeaf);
        }

        private static IReadOnlyList<MergeRecoveryJournalRecord> CompleteJournal()
        {
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(Binding());
            MergeRecoveryJournalRecord validated = MergeRecoveryJournal.CreateValidated(
                prepared,
                Validation());
            MergeRecoveryJournalRecord published = MergeRecoveryJournal.CreateOutputPublished(
                validated,
                Observation());
            MergeRecoveryJournalRecord committed = MergeRecoveryJournal.CreateAuditCommitted(
                published,
                AuditReceipt(validated));
            return new[] { prepared, validated, published, committed };
        }

        private static MergeRecoveryJournalRecord Validated()
        {
            MergeRecoveryJournalRecord prepared = MergeRecoveryJournal.CreatePrepared(Binding());
            return MergeRecoveryJournal.CreateValidated(prepared, Validation());
        }

        private static MergeRecoveryPreparedBinding Binding(
            string projectId = "project-17",
            string destinationLeaf = "Merged Review.xlsx") =>
            new MergeRecoveryPreparedBinding(
                projectId,
                "merge-42",
                "revision-master-3",
                HashUpper('B'),
                123456,
                HashUpper('A'),
                "evidence-generation-9",
                "mapping-generation-4",
                "decision-generation-11",
                HashUpper('C'),
                HashUpper('D'),
                "locator-7f5e8c",
                destinationLeaf);

        private static MergeRecoveryValidationEvidence Validation(
            string stageLeaf = "stage-e783.tmp") =>
            new MergeRecoveryValidationEvidence(
                stageLeaf,
                "merge-validation-v1",
                HashUpper('1'),
                654321,
                HashUpper('2'));

        private static MergeRecoveryOutputObservation Observation() =>
            new MergeRecoveryOutputObservation(
                "locator-7f5e8c",
                "Merged Review.xlsx",
                654321,
                HashUpper('2'),
                ObservedOffset);

        private static MergeRecoveryAuditReceipt AuditReceipt(
            MergeRecoveryJournalRecord validated) =>
            new MergeRecoveryAuditReceipt(
                validated.AuditCommitId,
                validated.OutputRevisionId,
                HashUpper('6'),
                HashUpper('7'),
                CommittedOffset);

        private static MergeRecoveryJournalRecord RawRecord(
            int sequence,
            MergeRecoveryJournalState state,
            string priorRecordId,
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation,
            string outputRevisionId,
            string auditCommitId,
            MergePublicationEvidenceKind? publicationKind = null,
            DateTimeOffset? observedUtc = null,
            MergeRecoveryAuditReceipt? auditReceipt = null)
        {
            MergeRecoveryJournalRecord provisional = MergeRecoveryJournal.CreateParsedRecord(
                MergeRecoveryJournalContract.Version,
                sequence,
                state,
                Hash('0'),
                priorRecordId,
                binding,
                validation,
                outputRevisionId,
                auditCommitId,
                publicationKind,
                observedUtc,
                auditReceipt);
            return MergeRecoveryJournal.CreateParsedRecord(
                MergeRecoveryJournalContract.Version,
                sequence,
                state,
                MergeRecoveryJournal.ComputeRecordId(provisional),
                priorRecordId,
                binding,
                validation,
                outputRevisionId,
                auditCommitId,
                publicationKind,
                observedUtc,
                auditReceipt);
        }

        private static void AssertCodecRejects(byte[] bytes) =>
            Assert.Throws<MergeRecoveryJournalException>(() =>
                MergeRecoveryJournalCodec.Deserialize(bytes));

        private static string Hash(char value) => new string(char.ToLowerInvariant(value), 64);
        private static string HashUpper(char value) => new string(char.ToUpperInvariant(value), 64);

        private static byte[] Utf8(string value) => new UTF8Encoding(false, true).GetBytes(value);
        private static string StrictUtf8(byte[] value) => new UTF8Encoding(false, true).GetString(value);

        private static string Sha256(byte[] bytes)
        {
            using (SHA256 sha = SHA256.Create())
                return string.Concat(sha.ComputeHash(bytes).Select(value =>
                    value.ToString("x2", CultureInfo.InvariantCulture)));
        }
    }
}
