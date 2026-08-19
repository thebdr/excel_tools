using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    /// <summary>
    /// Pure transition and replay engine for the portable merge-recovery ledger. The four cumulative records form
    /// a deterministic hash chain; replay derives trust from the complete chain rather than from a materialized
    /// current-state flag.
    /// </summary>
    public static class MergeRecoveryJournal
    {
        internal const string RecordIdDomain = "merge-recovery-journal-record-v1";
        internal const string OutputRevisionIdDomain = "merge-output-revision-v1";
        internal const string AuditCommitIdDomain = "merge-audit-commit-v1";
        private const string EmptyRecordId =
            "0000000000000000000000000000000000000000000000000000000000000000";

        public static MergeRecoveryJournalRecord CreatePrepared(
            MergeRecoveryPreparedBinding binding)
        {
            if (binding == null) throw new ArgumentNullException(nameof(binding));
            return CreateRecord(
                0,
                MergeRecoveryJournalState.Prepared,
                "",
                binding,
                null,
                "",
                "",
                null,
                null,
                null);
        }

        public static MergeRecoveryJournalRecord CreateValidated(
            MergeRecoveryJournalRecord prior,
            MergeRecoveryValidationEvidence validation)
        {
            RequireTransitionPrior(prior, MergeRecoveryJournalState.Prepared);
            if (validation == null) throw new ArgumentNullException(nameof(validation));
            RequireDistinctPublicationLeaves(prior.Binding, validation, 1);
            string outputRevisionId = ComputeOutputRevisionId(prior.Binding, validation);
            string auditCommitId = ComputeAuditCommitId(prior.Binding, validation, outputRevisionId);
            return CreateRecord(
                1,
                MergeRecoveryJournalState.Validated,
                prior.RecordId,
                prior.Binding,
                validation,
                outputRevisionId,
                auditCommitId,
                null,
                null,
                null);
        }

        public static MergeRecoveryJournalRecord CreateOutputPublished(
            MergeRecoveryJournalRecord prior,
            MergeRecoveryOutputObservation publisherEvidence)
        {
            RequireTransitionPrior(prior, MergeRecoveryJournalState.Validated);
            ValidateExpectedOutputObservation(prior, publisherEvidence);
            return CreateRecord(
                2,
                MergeRecoveryJournalState.OutputPublished,
                prior.RecordId,
                prior.Binding,
                prior.Validation,
                prior.OutputRevisionId,
                prior.AuditCommitId,
                MergePublicationEvidenceKind.DirectPublisher,
                publisherEvidence.ObservedUtc,
                null);
        }

        public static MergeRecoveryJournalRecord CreateRecoveredOutputPublished(
            MergeRecoveryJournalRecord prior,
            MergeRecoveryOutputObservation exactDestinationObservation)
        {
            RequireTransitionPrior(prior, MergeRecoveryJournalState.Validated);
            ValidateExpectedOutputObservation(prior, exactDestinationObservation);
            return CreateRecord(
                2,
                MergeRecoveryJournalState.OutputPublished,
                prior.RecordId,
                prior.Binding,
                prior.Validation,
                prior.OutputRevisionId,
                prior.AuditCommitId,
                MergePublicationEvidenceKind.RestartReconciled,
                exactDestinationObservation.ObservedUtc,
                null);
        }

        /// <summary>
        /// Explicit-time restart overload. The separately supplied time must describe the same sanitized exact
        /// observation, preventing an API caller from accidentally splitting one evidence capture across times.
        /// </summary>
        public static MergeRecoveryJournalRecord CreateRecoveredOutputPublished(
            MergeRecoveryJournalRecord prior,
            MergeRecoveryOutputObservation exactDestinationObservation,
            DateTimeOffset observedUtc)
        {
            if (exactDestinationObservation == null)
                throw new ArgumentNullException(nameof(exactDestinationObservation));
            DateTimeOffset normalized = MergeRecoveryJournalValidation.NormalizeTimestamp(
                observedUtc,
                nameof(observedUtc));
            if (!normalized.EqualsExact(exactDestinationObservation.ObservedUtc))
                throw new ArgumentException(
                    "The restart observation timestamp must match its exact evidence capture.",
                    nameof(observedUtc));
            return CreateRecoveredOutputPublished(prior, exactDestinationObservation);
        }

        public static MergeRecoveryJournalRecord CreateAuditCommitted(
            MergeRecoveryJournalRecord prior,
            MergeRecoveryAuditReceipt auditReceipt)
        {
            RequireTransitionPrior(prior, MergeRecoveryJournalState.OutputPublished);
            if (auditReceipt == null) throw new ArgumentNullException(nameof(auditReceipt));
            if (!string.Equals(prior.AuditCommitId, auditReceipt.AuditCommitId, StringComparison.Ordinal) ||
                !string.Equals(prior.OutputRevisionId, auditReceipt.OutputRevisionId, StringComparison.Ordinal))
                throw Failure(
                    MergeRecoveryJournalCodes.AuditReceipt,
                    prior.Sequence + 1,
                    "The audit receipt does not bind the deterministic validated audit and output identities.");
            if (auditReceipt.CommittedUtc < prior.ObservedUtc.Value)
                throw Failure(
                    MergeRecoveryJournalCodes.AuditReceipt,
                    prior.Sequence + 1,
                    "The audit commit timestamp precedes publication evidence.");

            return CreateRecord(
                3,
                MergeRecoveryJournalState.AuditCommitted,
                prior.RecordId,
                prior.Binding,
                prior.Validation,
                prior.OutputRevisionId,
                prior.AuditCommitId,
                prior.PublicationEvidenceKind,
                prior.ObservedUtc,
                auditReceipt);
        }

        /// <summary>
        /// Replays a complete or resumable prefix. A valid prefix always starts at Prepared and cannot skip,
        /// repeat, reorder, or extend the four terminal states.
        /// </summary>
        public static MergeRecoveryJournalReplay Replay(
            IEnumerable<MergeRecoveryJournalRecord> records)
        {
            if (records == null) throw new ArgumentNullException(nameof(records));
            var frozen = new List<MergeRecoveryJournalRecord>();
            foreach (MergeRecoveryJournalRecord record in records)
            {
                if (frozen.Count == MergeRecoveryJournalContract.MaxRecords)
                    throw Failure(
                        MergeRecoveryJournalCodes.Count,
                        frozen.Count,
                        "A recovery journal cannot contain more than four checkpoint records.");
                if (record == null)
                    throw Failure(
                        MergeRecoveryJournalCodes.Sequence,
                        frozen.Count,
                        "A recovery journal cannot contain a null checkpoint record.");
                frozen.Add(record);
            }
            if (frozen.Count == 0)
                throw Failure(
                    MergeRecoveryJournalCodes.Empty,
                    -1,
                    "A recovery journal must begin with a Prepared checkpoint.");

            MergeRecoveryPreparedBinding frozenBinding = null;
            MergeRecoveryValidationEvidence frozenValidation = null;
            string frozenOutputRevisionId = "";
            string frozenAuditCommitId = "";
            MergePublicationEvidenceKind? frozenPublicationKind = null;
            DateTimeOffset? frozenObservedUtc = null;

            for (int index = 0; index < frozen.Count; index++)
            {
                MergeRecoveryJournalRecord record = frozen[index];
                MergeRecoveryJournalState expectedState = (MergeRecoveryJournalState)(index + 1);
                if (record.ContractVersion != MergeRecoveryJournalContract.Version)
                    throw Failure(
                        MergeRecoveryJournalCodes.Version,
                        index,
                        "The recovery checkpoint contract version is not supported.");
                if (record.Sequence != index)
                    throw Failure(
                        MergeRecoveryJournalCodes.Sequence,
                        index,
                        "The recovery checkpoint sequence is not contiguous and zero based.");
                if (record.State != expectedState)
                    throw Failure(
                        MergeRecoveryJournalCodes.State,
                        index,
                        "The recovery journal must contain exactly Prepared, Validated, OutputPublished, then AuditCommitted.");
                if (index >= 1)
                    RequireDistinctPublicationLeaves(record.Binding, record.Validation, index);

                string expectedPrior = index == 0 ? "" : frozen[index - 1].RecordId;
                if (!string.Equals(record.PriorRecordId, expectedPrior, StringComparison.Ordinal))
                    throw Failure(
                        MergeRecoveryJournalCodes.PriorChain,
                        index,
                        "The recovery checkpoint does not name the exact preceding record ID.");

                string expectedRecordId = ComputeRecordId(record);
                if (!string.Equals(record.RecordId, expectedRecordId, StringComparison.Ordinal))
                    throw Failure(
                        MergeRecoveryJournalCodes.RecordId,
                        index,
                        "The recovery checkpoint ID disagrees with its canonical cumulative fields.");

                if (index == 0)
                {
                    frozenBinding = record.Binding;
                }
                else if (!frozenBinding.HasSameFrozenFields(record.Binding))
                {
                    throw Failure(
                        MergeRecoveryJournalCodes.FrozenFieldDrift,
                        index,
                        "A binding frozen at Prepared changed in a later checkpoint.");
                }

                if (index == 1)
                {
                    frozenValidation = record.Validation;
                    frozenOutputRevisionId = ComputeOutputRevisionId(frozenBinding, frozenValidation);
                    frozenAuditCommitId = ComputeAuditCommitId(
                        frozenBinding,
                        frozenValidation,
                        frozenOutputRevisionId);
                }
                if (index >= 1)
                {
                    if (!frozenValidation.HasSameFrozenFields(record.Validation))
                        throw Failure(
                            MergeRecoveryJournalCodes.FrozenFieldDrift,
                            index,
                            "Validation evidence changed in a later checkpoint.");
                    if (!string.Equals(
                            frozenOutputRevisionId,
                            record.OutputRevisionId,
                            StringComparison.Ordinal) ||
                        !string.Equals(frozenAuditCommitId, record.AuditCommitId, StringComparison.Ordinal))
                        throw Failure(
                            MergeRecoveryJournalCodes.DerivedId,
                            index,
                            "A deterministic output or audit identity disagrees with the frozen plan and validation evidence.");
                }

                if (index == 2)
                {
                    frozenPublicationKind = record.PublicationEvidenceKind;
                    frozenObservedUtc = record.ObservedUtc;
                }
                if (index >= 2 &&
                    (record.PublicationEvidenceKind != frozenPublicationKind ||
                     !record.ObservedUtc.Value.EqualsExact(frozenObservedUtc.Value)))
                    throw Failure(
                        MergeRecoveryJournalCodes.FrozenFieldDrift,
                        index,
                        "Publication evidence changed after OutputPublished.");

                if (index == 3)
                {
                    if (!string.Equals(
                            record.AuditReceipt.AuditCommitId,
                            frozenAuditCommitId,
                            StringComparison.Ordinal) ||
                        !string.Equals(
                            record.AuditReceipt.OutputRevisionId,
                            frozenOutputRevisionId,
                            StringComparison.Ordinal))
                        throw Failure(
                            MergeRecoveryJournalCodes.AuditReceipt,
                            index,
                            "The terminal audit receipt does not bind the deterministic validated identities.");
                }
            }

            return new MergeRecoveryJournalReplay(
                new ReadOnlyCollection<MergeRecoveryJournalRecord>(frozen));
        }

        internal static string ComputeRecordId(MergeRecoveryJournalRecord record)
        {
            if (record == null) throw new ArgumentNullException(nameof(record));
            var components = new List<string>
            {
                record.ContractVersion.ToString(CultureInfo.InvariantCulture),
                record.Sequence.ToString(CultureInfo.InvariantCulture),
                record.State.ToString(),
                record.PriorRecordId,
                record.ProjectId,
                record.MergeId,
                record.SourceMasterRevisionId,
                record.SourceProjectStateRoot,
                record.BaseByteLength.ToString(CultureInfo.InvariantCulture),
                record.BaseSha256,
                record.EvidenceGenerationId,
                record.MappingGenerationId,
                record.DecisionGenerationId,
                record.DecisionSetHash,
                record.MergePlanHash,
                record.DestinationLocatorId,
                record.DestinationLeaf,
                record.Validation == null ? "validation:absent" : "validation:present",
            };
            if (record.Validation != null)
            {
                components.Add(record.Validation.StageLeaf);
                components.Add(record.Validation.ValidationContractVersion);
                components.Add(record.Validation.ValidationReceiptHash);
                components.Add(record.Validation.ExpectedOutputLength.ToString(CultureInfo.InvariantCulture));
                components.Add(record.Validation.ExpectedOutputSha256);
                components.Add(record.OutputRevisionId);
                components.Add(record.AuditCommitId);
            }
            components.Add(record.PublicationEvidenceKind.HasValue
                ? "publication:present"
                : "publication:absent");
            if (record.PublicationEvidenceKind.HasValue)
            {
                components.Add(record.PublicationEvidenceKind.Value.ToString());
                components.Add(MergeRecoveryJournalValidation.TimestampText(record.ObservedUtc.Value));
            }
            components.Add(record.AuditReceipt == null ? "audit:absent" : "audit:present");
            if (record.AuditReceipt != null)
            {
                components.Add(record.AuditReceipt.AuditCommitId);
                components.Add(record.AuditReceipt.OutputRevisionId);
                components.Add(record.AuditReceipt.CommittedProjectStateRoot);
                components.Add(record.AuditReceipt.CommittedReviewPackageSha256);
                components.Add(MergeRecoveryJournalValidation.TimestampText(record.AuditReceipt.CommittedUtc));
            }
            return LowerStableId(RecordIdDomain, components.ToArray());
        }

        internal static string ComputeOutputRevisionId(
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation)
        {
            if (binding == null) throw new ArgumentNullException(nameof(binding));
            if (validation == null) throw new ArgumentNullException(nameof(validation));
            return LowerStableId(
                OutputRevisionIdDomain,
                binding.ProjectId,
                binding.MergeId,
                binding.SourceMasterRevisionId,
                binding.SourceProjectStateRoot,
                binding.MergePlanHash,
                validation.ValidationContractVersion,
                validation.ValidationReceiptHash,
                validation.ExpectedOutputLength.ToString(CultureInfo.InvariantCulture),
                validation.ExpectedOutputSha256);
        }

        internal static string ComputeAuditCommitId(
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation,
            string outputRevisionId)
        {
            if (binding == null) throw new ArgumentNullException(nameof(binding));
            if (validation == null) throw new ArgumentNullException(nameof(validation));
            outputRevisionId = MergeRecoveryJournalValidation.RequireSha256(
                outputRevisionId,
                nameof(outputRevisionId));
            return LowerStableId(
                AuditCommitIdDomain,
                binding.ProjectId,
                binding.MergeId,
                binding.SourceProjectStateRoot,
                binding.DecisionGenerationId,
                binding.DecisionSetHash,
                binding.MergePlanHash,
                validation.ValidationReceiptHash,
                validation.ExpectedOutputSha256,
                outputRevisionId);
        }

        internal static MergeRecoveryJournalRecord CreateParsedRecord(
            int contractVersion,
            int sequence,
            MergeRecoveryJournalState state,
            string recordId,
            string priorRecordId,
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation,
            string outputRevisionId,
            string auditCommitId,
            MergePublicationEvidenceKind? publicationEvidenceKind,
            DateTimeOffset? observedUtc,
            MergeRecoveryAuditReceipt auditReceipt)
        {
            if (state >= MergeRecoveryJournalState.Validated)
                RequireDistinctPublicationLeaves(binding, validation, sequence);
            return new MergeRecoveryJournalRecord(
                contractVersion,
                sequence,
                state,
                recordId,
                priorRecordId,
                binding,
                validation,
                outputRevisionId,
                auditCommitId,
                publicationEvidenceKind,
                observedUtc,
                auditReceipt);
        }

        private static MergeRecoveryJournalRecord CreateRecord(
            int sequence,
            MergeRecoveryJournalState state,
            string priorRecordId,
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation,
            string outputRevisionId,
            string auditCommitId,
            MergePublicationEvidenceKind? publicationEvidenceKind,
            DateTimeOffset? observedUtc,
            MergeRecoveryAuditReceipt auditReceipt)
        {
            var provisional = new MergeRecoveryJournalRecord(
                MergeRecoveryJournalContract.Version,
                sequence,
                state,
                EmptyRecordId,
                priorRecordId,
                binding,
                validation,
                outputRevisionId,
                auditCommitId,
                publicationEvidenceKind,
                observedUtc,
                auditReceipt);
            string recordId = ComputeRecordId(provisional);
            return new MergeRecoveryJournalRecord(
                MergeRecoveryJournalContract.Version,
                sequence,
                state,
                recordId,
                priorRecordId,
                binding,
                validation,
                outputRevisionId,
                auditCommitId,
                publicationEvidenceKind,
                observedUtc,
                auditReceipt);
        }

        private static void RequireTransitionPrior(
            MergeRecoveryJournalRecord prior,
            MergeRecoveryJournalState expectedState)
        {
            if (prior == null) throw new ArgumentNullException(nameof(prior));
            if (prior.ContractVersion != MergeRecoveryJournalContract.Version ||
                prior.State != expectedState ||
                prior.Sequence != (int)expectedState - 1)
                throw Failure(
                    MergeRecoveryJournalCodes.Transition,
                    prior.Sequence,
                    "The requested transition does not follow the exact recovery state sequence.");
            string expectedId = ComputeRecordId(prior);
            if (!string.Equals(prior.RecordId, expectedId, StringComparison.Ordinal))
                throw Failure(
                    MergeRecoveryJournalCodes.RecordId,
                    prior.Sequence,
                    "The transition prior checkpoint has an invalid deterministic record ID.");
            if (prior.Validation != null)
            {
                RequireDistinctPublicationLeaves(prior.Binding, prior.Validation, prior.Sequence);
                string outputRevisionId = ComputeOutputRevisionId(prior.Binding, prior.Validation);
                string auditCommitId = ComputeAuditCommitId(
                    prior.Binding,
                    prior.Validation,
                    outputRevisionId);
                if (!string.Equals(prior.OutputRevisionId, outputRevisionId, StringComparison.Ordinal) ||
                    !string.Equals(prior.AuditCommitId, auditCommitId, StringComparison.Ordinal))
                    throw Failure(
                        MergeRecoveryJournalCodes.DerivedId,
                        prior.Sequence,
                        "The transition prior has invalid deterministic output or audit identities.");
            }
        }

        private static void ValidateExpectedOutputObservation(
            MergeRecoveryJournalRecord validated,
            MergeRecoveryOutputObservation observation)
        {
            if (observation == null) throw new ArgumentNullException(nameof(observation));
            if (!string.Equals(
                    validated.DestinationLocatorId,
                    observation.DestinationLocatorId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    validated.DestinationLeaf,
                    observation.DestinationLeaf,
                    StringComparison.Ordinal) ||
                validated.Validation.ExpectedOutputLength != observation.ByteLength ||
                !string.Equals(
                    validated.Validation.ExpectedOutputSha256,
                    observation.Sha256,
                    StringComparison.Ordinal))
                throw Failure(
                    MergeRecoveryJournalCodes.Observation,
                    validated.Sequence + 1,
                    "Publication evidence does not exactly match the frozen destination locator, leaf, length, and SHA-256.");
        }

        private static void RequireDistinctPublicationLeaves(
            MergeRecoveryPreparedBinding binding,
            MergeRecoveryValidationEvidence validation,
            int sequence)
        {
            if (binding == null || validation == null)
                return;
            if (string.Equals(
                    binding.DestinationLeaf,
                    validation.StageLeaf,
                    StringComparison.OrdinalIgnoreCase))
                throw Failure(
                    MergeRecoveryJournalCodes.LeafAlias,
                    sequence,
                    "The validated stage leaf aliases the destination leaf under the supported Windows case-insensitive publication semantics.");
        }

        private static string LowerStableId(string domain, params string[] components) =>
            StableId.Compute(domain, components).ToLowerInvariant();

        private static MergeRecoveryJournalException Failure(
            string code,
            int sequence,
            string message) =>
            new MergeRecoveryJournalException(code, sequence, message);
    }
}
