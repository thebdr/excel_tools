using System;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    /// <summary>What a bounded local probe established about one recovery artifact.</summary>
    public enum MergeRecoveryArtifactObservationKind
    {
        Absent = 0,
        Observed = 1,
        Unavailable = 2
    }

    /// <summary>
    /// Path-free input to the pure restart matrix. Local paths and occurrence handles remain outside this model.
    /// </summary>
    public sealed class MergeRecoveryArtifactObservation
    {
        private MergeRecoveryArtifactObservation(
            MergeRecoveryArtifactObservationKind kind,
            long byteLength,
            string sha256,
            string reasonCode)
        {
            Kind = kind;
            ByteLength = byteLength;
            Sha256 = sha256 ?? "";
            ReasonCode = reasonCode ?? "";
        }

        public MergeRecoveryArtifactObservationKind Kind { get; }
        public long ByteLength { get; }
        public string Sha256 { get; }
        public string ReasonCode { get; }

        public static MergeRecoveryArtifactObservation Absent() =>
            new MergeRecoveryArtifactObservation(
                MergeRecoveryArtifactObservationKind.Absent,
                0,
                "",
                "");

        public static MergeRecoveryArtifactObservation Observed(long byteLength, string sha256) =>
            new MergeRecoveryArtifactObservation(
                MergeRecoveryArtifactObservationKind.Observed,
                MergeRecoveryJournalValidation.RequireLength(byteLength, nameof(byteLength)),
                MergeRecoveryJournalValidation.RequireSha256(sha256, nameof(sha256)),
                "");

        public static MergeRecoveryArtifactObservation Unavailable(string reasonCode)
        {
            return new MergeRecoveryArtifactObservation(
                MergeRecoveryArtifactObservationKind.Unavailable,
                0,
                "",
                MergeRecoveryJournalValidation.RequireIdentity(
                    reasonCode,
                    nameof(reasonCode)));
        }
    }

    public enum MergeRecoveryAction
    {
        ResumeValidation = 0,
        PublishValidatedStage = 1,
        AppendRecoveredOutputPublished = 2,
        CommitAudit = 3,
        Complete = 4,
        Block = 5
    }

    public static class MergeRecoveryReconciliationCodes
    {
        public const string ResumeValidation = "CMP-MERGE-RECOVERY-RESUME-VALIDATION";
        public const string PublishValidatedStage = "CMP-MERGE-RECOVERY-PUBLISH-VALIDATED-STAGE";
        public const string ReconcilePublishedOutput = "CMP-MERGE-RECOVERY-RECONCILE-PUBLISHED-OUTPUT";
        public const string CommitAudit = "CMP-MERGE-RECOVERY-COMMIT-AUDIT";
        public const string Complete = "CMP-MERGE-RECOVERY-COMPLETE";
        public const string DestinationUnexpected = "CMP-MERGE-RECOVERY-DESTINATION-UNEXPECTED";
        public const string DestinationMismatch = "CMP-MERGE-RECOVERY-DESTINATION-MISMATCH";
        public const string DestinationUnavailable = "CMP-MERGE-RECOVERY-DESTINATION-UNAVAILABLE";
        public const string StageMissing = "CMP-MERGE-RECOVERY-STAGE-MISSING";
        public const string StageMismatch = "CMP-MERGE-RECOVERY-STAGE-MISMATCH";
        public const string StageUnavailable = "CMP-MERGE-RECOVERY-STAGE-UNAVAILABLE";
        public const string CommittedOutputCorrupt = "CMP-MERGE-RECOVERY-COMMITTED-OUTPUT-CORRUPT";
    }

    /// <summary>One deterministic decision from the four-state restart matrix.</summary>
    public sealed class MergeRecoveryReconciliation
    {
        internal MergeRecoveryReconciliation(
            MergeRecoveryAction action,
            string code,
            string message,
            string correctiveAction)
        {
            Action = action;
            Code = code ?? "";
            Message = message ?? "";
            CorrectiveAction = correctiveAction ?? "";
        }

        public MergeRecoveryAction Action { get; }
        public string Code { get; }
        public string Message { get; }
        public string CorrectiveAction { get; }
        public bool IsBlocked => Action == MergeRecoveryAction.Block;
    }

    /// <summary>
    /// Pure restart policy. It never opens, writes, deletes, renames, or guesses an artifact. Callers must retain
    /// or re-prove the exact local occurrence immediately before carrying out the returned action.
    /// </summary>
    public static class MergeRecoveryReconciler
    {
        public static MergeRecoveryReconciliation Reconcile(
            MergeRecoveryJournalReplay journal,
            MergeRecoveryArtifactObservation stage,
            MergeRecoveryArtifactObservation destination)
        {
            if (journal == null) throw new ArgumentNullException(nameof(journal));
            if (journal.Current == null)
                throw new ArgumentException("A non-empty replayed recovery journal is required.", nameof(journal));
            if (stage == null) throw new ArgumentNullException(nameof(stage));
            if (destination == null) throw new ArgumentNullException(nameof(destination));

            MergeRecoveryJournalRecord current = journal.Current;
            switch (current.State)
            {
                case MergeRecoveryJournalState.Prepared:
                    return ReconcilePrepared(destination);
                case MergeRecoveryJournalState.Validated:
                    return ReconcileValidated(current, stage, destination);
                case MergeRecoveryJournalState.OutputPublished:
                    return ReconcileOutputPublished(current, destination);
                case MergeRecoveryJournalState.AuditCommitted:
                    return ReconcileAuditCommitted(current, destination);
                default:
                    throw new ArgumentOutOfRangeException(nameof(journal));
            }
        }

        private static MergeRecoveryReconciliation ReconcilePrepared(
            MergeRecoveryArtifactObservation destination)
        {
            if (destination.Kind == MergeRecoveryArtifactObservationKind.Absent)
                return Result(
                    MergeRecoveryAction.ResumeValidation,
                    MergeRecoveryReconciliationCodes.ResumeValidation,
                    "The destination is absent; the frozen base and plan must be revalidated before writing resumes.",
                    "Revalidate the source, decisions, mapping, plan, and any retained stage before creating Validated.");
            if (destination.Kind == MergeRecoveryArtifactObservationKind.Unavailable)
                return Block(
                    MergeRecoveryReconciliationCodes.DestinationUnavailable,
                    "The destination could not be proven absent; local reason code: " +
                    destination.ReasonCode + ".",
                    "Resolve the access or filesystem error and inspect the destination before resuming.");
            return Block(
                MergeRecoveryReconciliationCodes.DestinationUnexpected,
                "A destination exists before any expected output identity was validated.",
                "Do not overwrite, delete, or adopt it automatically; choose another destination or investigate it.");
        }

        private static MergeRecoveryReconciliation ReconcileValidated(
            MergeRecoveryJournalRecord current,
            MergeRecoveryArtifactObservation stage,
            MergeRecoveryArtifactObservation destination)
        {
            if (destination.Kind == MergeRecoveryArtifactObservationKind.Unavailable)
                return Block(
                    MergeRecoveryReconciliationCodes.DestinationUnavailable,
                    "The destination could not be inspected; local reason code: " +
                    destination.ReasonCode + ".",
                    "Resolve the access or filesystem error, then re-prove the destination occurrence.");
            if (destination.Kind == MergeRecoveryArtifactObservationKind.Observed)
            {
                if (MatchesExpected(current, destination))
                    return Result(
                        MergeRecoveryAction.AppendRecoveredOutputPublished,
                        MergeRecoveryReconciliationCodes.ReconcilePublishedOutput,
                        "The destination exactly matches the validated output identity.",
                        "Re-prove the same occurrence and append OutputPublished with RestartReconciled evidence.");
                return Block(
                    MergeRecoveryReconciliationCodes.DestinationMismatch,
                    "The destination exists but does not match the validated output length and SHA-256.",
                    "Do not overwrite or delete it automatically; quarantine and investigate the conflicting file.");
            }

            if (stage.Kind == MergeRecoveryArtifactObservationKind.Unavailable)
                return Block(
                    MergeRecoveryReconciliationCodes.StageUnavailable,
                    "The validated stage could not be inspected; local reason code: " +
                    stage.ReasonCode + ".",
                    "Resolve the access or filesystem error and re-prove the exact stage before publication.");
            if (stage.Kind == MergeRecoveryArtifactObservationKind.Absent)
                return Block(
                    MergeRecoveryReconciliationCodes.StageMissing,
                    "The destination is absent and the validated stage is missing.",
                    "Rebuild and fully revalidate a new stage under a new merge attempt; do not guess recovery.");
            if (!MatchesExpected(current, stage))
                return Block(
                    MergeRecoveryReconciliationCodes.StageMismatch,
                    "The retained stage does not match the validated output length and SHA-256.",
                    "Keep it for diagnosis and create a new merge attempt after investigating the mismatch.");

            return Result(
                MergeRecoveryAction.PublishValidatedStage,
                MergeRecoveryReconciliationCodes.PublishValidatedStage,
                "The destination is absent and the retained stage exactly matches the validated output identity.",
                "Re-prove the stage occurrence and frozen preconditions immediately before create-new publication.");
        }

        private static MergeRecoveryReconciliation ReconcileOutputPublished(
            MergeRecoveryJournalRecord current,
            MergeRecoveryArtifactObservation destination)
        {
            if (destination.Kind == MergeRecoveryArtifactObservationKind.Observed &&
                MatchesExpected(current, destination))
                return Result(
                    MergeRecoveryAction.CommitAudit,
                    MergeRecoveryReconciliationCodes.CommitAudit,
                    "The published destination still exactly matches the validated output identity.",
                    "Re-prove the same occurrence, commit the project audit, then append AuditCommitted.");

            if (destination.Kind == MergeRecoveryArtifactObservationKind.Unavailable)
                return Block(
                    MergeRecoveryReconciliationCodes.DestinationUnavailable,
                    "The published destination could not be inspected; local reason code: " +
                    destination.ReasonCode + ".",
                    "Resolve the access or filesystem error before committing any project audit.");
            return Block(
                MergeRecoveryReconciliationCodes.DestinationMismatch,
                destination.Kind == MergeRecoveryArtifactObservationKind.Absent
                    ? "The output recorded as published is now missing."
                    : "The output recorded as published no longer matches its validated identity.",
                "Do not advance the master; investigate the missing or changed output occurrence.");
        }

        private static MergeRecoveryReconciliation ReconcileAuditCommitted(
            MergeRecoveryJournalRecord current,
            MergeRecoveryArtifactObservation destination)
        {
            if (destination.Kind == MergeRecoveryArtifactObservationKind.Observed &&
                MatchesExpected(current, destination))
                return Result(
                    MergeRecoveryAction.Complete,
                    MergeRecoveryReconciliationCodes.Complete,
                    "The committed output still matches the immutable merge audit.",
                    "No recovery action is required.");

            string detail = destination.Kind == MergeRecoveryArtifactObservationKind.Unavailable
                ? " Local inspection reason code: " + destination.ReasonCode + "."
                : "";
            return Block(
                MergeRecoveryReconciliationCodes.CommittedOutputCorrupt,
                "The audit is committed, but the bound output is missing, changed, or unavailable." + detail,
                "Treat the project/output pair as corrupted and investigate; never rewrite the completed journal.");
        }

        private static bool MatchesExpected(
            MergeRecoveryJournalRecord record,
            MergeRecoveryArtifactObservation artifact)
        {
            return record.Validation != null &&
                   artifact.Kind == MergeRecoveryArtifactObservationKind.Observed &&
                   artifact.ByteLength == record.Validation.ExpectedOutputLength &&
                   string.Equals(
                       artifact.Sha256,
                       record.Validation.ExpectedOutputSha256,
                       StringComparison.Ordinal);
        }

        private static MergeRecoveryReconciliation Result(
            MergeRecoveryAction action,
            string code,
            string message,
            string correctiveAction) =>
            new MergeRecoveryReconciliation(action, code, message, correctiveAction);

        private static MergeRecoveryReconciliation Block(
            string code,
            string message,
            string correctiveAction) =>
            Result(MergeRecoveryAction.Block, code, message, correctiveAction);
    }
}
