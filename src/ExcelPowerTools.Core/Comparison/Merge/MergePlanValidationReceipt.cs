using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    public static class MergePlanValidationContract
    {
        public const int Version = 1;
        public const int MaxProofs = 16;
    }

    public static class MergePlanValidationCodes
    {
        public const string ContextStale = "CMP-MERGE-VALIDATION-CONTEXT-STALE";
        public const string SynthesisMismatch = "CMP-MERGE-VALIDATION-SYNTHESIS-MISMATCH";
        public const string ProofInvalid = "CMP-MERGE-VALIDATION-PROOF-INVALID";
        public const string ProofMissing = "CMP-MERGE-VALIDATION-PROOF-MISSING";
        public const string ProofDuplicate = "CMP-MERGE-VALIDATION-PROOF-DUPLICATE";
        public const string ProofUnexpected = "CMP-MERGE-VALIDATION-PROOF-UNEXPECTED";
    }

    public sealed class MergePlanValidationException : InvalidOperationException
    {
        internal MergePlanValidationException(
            string code,
            string detail,
            string correctiveAction,
            MergePlanRequiredValidation check = MergePlanRequiredValidation.None)
            : base(
                MergePlanValidation.RequireIdentity(code, nameof(code)) + ": " +
                (detail ?? "") +
                " Corrective action: " +
                (correctiveAction ?? "Rebuild and revalidate the exact merge plan."))
        {
            Code = code;
            Detail = detail ?? "";
            CorrectiveAction = correctiveAction ?? "";
            Check = check;
        }

        public string Code { get; }
        public string Detail { get; }
        public string CorrectiveAction { get; }
        public MergePlanRequiredValidation Check { get; }
    }

    /// <summary>
    /// One path-free deterministic proof emitted by a trusted validator. The proof is bound to both PlanHash and
    /// SynthesisHash so it cannot be replayed onto a different selection or cell-write compilation.
    /// </summary>
    public sealed class MergePlanValidationProof
    {
        private MergePlanValidationProof(
            string planHash,
            string synthesisHash,
            MergePlanRequiredValidation check,
            string verifierContractId,
            string evidenceHash)
        {
            PlanHash = RequireHash(planHash, nameof(planHash));
            SynthesisHash = RequireHash(synthesisHash, nameof(synthesisHash));
            MergeValidationReceiptPolicy.ValidateSingleCheck(check, nameof(check));
            Check = check;
            VerifierContractId = MergePlanValidation.RequireIdentity(
                verifierContractId,
                nameof(verifierContractId));
            EvidenceHash = RequireHash(evidenceHash, nameof(evidenceHash));
            ProofHash = StableId.Compute(
                "merge-plan-validation-proof-v1",
                MergePlanValidationContract.Version.ToString(
                    CultureInfo.InvariantCulture),
                PlanHash,
                SynthesisHash,
                ((int)Check).ToString(CultureInfo.InvariantCulture),
                VerifierContractId,
                EvidenceHash);
        }

        public string PlanHash { get; }
        public string SynthesisHash { get; }
        public MergePlanRequiredValidation Check { get; }
        public string VerifierContractId { get; }
        public string EvidenceHash { get; }
        public string ProofHash { get; }

        /// <summary>
        /// Internal trust boundary for concrete validators. There is deliberately no public arbitrary-proof
        /// factory.
        /// </summary>
        internal static MergePlanValidationProof CreateTrusted(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            MergePlanRequiredValidation check,
            string verifierContractId,
            string evidenceHash)
        {
            MergeValidationReceiptPolicy.ValidateBindings(plan, synthesis);
            return new MergePlanValidationProof(
                plan.PlanHash,
                synthesis.SynthesisHash,
                check,
                verifierContractId,
                evidenceHash);
        }

        private static string RequireHash(string value, string parameterName)
        {
            if (!StableId.IsSha256(value))
                throw new ArgumentException(
                    "A validation proof requires a SHA-256 identity.",
                    parameterName);
            return value;
        }
    }

    /// <summary>
    /// Complete exact proof set suitable for the recovery journal's ValidationReceiptHash. Construction fails
    /// unless every flag required by the frozen plan appears exactly once and no extra proof is present.
    /// </summary>
    public sealed class MergePlanValidationReceipt
    {
        private MergePlanValidationReceipt(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            IEnumerable<MergePlanValidationProof> proofs)
        {
            MergeValidationReceiptPolicy.ValidateBindings(plan, synthesis);
            var frozen = FreezeProofs(plan, synthesis, proofs);
            PlanHash = plan.PlanHash;
            DecisionSetHash = plan.DecisionSetHash;
            SynthesisHash = synthesis.SynthesisHash;
            RequiredValidations = plan.RequiredValidations;
            CoveredValidations = frozen.Aggregate(
                MergePlanRequiredValidation.None,
                (current, proof) => current | proof.Check);
            Proofs = new ReadOnlyCollection<MergePlanValidationProof>(frozen);

            var components = new List<string>
            {
                MergePlanValidationContract.Version.ToString(
                    CultureInfo.InvariantCulture),
                PlanHash,
                DecisionSetHash,
                SynthesisHash,
                ((int)RequiredValidations).ToString(CultureInfo.InvariantCulture),
                ((int)CoveredValidations).ToString(CultureInfo.InvariantCulture),
                Proofs.Count.ToString(CultureInfo.InvariantCulture),
            };
            components.AddRange(Proofs.Select(item => item.ProofHash));
            ReceiptHash = StableId.Compute(
                "merge-plan-validation-receipt-v1",
                components.ToArray());
        }

        public int ContractVersion => MergePlanValidationContract.Version;
        public string PlanHash { get; }
        public string DecisionSetHash { get; }
        public string SynthesisHash { get; }
        public MergePlanRequiredValidation RequiredValidations { get; }
        public MergePlanRequiredValidation CoveredValidations { get; }
        public IReadOnlyList<MergePlanValidationProof> Proofs { get; }
        public string ReceiptHash { get; }

        /// <summary>
        /// Internal because only trusted validators in this assembly may create the proof objects accepted here.
        /// </summary>
        internal static MergePlanValidationReceipt Assemble(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            IEnumerable<MergePlanValidationProof> proofs) =>
            new MergePlanValidationReceipt(plan, synthesis, proofs);

        private static List<MergePlanValidationProof> FreezeProofs(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            IEnumerable<MergePlanValidationProof> proofs)
        {
            if (proofs == null) throw new ArgumentNullException(nameof(proofs));
            var result = new List<MergePlanValidationProof>();
            var seen = new HashSet<MergePlanRequiredValidation>();
            foreach (MergePlanValidationProof proof in proofs)
            {
                if (proof == null)
                {
                    Fail(
                        MergePlanValidationCodes.ProofInvalid,
                        "The validation proof collection contains a null entry.",
                        "Rerun the required validators.");
                }
                if (result.Count == MergePlanValidationContract.MaxProofs)
                {
                    Fail(
                        MergePlanValidationCodes.ProofUnexpected,
                        "The validation proof count exceeds the versioned contract.",
                        "Use only one proof for each validation required by the plan.");
                }
                if (!string.Equals(
                        proof.PlanHash,
                        plan.PlanHash,
                        StringComparison.Ordinal) ||
                    !string.Equals(
                        proof.SynthesisHash,
                        synthesis.SynthesisHash,
                        StringComparison.Ordinal))
                {
                    Fail(
                        MergePlanValidationCodes.ProofInvalid,
                        "A validation proof belongs to a different plan or cell-write synthesis.",
                        "Discard stale proofs and rerun validation for this exact plan.",
                        proof.Check);
                }
                if (!plan.RequiredValidations.HasFlag(proof.Check))
                {
                    Fail(
                        MergePlanValidationCodes.ProofUnexpected,
                        "A proof was supplied for validation not required by this plan: " +
                        proof.Check + ".",
                        "Remove the unrelated proof and validate the exact required set.",
                        proof.Check);
                }
                if (!seen.Add(proof.Check))
                {
                    Fail(
                        MergePlanValidationCodes.ProofDuplicate,
                        "Validation " + proof.Check + " has more than one proof.",
                        "Retain one deterministic proof from the authoritative validator.",
                        proof.Check);
                }
                result.Add(proof);
            }

            foreach (MergePlanRequiredValidation required in
                     MergeValidationReceiptPolicy.EnumerateChecks(plan.RequiredValidations))
            {
                if (!seen.Contains(required))
                {
                    Fail(
                        MergePlanValidationCodes.ProofMissing,
                        "Required validation " + required + " has no proof.",
                        "Run that validator successfully before preparing a validated output.",
                        required);
                }
            }
            result.Sort((left, right) =>
                ((int)left.Check).CompareTo((int)right.Check));
            return result;
        }

        private static void Fail(
            string code,
            string detail,
            string correctiveAction,
            MergePlanRequiredValidation check = MergePlanRequiredValidation.None)
        {
            throw new MergePlanValidationException(
                code,
                detail,
                correctiveAction,
                check);
        }
    }

    /// <summary>Concrete deterministic validators available before any workbook write occurs.</summary>
    public static class MergePlanValidationVerifier
    {
        public const string FrozenContextVerifierContract =
            "merge-frozen-decision-mapping-context-v1";
        public const string CellSynthesisVerifierContract =
            "merge-cell-write-synthesis-proof-v1";

        public static MergePlanValidationProof ProveFrozenContext(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            ReviewDecisionGeneration currentGeneration,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (currentGeneration == null)
                throw new ArgumentNullException(nameof(currentGeneration));
            MergeValidationReceiptPolicy.ValidateBindings(plan, synthesis);
            cancellationToken.ThrowIfCancellationRequested();

            MergePlan rebuilt;
            MergeCellWriteSynthesis rebuiltSynthesis;
            try
            {
                var selections = plan.Operations.Select(item =>
                    new MergePlanSelection(
                        item.SubjectId,
                        item.DecisionEventId));
                var request = new MergePlanRequest(
                    plan.ProjectId,
                    plan.EvidenceGenerationId,
                    plan.MappingGenerationId,
                    plan.CatalogId,
                    plan.DecisionGenerationId,
                    plan.MasterRevisionId,
                    selections);
                rebuilt = MergePlanBuilder.Build(
                    currentGeneration,
                    request,
                    cancellationToken);
                rebuiltSynthesis = MergeCellWriteSynthesizer.Synthesize(
                    rebuilt,
                    cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is InvalidOperationException)
            {
                throw new MergePlanValidationException(
                    MergePlanValidationCodes.ContextStale,
                    "The frozen decision/mapping context could not be reproduced: " +
                    exception.Message,
                    "Refresh the review workspace, inspect changed decisions or mappings, and build a new plan.",
                    MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity);
            }

            if (!string.Equals(
                    rebuilt.PlanHash,
                    plan.PlanHash,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    rebuilt.DecisionSetHash,
                    plan.DecisionSetHash,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    rebuiltSynthesis.SynthesisHash,
                    synthesis.SynthesisHash,
                    StringComparison.Ordinal))
            {
                throw new MergePlanValidationException(
                    MergePlanValidationCodes.ContextStale,
                    "Rebuilding the current context produced a different plan or synthesis identity.",
                    "Discard the stale plan and build a new one from the current review generation.",
                    MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity);
            }

            string evidenceHash = StableId.Compute(
                "merge-frozen-context-proof-evidence-v1",
                currentGeneration.ProjectId,
                currentGeneration.BaseGenerationId,
                currentGeneration.MappingGenerationId,
                currentGeneration.CatalogId,
                currentGeneration.GenerationId,
                currentGeneration.Catalog.MasterRevisionId,
                rebuilt.PlanHash,
                rebuiltSynthesis.SynthesisHash);
            return MergePlanValidationProof.CreateTrusted(
                plan,
                synthesis,
                MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity,
                FrozenContextVerifierContract,
                evidenceHash);
        }

        public static MergePlanValidationProof ProveCellWriteSynthesis(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            MergeValidationReceiptPolicy.ValidateBindings(plan, synthesis);
            cancellationToken.ThrowIfCancellationRequested();
            MergeCellWriteSynthesis rebuilt;
            try
            {
                rebuilt = MergeCellWriteSynthesizer.Synthesize(
                    plan,
                    cancellationToken);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (
                exception is ArgumentException ||
                exception is InvalidOperationException)
            {
                throw new MergePlanValidationException(
                    MergePlanValidationCodes.SynthesisMismatch,
                    "The frozen cell-write synthesis could not be reproduced: " +
                    exception.Message,
                    "Rebuild cell-write synthesis from the exact plan.",
                    MergePlanRequiredValidation.CellWriteSynthesis);
            }
            if (!string.Equals(
                    rebuilt.SynthesisHash,
                    synthesis.SynthesisHash,
                    StringComparison.Ordinal))
            {
                throw new MergePlanValidationException(
                    MergePlanValidationCodes.SynthesisMismatch,
                    "Rebuilding cell-write synthesis produced a different identity.",
                    "Discard the stale synthesis and rebuild it from the exact plan.",
                    MergePlanRequiredValidation.CellWriteSynthesis);
            }
            return MergePlanValidationProof.CreateTrusted(
                plan,
                synthesis,
                MergePlanRequiredValidation.CellWriteSynthesis,
                CellSynthesisVerifierContract,
                synthesis.SynthesisHash);
        }
    }

    internal static class MergeValidationReceiptPolicy
    {
        internal const MergePlanRequiredValidation SupportedChecks =
            MergePlanRequiredValidation.SourceMasterByteIdentity |
            MergePlanRequiredValidation.ProjectAndArtifactFindings |
            MergePlanRequiredValidation.MasterTargetPreconditions |
            MergePlanRequiredValidation.CellWriteSynthesis |
            MergePlanRequiredValidation.WriterSafetyEnvelope |
            MergePlanRequiredValidation.FullOutputSemanticDiff |
            MergePlanRequiredValidation.AtomicPublication |
            MergePlanRequiredValidation.FormulaDependencies |
            MergePlanRequiredValidation.FormulaGroups |
            MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity;

        internal static void ValidateBindings(
            MergePlan plan,
            MergeCellWriteSynthesis synthesis)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            if (synthesis == null) throw new ArgumentNullException(nameof(synthesis));
            if (!StableId.IsSha256(plan.PlanHash) ||
                !StableId.IsSha256(plan.DecisionSetHash) ||
                !StableId.IsSha256(synthesis.SynthesisHash) ||
                !string.Equals(
                    synthesis.PlanHash,
                    plan.PlanHash,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    synthesis.DecisionSetHash,
                    plan.DecisionSetHash,
                    StringComparison.Ordinal) ||
                synthesis.SatisfiedValidation !=
                    MergePlanRequiredValidation.CellWriteSynthesis ||
                synthesis.RemainingRequiredValidations !=
                    (plan.RequiredValidations &
                     ~MergePlanRequiredValidation.CellWriteSynthesis) ||
                synthesis.Writes == null ||
                synthesis.Writes.Count == 0)
            {
                throw new MergePlanValidationException(
                    MergePlanValidationCodes.SynthesisMismatch,
                    "The plan and cell-write synthesis identities do not match.",
                    "Rebuild synthesis from the exact plan.");
            }
            if (plan.RequiredValidations == MergePlanRequiredValidation.None ||
                (plan.RequiredValidations & ~SupportedChecks) != 0)
            {
                throw new MergePlanValidationException(
                    MergePlanValidationCodes.ProofInvalid,
                    "The plan names no validation or an unsupported validation flag.",
                    "Rebuild the plan with the current validation contract.");
            }
        }

        internal static IEnumerable<MergePlanRequiredValidation> EnumerateChecks(
            MergePlanRequiredValidation checks)
        {
            for (int bit = 1; bit <= (int)SupportedChecks; bit <<= 1)
            {
                var candidate = (MergePlanRequiredValidation)bit;
                if ((checks & candidate) != 0)
                    yield return candidate;
            }
        }

        internal static void ValidateSingleCheck(
            MergePlanRequiredValidation check,
            string parameterName)
        {
            int value = (int)check;
            if (value <= 0 ||
                (value & (value - 1)) != 0 ||
                (check & ~SupportedChecks) != 0)
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    "A validation proof must identify exactly one supported check.");
        }

    }
}
