using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Merge;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergePlanValidationReceiptTests
    {
        private static readonly DateTimeOffset When =
            new DateTimeOffset(2026, 7, 23, 14, 0, 0, TimeSpan.Zero);

        [Fact]
        public void Frozen_context_proof_rebuilds_exact_plan_and_synthesis()
        {
            Fixture fixture = CreateFixture();

            MergePlanValidationProof first =
                MergePlanValidationVerifier.ProveFrozenContext(
                    fixture.Plan,
                    fixture.Synthesis,
                    fixture.Generation);
            MergePlanValidationProof second =
                MergePlanValidationVerifier.ProveFrozenContext(
                    fixture.Plan,
                    fixture.Synthesis,
                    fixture.Generation);

            Assert.Equal(
                MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity,
                first.Check);
            Assert.Equal(
                MergePlanValidationVerifier.FrozenContextVerifierContract,
                first.VerifierContractId);
            Assert.Equal(fixture.Plan.PlanHash, first.PlanHash);
            Assert.Equal(fixture.Synthesis.SynthesisHash, first.SynthesisHash);
            Assert.Equal(first.EvidenceHash, second.EvidenceHash);
            Assert.Equal(first.ProofHash, second.ProofHash);
            Assert.Equal(64, first.ProofHash.Length);
        }

        [Fact]
        public void Unrelated_later_review_decision_invalidates_frozen_context()
        {
            Fixture fixture = CreateFixture();
            ReviewDecisionGeneration advanced = ReviewDecisionEngine.Apply(
                fixture.Generation,
                ReviewDecisionRequest.ForSubject(
                    fixture.Generation.ProjectId,
                    fixture.Generation.BaseGenerationId,
                    fixture.Generation.MappingGenerationId,
                    fixture.Generation.GenerationId,
                    "other-subject",
                    ReviewDecisionStatus.Equivalent,
                    ReviewerSnapshot.Empty,
                    When.AddMinutes(1)));

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationVerifier.ProveFrozenContext(
                        fixture.Plan,
                        fixture.Synthesis,
                        advanced));

            Assert.Equal(MergePlanValidationCodes.ContextStale, exception.Code);
            Assert.Equal(
                MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity,
                exception.Check);
            Assert.Contains("Corrective action:", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Cell_synthesis_proof_recomputes_exact_synthesis()
        {
            Fixture fixture = CreateFixture();

            MergePlanValidationProof proof =
                MergePlanValidationVerifier.ProveCellWriteSynthesis(
                    fixture.Plan,
                    fixture.Synthesis);

            Assert.Equal(MergePlanRequiredValidation.CellWriteSynthesis, proof.Check);
            Assert.Equal(
                MergePlanValidationVerifier.CellSynthesisVerifierContract,
                proof.VerifierContractId);
            Assert.Equal(fixture.Synthesis.SynthesisHash, proof.EvidenceHash);
        }

        [Fact]
        public void Complete_receipt_is_exact_deterministic_and_recovery_compatible()
        {
            Fixture fixture = CreateFixture();
            List<MergePlanValidationProof> proofs = CompleteProofs(fixture);

            MergePlanValidationReceipt first =
                MergePlanValidationReceipt.Assemble(
                    fixture.Plan,
                    fixture.Synthesis,
                    proofs.AsEnumerable().Reverse());
            MergePlanValidationReceipt second =
                MergePlanValidationReceipt.Assemble(
                    fixture.Plan,
                    fixture.Synthesis,
                    proofs);

            Assert.Equal(fixture.Plan.RequiredValidations, first.RequiredValidations);
            Assert.Equal(first.RequiredValidations, first.CoveredValidations);
            Assert.Equal(
                first.Proofs.OrderBy(item => (int)item.Check).Select(item => item.Check),
                first.Proofs.Select(item => item.Check));
            Assert.Equal(first.ReceiptHash, second.ReceiptHash);
            Assert.Equal(64, first.ReceiptHash.Length);

            var recoveryEvidence = new MergeRecoveryValidationEvidence(
                "validated-stage.xlsx",
                "merge-validation-receipt-v1",
                first.ReceiptHash,
                123,
                Hash("output"));
            Assert.Equal(
                first.ReceiptHash.ToLowerInvariant(),
                recoveryEvidence.ValidationReceiptHash);
        }

        [Fact]
        public void Missing_required_proof_blocks_receipt()
        {
            Fixture fixture = CreateFixture();
            List<MergePlanValidationProof> proofs = CompleteProofs(fixture);
            MergePlanValidationProof removed = proofs.Single(
                item => item.Check ==
                        MergePlanRequiredValidation.MasterTargetPreconditions);
            proofs.Remove(removed);

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationReceipt.Assemble(
                        fixture.Plan,
                        fixture.Synthesis,
                        proofs));

            Assert.Equal(MergePlanValidationCodes.ProofMissing, exception.Code);
            Assert.Equal(
                MergePlanRequiredValidation.MasterTargetPreconditions,
                exception.Check);
        }

        [Fact]
        public void Duplicate_proof_blocks_receipt()
        {
            Fixture fixture = CreateFixture();
            List<MergePlanValidationProof> proofs = CompleteProofs(fixture);
            proofs.Add(proofs[0]);

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationReceipt.Assemble(
                        fixture.Plan,
                        fixture.Synthesis,
                        proofs));

            Assert.Equal(MergePlanValidationCodes.ProofDuplicate, exception.Code);
        }

        [Fact]
        public void Unexpected_formula_proof_blocks_value_only_receipt()
        {
            Fixture fixture = CreateFixture();
            List<MergePlanValidationProof> proofs = CompleteProofs(fixture);
            proofs.Add(MergePlanValidationProof.CreateTrusted(
                fixture.Plan,
                fixture.Synthesis,
                MergePlanRequiredValidation.FormulaGroups,
                "test-formula-groups-v1",
                Hash("formula-groups")));

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationReceipt.Assemble(
                        fixture.Plan,
                        fixture.Synthesis,
                        proofs));

            Assert.Equal(MergePlanValidationCodes.ProofUnexpected, exception.Code);
            Assert.Equal(MergePlanRequiredValidation.FormulaGroups, exception.Check);
        }

        [Fact]
        public void Proof_from_another_plan_is_rejected()
        {
            Fixture fixture = CreateFixture();
            Fixture other = CreateFixture(selectedCandidateText: "different");
            List<MergePlanValidationProof> proofs = CompleteProofs(fixture);
            int index = proofs.FindIndex(item =>
                item.Check == MergePlanRequiredValidation.WriterSafetyEnvelope);
            proofs[index] = MergePlanValidationProof.CreateTrusted(
                other.Plan,
                other.Synthesis,
                MergePlanRequiredValidation.WriterSafetyEnvelope,
                "test-writer-v1",
                Hash("writer"));

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationReceipt.Assemble(
                        fixture.Plan,
                        fixture.Synthesis,
                        proofs));

            Assert.Equal(MergePlanValidationCodes.ProofInvalid, exception.Code);
        }

        [Fact]
        public void Formula_plan_requires_dependency_and_group_proofs()
        {
            Fixture fixture = CreateFixture(formula: true);
            Assert.True(fixture.Plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FormulaDependencies));
            Assert.True(fixture.Plan.RequiredValidations.HasFlag(
                MergePlanRequiredValidation.FormulaGroups));
            List<MergePlanValidationProof> proofs = CompleteProofs(fixture);
            proofs.RemoveAll(item =>
                item.Check == MergePlanRequiredValidation.FormulaDependencies);

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationReceipt.Assemble(
                        fixture.Plan,
                        fixture.Synthesis,
                        proofs));
            Assert.Equal(MergePlanValidationCodes.ProofMissing, exception.Code);
            Assert.Equal(
                MergePlanRequiredValidation.FormulaDependencies,
                exception.Check);

            MergePlanValidationReceipt complete =
                MergePlanValidationReceipt.Assemble(
                    fixture.Plan,
                    fixture.Synthesis,
                    CompleteProofs(fixture));
            Assert.Equal(complete.RequiredValidations, complete.CoveredValidations);
        }

        [Fact]
        public void Mismatched_synthesis_is_rejected_before_proof()
        {
            Fixture fixture = CreateFixture();
            Fixture other = CreateFixture(selectedCandidateText: "different");

            MergePlanValidationException exception =
                Assert.Throws<MergePlanValidationException>(() =>
                    MergePlanValidationVerifier.ProveCellWriteSynthesis(
                        fixture.Plan,
                        other.Synthesis));

            Assert.Equal(MergePlanValidationCodes.SynthesisMismatch, exception.Code);
        }

        [Fact]
        public void Validation_can_be_cancelled()
        {
            Fixture fixture = CreateFixture();
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                MergePlanValidationVerifier.ProveFrozenContext(
                    fixture.Plan,
                    fixture.Synthesis,
                    fixture.Generation,
                    cancellation.Token));
            Assert.Throws<OperationCanceledException>(() =>
                MergePlanValidationVerifier.ProveCellWriteSynthesis(
                    fixture.Plan,
                    fixture.Synthesis,
                    cancellation.Token));
        }

        [Fact]
        public void Arbitrary_proof_and_receipt_factories_are_not_public()
        {
            Assert.DoesNotContain(
                typeof(MergePlanValidationProof).GetConstructors(
                    BindingFlags.Instance |
                    BindingFlags.Public |
                    BindingFlags.NonPublic),
                constructor => constructor.IsPublic);
            Assert.DoesNotContain(
                typeof(MergePlanValidationProof).GetMethods(
                    BindingFlags.Static |
                    BindingFlags.Public |
                    BindingFlags.NonPublic),
                method => method.IsPublic &&
                          method.ReturnType == typeof(MergePlanValidationProof));
            Assert.DoesNotContain(
                typeof(MergePlanValidationReceipt).GetConstructors(
                    BindingFlags.Instance |
                    BindingFlags.Public |
                    BindingFlags.NonPublic),
                constructor => constructor.IsPublic);
            Assert.DoesNotContain(
                typeof(MergePlanValidationReceipt).GetMethods(
                    BindingFlags.Static |
                    BindingFlags.Public |
                    BindingFlags.NonPublic),
                method => method.IsPublic &&
                          method.ReturnType == typeof(MergePlanValidationReceipt));
        }

        private static List<MergePlanValidationProof> CompleteProofs(
            Fixture fixture)
        {
            var proofs = new List<MergePlanValidationProof>();
            foreach (MergePlanRequiredValidation check in
                     MergeValidationReceiptPolicy.EnumerateChecks(
                         fixture.Plan.RequiredValidations))
            {
                if (check ==
                    MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity)
                {
                    proofs.Add(MergePlanValidationVerifier.ProveFrozenContext(
                        fixture.Plan,
                        fixture.Synthesis,
                        fixture.Generation));
                }
                else if (check == MergePlanRequiredValidation.CellWriteSynthesis)
                {
                    proofs.Add(MergePlanValidationVerifier.ProveCellWriteSynthesis(
                        fixture.Plan,
                        fixture.Synthesis));
                }
                else
                {
                    proofs.Add(MergePlanValidationProof.CreateTrusted(
                        fixture.Plan,
                        fixture.Synthesis,
                        check,
                        "test-" + check + "-v1",
                        Hash("evidence-" + check)));
                }
            }
            return proofs;
        }

        private static Fixture CreateFixture(
            bool formula = false,
            string selectedCandidateText = "2")
        {
            ReviewDecisionSubject selectedSubject = formula
                ? new ReviewDecisionSubject(
                    "selected-subject",
                    "logical-cell-A1",
                    ComparisonAspect.Formula,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable)
                : new ReviewDecisionSubject(
                    "selected-subject",
                    "logical-cell-A1",
                    ComparisonAspect.Value,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            var otherSubject = new ReviewDecisionSubject(
                "other-subject",
                "logical-cell-B1",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            Observation selectedMaster = formula
                ? FormulaObservation(
                    selectedSubject.SubjectId,
                    "master",
                    "=OLD()",
                    "A1")
                : ValueObservation(
                    selectedSubject.SubjectId,
                    "master",
                    "1",
                    "A1");
            Observation selectedCandidate = formula
                ? FormulaObservation(
                    selectedSubject.SubjectId,
                    "candidate",
                    "=NEW(" + selectedCandidateText + ")",
                    "A1")
                : ValueObservation(
                    selectedSubject.SubjectId,
                    "candidate",
                    selectedCandidateText,
                    "A1");
            var evidence = new ReviewEvidenceGeneration(
                "project",
                "mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity(
                        "candidate",
                        0,
                        Hash("candidate-file")),
                    new ReviewEvidenceRevisionIdentity(
                        "master",
                        1,
                        Hash("master-file")),
                });
            var catalog = new ReviewDecisionCatalog(
                evidence,
                "master",
                new[] { selectedSubject, otherSubject },
                new[]
                {
                    selectedCandidate,
                    selectedMaster,
                    ValueObservation("other-subject", "candidate", "other-2", "B1"),
                    ValueObservation("other-subject", "master", "other-1", "B1"),
                });
            ReviewDecisionGeneration generation =
                ReviewDecisionEngine.Create(catalog);
            generation = ReviewDecisionEngine.Apply(
                generation,
                ReviewDecisionRequest.ForSubject(
                    generation.ProjectId,
                    generation.BaseGenerationId,
                    generation.MappingGenerationId,
                    generation.GenerationId,
                    selectedSubject.SubjectId,
                    ReviewDecisionStatus.UseCandidate,
                    ReviewerSnapshot.Empty,
                    When,
                    selectedCandidate.ObservationId));
            MergePlan plan = MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(
                    generation,
                    new[] { selectedSubject.SubjectId }));
            MergeCellWriteSynthesis synthesis =
                MergeCellWriteSynthesizer.Synthesize(plan);
            return new Fixture(generation, plan, synthesis);
        }

        private static Observation FormulaObservation(
            string subjectId,
            string revisionId,
            string formula,
            string address)
        {
            string contentHash = Hash("formula:" + formula);
            return new Observation(
                ComparisonIds.Observation(
                    subjectId,
                    revisionId,
                    ObservationExactType.Formula,
                    contentHash),
                subjectId,
                revisionId,
                ObservationExactType.Formula,
                contentHash,
                formula,
                provenance: Provenance(revisionId, address));
        }

        private static Observation ValueObservation(
            string subjectId,
            string revisionId,
            string value,
            string address)
        {
            string contentHash = Hash("value:" + value);
            return new Observation(
                ComparisonIds.Observation(
                    subjectId,
                    revisionId,
                    ObservationExactType.String,
                    contentHash),
                subjectId,
                revisionId,
                ObservationExactType.String,
                contentHash,
                value: new CanonicalScalar(
                    ObservationExactType.String,
                    value),
                provenance: Provenance(revisionId, address));
        }

        private static ObservationProvenance Provenance(
            string revisionId,
            string address) =>
            new ObservationProvenance(
                revisionId + ".xlsx",
                "Data",
                address,
                1,
                address == "A1" ? 1 : 2,
                "logical-sheet");

        private static string Hash(string value) =>
            StableId.ForBytes(Encoding.UTF8.GetBytes(value));

        private sealed class Fixture
        {
            internal Fixture(
                ReviewDecisionGeneration generation,
                MergePlan plan,
                MergeCellWriteSynthesis synthesis)
            {
                Generation = generation;
                Plan = plan;
                Synthesis = synthesis;
            }

            internal ReviewDecisionGeneration Generation { get; }
            internal MergePlan Plan { get; }
            internal MergeCellWriteSynthesis Synthesis { get; }
        }
    }
}
