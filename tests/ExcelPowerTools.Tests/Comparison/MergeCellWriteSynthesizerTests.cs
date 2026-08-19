using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Merge;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergeCellWriteSynthesizerTests
    {
        private static readonly DateTimeOffset When =
            new DateTimeOffset(2026, 7, 23, 13, 0, 0, TimeSpan.Zero);

        [Fact]
        public void Formula_change_synthesizes_complete_set_formula_write()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Formula,
                    "=OLD(A1)",
                    ObservationExactType.Formula,
                    "=NEW(A1)"),
                value: null);

            MergeCellWriteSynthesis synthesis =
                MergeCellWriteSynthesizer.Synthesize(plan);
            MergeCellWrite write = Assert.Single(synthesis.Writes);

            Assert.Equal(MergeCellWriteKind.SetFormula, write.Kind);
            Assert.Equal("=NEW(A1)", write.OutputFormula);
            Assert.Equal(ObservationExactType.Blank, write.OutputValue.Type);
            Assert.NotEmpty(write.FormulaOperationId);
            Assert.Empty(write.ValueOperationId);
            Assert.Equal(
                MergeCellWriteProofRequirement.None,
                write.ProofRequirements);
            Assert.Equal(
                MergePlanRequiredValidation.CellWriteSynthesis,
                synthesis.SatisfiedValidation);
            Assert.False(synthesis.RemainingRequiredValidations.HasFlag(
                MergePlanRequiredValidation.CellWriteSynthesis));
            Assert.True(synthesis.RemainingRequiredValidations.HasFlag(
                MergePlanRequiredValidation.FullOutputSemanticDiff));
            Assert.Equal(64, synthesis.SynthesisHash.Length);
        }

        [Fact]
        public void Formula_removal_synthesizes_clear_content_hybrid()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Formula,
                    "=OLD()",
                    ObservationExactType.Blank,
                    ""),
                value: null);

            MergeCellWrite write = Assert.Single(
                MergeCellWriteSynthesizer.Synthesize(plan).Writes);

            Assert.Equal(MergeCellWriteKind.ClearContent, write.Kind);
            Assert.Empty(write.OutputFormula);
            Assert.Equal(ObservationExactType.Blank, write.OutputValue.Type);
            Assert.Equal(
                MergeCellWriteProofRequirement.None,
                write.ProofRequirements);
        }

        [Fact]
        public void Constant_to_formula_uses_formula_and_explicit_blank_value()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Blank,
                    "",
                    ObservationExactType.Formula,
                    "=A1+1"),
                Value(
                    ObservationExactType.Number,
                    "10",
                    ObservationExactType.Blank,
                    ""));

            MergeCellWrite write = Assert.Single(
                MergeCellWriteSynthesizer.Synthesize(plan).Writes);

            Assert.Equal(MergeCellWriteKind.SetFormula, write.Kind);
            Assert.Equal("=A1+1", write.OutputFormula);
            Assert.Equal(ObservationExactType.Blank, write.OutputValue.Type);
            Assert.NotEmpty(write.FormulaOperationId);
            Assert.NotEmpty(write.ValueOperationId);
            Assert.Equal(
                MergeCellWriteProofRequirement.None,
                write.ProofRequirements);
        }

        [Fact]
        public void Formula_to_constant_uses_selected_typed_value()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Formula,
                    "=OLD()",
                    ObservationExactType.Blank,
                    ""),
                Value(
                    ObservationExactType.Blank,
                    "",
                    ObservationExactType.Number,
                    "42"));

            MergeCellWrite write = Assert.Single(
                MergeCellWriteSynthesizer.Synthesize(plan).Writes);

            Assert.Equal(MergeCellWriteKind.SetConstant, write.Kind);
            Assert.Empty(write.OutputFormula);
            Assert.Equal(ObservationExactType.Number, write.OutputValue.Type);
            Assert.Equal("42", write.OutputValue.CanonicalText);
            Assert.Equal(
                MergeCellWriteProofRequirement.None,
                write.ProofRequirements);
        }

        [Fact]
        public void Value_only_change_from_nonblank_master_needs_no_formula_proof()
        {
            MergePlan plan = Plan(
                formula: null,
                Value(
                    ObservationExactType.Number,
                    "10",
                    ObservationExactType.Number,
                    "11"));

            MergeCellWriteSynthesis synthesis =
                MergeCellWriteSynthesizer.Synthesize(plan);
            MergeCellWrite write = Assert.Single(synthesis.Writes);

            Assert.Equal(MergeCellWriteKind.SetConstant, write.Kind);
            Assert.Equal("11", write.OutputValue.CanonicalText);
            Assert.Equal(
                MergeCellWriteProofRequirement.None,
                write.ProofRequirements);
            Assert.Equal(0, synthesis.WritesRequiringMasterProof);
        }

        [Fact]
        public void Value_only_change_from_blank_requires_proof_master_has_no_formula()
        {
            MergePlan plan = Plan(
                formula: null,
                Value(
                    ObservationExactType.Blank,
                    "",
                    ObservationExactType.String,
                    "chosen"));

            MergeCellWriteSynthesis synthesis =
                MergeCellWriteSynthesizer.Synthesize(plan);
            MergeCellWrite write = Assert.Single(synthesis.Writes);

            Assert.Equal(MergeCellWriteKind.SetConstant, write.Kind);
            Assert.Equal(
                MergeCellWriteProofRequirement.MasterHasNoFormula,
                write.ProofRequirements);
            Assert.Equal(1, synthesis.WritesRequiringMasterProof);
            Assert.True(synthesis.RemainingRequiredValidations.HasFlag(
                MergePlanRequiredValidation.MasterTargetPreconditions));
        }

        [Fact]
        public void Formula_addition_without_value_choice_requires_blank_master_constant_proof()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Blank,
                    "",
                    ObservationExactType.Formula,
                    "=CHOSEN()"),
                value: null);

            MergeCellWrite write = Assert.Single(
                MergeCellWriteSynthesizer.Synthesize(plan).Writes);

            Assert.Equal(MergeCellWriteKind.SetFormula, write.Kind);
            Assert.Equal(
                MergeCellWriteProofRequirement.MasterConstantValueIsBlank,
                write.ProofRequirements);
        }

        [Fact]
        public void Rejects_selected_formula_plus_nonblank_constant_hybrid()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Blank,
                    "",
                    ObservationExactType.Formula,
                    "=CHOSEN()"),
                Value(
                    ObservationExactType.Number,
                    "1",
                    ObservationExactType.Number,
                    "2"));

            MergeCellWriteSynthesisException exception =
                Assert.Throws<MergeCellWriteSynthesisException>(() =>
                    MergeCellWriteSynthesizer.Synthesize(plan));

            Assert.Equal(MergeCellWriteSynthesisCodes.AspectConflict, exception.Code);
            Assert.NotEmpty(exception.TargetId);
            Assert.Contains("Corrective action:", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Rejects_master_evidence_claiming_formula_and_nonblank_constant()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Formula,
                    "=OLD()",
                    ObservationExactType.Formula,
                    "=NEW()"),
                Value(
                    ObservationExactType.Number,
                    "1",
                    ObservationExactType.Blank,
                    ""));

            MergeCellWriteSynthesisException exception =
                Assert.Throws<MergeCellWriteSynthesisException>(() =>
                    MergeCellWriteSynthesizer.Synthesize(plan));

            Assert.Equal(
                MergeCellWriteSynthesisCodes.EvidenceInconsistent,
                exception.Code);
        }

        [Fact]
        public void Synthesis_is_deterministic_and_does_not_mutate_plan()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Formula,
                    "=OLD()",
                    ObservationExactType.Formula,
                    "=NEW()"),
                value: null);
            string originalPlanHash = plan.PlanHash;
            string[] originalOperations = plan.Operations
                .Select(item => item.OperationId)
                .ToArray();

            MergeCellWriteSynthesis first =
                MergeCellWriteSynthesizer.Synthesize(plan);
            MergeCellWriteSynthesis second =
                MergeCellWriteSynthesizer.Synthesize(plan);

            Assert.Equal(first.SynthesisHash, second.SynthesisHash);
            Assert.Equal(
                first.Writes.Select(item => item.WriteId),
                second.Writes.Select(item => item.WriteId));
            Assert.Equal(originalPlanHash, plan.PlanHash);
            Assert.Equal(
                originalOperations,
                plan.Operations.Select(item => item.OperationId));
        }

        [Fact]
        public void Null_plan_is_rejected()
        {
            Assert.Throws<ArgumentNullException>(() =>
                MergeCellWriteSynthesizer.Synthesize(null!));
        }

        [Fact]
        public void Cancellation_stops_synthesis()
        {
            MergePlan plan = Plan(
                Formula(
                    ObservationExactType.Formula,
                    "=OLD()",
                    ObservationExactType.Formula,
                    "=NEW()"),
                value: null);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() =>
                MergeCellWriteSynthesizer.Synthesize(plan, cancellation.Token));
        }

        private static MergePlan Plan(
            FormulaSpec? formula,
            ValueSpec? value)
        {
            var subjects = new List<ReviewDecisionSubject>();
            var observations = new List<Observation>();
            var selectedSubjects = new List<string>();
            if (formula != null)
            {
                subjects.Add(new ReviewDecisionSubject(
                    "formula-subject",
                    "logical-cell",
                    ComparisonAspect.Formula,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable));
                observations.Add(Observation(
                    "formula-subject",
                    ComparisonAspect.Formula,
                    "master",
                    formula.MasterType,
                    formula.MasterText));
                observations.Add(Observation(
                    "formula-subject",
                    ComparisonAspect.Formula,
                    "candidate",
                    formula.CandidateType,
                    formula.CandidateText));
                selectedSubjects.Add("formula-subject");
            }
            if (value != null)
            {
                subjects.Add(new ReviewDecisionSubject(
                    "value-subject",
                    "logical-cell",
                    ComparisonAspect.Value,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable));
                observations.Add(Observation(
                    "value-subject",
                    ComparisonAspect.Value,
                    "master",
                    value.MasterType,
                    value.MasterText));
                observations.Add(Observation(
                    "value-subject",
                    ComparisonAspect.Value,
                    "candidate",
                    value.CandidateType,
                    value.CandidateText));
                selectedSubjects.Add("value-subject");
            }

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
                subjects,
                observations);
            ReviewDecisionGeneration generation =
                ReviewDecisionEngine.Create(catalog);
            for (int index = 0; index < selectedSubjects.Count; index++)
            {
                string subjectId = selectedSubjects[index];
                string selectedObservationId = catalog.Observations.Single(
                    item => item.SubjectId == subjectId &&
                            item.RevisionId == "candidate").ObservationId;
                generation = ReviewDecisionEngine.Apply(
                    generation,
                    ReviewDecisionRequest.ForSubject(
                        generation.ProjectId,
                        generation.BaseGenerationId,
                        generation.MappingGenerationId,
                        generation.GenerationId,
                        subjectId,
                        ReviewDecisionStatus.UseCandidate,
                        ReviewerSnapshot.Empty,
                        When.AddMinutes(index),
                        selectedObservationId));
            }
            return MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(
                    generation,
                    selectedSubjects.AsEnumerable().Reverse()));
        }

        private static Observation Observation(
            string subjectId,
            ComparisonAspect aspect,
            string revisionId,
            ObservationExactType type,
            string text)
        {
            string contentHash = Hash(aspect + ":" + type + ":" + text);
            string? formula = type == ObservationExactType.Formula ? text : null;
            CanonicalScalar? value;
            switch (type)
            {
                case ObservationExactType.Formula:
                    value = null;
                    break;
                case ObservationExactType.Blank:
                    value = CanonicalScalar.Blank;
                    break;
                default:
                    value = new CanonicalScalar(type, text);
                    break;
            }
            return new Observation(
                ComparisonIds.Observation(
                    subjectId,
                    revisionId,
                    type,
                    contentHash),
                subjectId,
                revisionId,
                type,
                contentHash,
                formula,
                value,
                provenance: new ObservationProvenance(
                    revisionId + ".xlsx",
                    "Data",
                    "A1",
                    1,
                    1,
                    "logical-sheet"));
        }

        private static FormulaSpec Formula(
            ObservationExactType masterType,
            string masterText,
            ObservationExactType candidateType,
            string candidateText) =>
            new FormulaSpec(
                masterType,
                masterText,
                candidateType,
                candidateText);

        private static ValueSpec Value(
            ObservationExactType masterType,
            string masterText,
            ObservationExactType candidateType,
            string candidateText) =>
            new ValueSpec(
                masterType,
                masterText,
                candidateType,
                candidateText);

        private static string Hash(string value) =>
            StableId.ForBytes(Encoding.UTF8.GetBytes(value));

        private sealed class FormulaSpec
        {
            internal FormulaSpec(
                ObservationExactType masterType,
                string masterText,
                ObservationExactType candidateType,
                string candidateText)
            {
                MasterType = masterType;
                MasterText = masterText;
                CandidateType = candidateType;
                CandidateText = candidateText;
            }

            internal ObservationExactType MasterType { get; }
            internal string MasterText { get; }
            internal ObservationExactType CandidateType { get; }
            internal string CandidateText { get; }
        }

        private sealed class ValueSpec
        {
            internal ValueSpec(
                ObservationExactType masterType,
                string masterText,
                ObservationExactType candidateType,
                string candidateText)
            {
                MasterType = masterType;
                MasterText = masterText;
                CandidateType = candidateType;
                CandidateText = candidateText;
            }

            internal ObservationExactType MasterType { get; }
            internal string MasterText { get; }
            internal ObservationExactType CandidateType { get; }
            internal string CandidateText { get; }
        }
    }
}
