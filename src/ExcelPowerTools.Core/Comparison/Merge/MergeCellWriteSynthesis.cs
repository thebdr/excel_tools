using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    public static class MergeCellWriteSynthesisContract
    {
        public const int Version = 1;
        public const int MaxWrites = MergePlanContract.MaxOperations;
    }

    public enum MergeCellWriteKind
    {
        SetFormula = 1,
        SetConstant = 2,
        ClearContent = 3,
    }

    /// <summary>
    /// Extra exact master facts required because the selected aspect set alone cannot distinguish two physical
    /// cell states. These are proof obligations, never assumptions.
    /// </summary>
    [Flags]
    public enum MergeCellWriteProofRequirement
    {
        None = 0,
        MasterHasNoFormula = 1,
        MasterConstantValueIsBlank = 2,
    }

    public static class MergeCellWriteSynthesisCodes
    {
        public const string PlanInvalid = "CMP-MERGE-SYNTHESIS-PLAN-INVALID";
        public const string EvidenceInconsistent = "CMP-MERGE-SYNTHESIS-EVIDENCE-INCONSISTENT";
        public const string AspectConflict = "CMP-MERGE-SYNTHESIS-ASPECT-CONFLICT";
        public const string QuotaExceeded = "CMP-MERGE-SYNTHESIS-QUOTA-EXCEEDED";
    }

    public sealed class MergeCellWriteSynthesisException : InvalidOperationException
    {
        internal MergeCellWriteSynthesisException(
            string code,
            string detail,
            string correctiveAction,
            string targetId = "")
            : base(
                MergePlanValidation.RequireIdentity(code, nameof(code)) + ": " +
                (detail ?? "") +
                " Corrective action: " +
                (correctiveAction ?? "Rebuild and review the selected cell changes."))
        {
            Code = code;
            Detail = detail ?? "";
            CorrectiveAction = correctiveAction ?? "";
            TargetId = targetId ?? "";
        }

        public string Code { get; }
        public string Detail { get; }
        public string CorrectiveAction { get; }
        public string TargetId { get; }
    }

    /// <summary>
    /// Complete intended content of one physical cell. It never carries formatting, comments, validation,
    /// cached results, or any other non-content instruction.
    /// </summary>
    public sealed class MergeCellWrite
    {
        internal MergeCellWrite(
            MergePlanTarget target,
            MergeCellWriteKind kind,
            string outputFormula,
            CanonicalScalar outputValue,
            MergeCellWriteProofRequirement proofRequirements,
            MergePlanOperation formulaOperation,
            MergePlanOperation valueOperation)
        {
            Target = target ?? throw new ArgumentNullException(nameof(target));
            if (!Enum.IsDefined(typeof(MergeCellWriteKind), kind))
                throw new ArgumentOutOfRangeException(nameof(kind));
            if ((proofRequirements & ~(
                    MergeCellWriteProofRequirement.MasterHasNoFormula |
                    MergeCellWriteProofRequirement.MasterConstantValueIsBlank)) != 0)
                throw new ArgumentOutOfRangeException(nameof(proofRequirements));

            Kind = kind;
            OutputFormula = outputFormula ?? "";
            OutputValue = outputValue == null
                ? throw new ArgumentNullException(nameof(outputValue))
                : new CanonicalScalar(
                    outputValue.Type,
                    outputValue.CanonicalText,
                    outputValue.DateSystem);
            ProofRequirements = proofRequirements;
            FormulaOperationId = formulaOperation?.OperationId ?? "";
            ValueOperationId = valueOperation?.OperationId ?? "";
            FormulaMasterPrecondition = formulaOperation?.MasterPrecondition;
            ValueMasterPrecondition = valueOperation?.MasterPrecondition;
            FormulaSelectedOutput = formulaOperation?.SelectedOutput;
            ValueSelectedOutput = valueOperation?.SelectedOutput;

            ValidateShape();
            WriteId = StableId.Compute(
                "merge-cell-write-v1",
                Target.TargetId,
                ((int)Kind).ToString(CultureInfo.InvariantCulture),
                ((int)ProofRequirements).ToString(CultureInfo.InvariantCulture),
                FormulaOperationId,
                ValueOperationId,
                FormulaMasterPrecondition?.StateId ?? "",
                ValueMasterPrecondition?.StateId ?? "",
                FormulaSelectedOutput?.StateId ?? "",
                ValueSelectedOutput?.StateId ?? "",
                OutputFormula,
                ((int)OutputValue.Type).ToString(CultureInfo.InvariantCulture),
                OutputValue.CanonicalText,
                ((int)OutputValue.DateSystem).ToString(CultureInfo.InvariantCulture));
        }

        public string WriteId { get; }
        public MergePlanTarget Target { get; }
        public MergeCellWriteKind Kind { get; }
        public string OutputFormula { get; }
        /// <summary>
        /// Exact constant-value aspect after the write. Formula and clear-content writes both have Blank here.
        /// </summary>
        public CanonicalScalar OutputValue { get; }
        public MergeCellWriteProofRequirement ProofRequirements { get; }
        public string FormulaOperationId { get; }
        public string ValueOperationId { get; }
        public MergePlanCellState FormulaMasterPrecondition { get; }
        public MergePlanCellState ValueMasterPrecondition { get; }
        public MergePlanCellState FormulaSelectedOutput { get; }
        public MergePlanCellState ValueSelectedOutput { get; }

        private void ValidateShape()
        {
            switch (Kind)
            {
                case MergeCellWriteKind.SetFormula:
                    if (OutputFormula.Length == 0 ||
                        OutputValue.Type != ObservationExactType.Blank ||
                        FormulaOperationId.Length == 0)
                        throw new ArgumentException(
                            "A formula write requires exact formula text, a blank constant aspect, and a formula operation.");
                    break;

                case MergeCellWriteKind.SetConstant:
                    if (OutputFormula.Length != 0 ||
                        OutputValue.Type == ObservationExactType.Blank ||
                        OutputValue.Type == ObservationExactType.Missing ||
                        OutputValue.Type == ObservationExactType.Formula ||
                        OutputValue.Type == ObservationExactType.Unsupported ||
                        ValueOperationId.Length == 0)
                        throw new ArgumentException(
                            "A constant write requires one supported non-blank value operation and no formula.");
                    break;

                case MergeCellWriteKind.ClearContent:
                    if (OutputFormula.Length != 0 ||
                        OutputValue.Type != ObservationExactType.Blank ||
                        (FormulaOperationId.Length == 0 && ValueOperationId.Length == 0))
                        throw new ArgumentException(
                            "A clear-content write requires a blank result and at least one selected aspect.");
                    break;
            }
        }
    }

    /// <summary>
    /// Deterministic full-cell compilation of one frozen merge plan. It discharges only CellWriteSynthesis; all
    /// proof flags and every other plan validation remain mandatory.
    /// </summary>
    public sealed class MergeCellWriteSynthesis
    {
        internal MergeCellWriteSynthesis(
            MergePlan plan,
            string synthesisHash,
            IEnumerable<MergeCellWrite> writes)
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            PlanHash = plan.PlanHash;
            DecisionSetHash = plan.DecisionSetHash;
            SynthesisHash = synthesisHash;
            SatisfiedValidation = MergePlanRequiredValidation.CellWriteSynthesis;
            RemainingRequiredValidations =
                plan.RequiredValidations & ~MergePlanRequiredValidation.CellWriteSynthesis;
            var frozen = new List<MergeCellWrite>(
                writes ?? throw new ArgumentNullException(nameof(writes)));
            Writes = new ReadOnlyCollection<MergeCellWrite>(frozen);
            WritesRequiringMasterProof = frozen.Count(
                item => item.ProofRequirements != MergeCellWriteProofRequirement.None);
        }

        public int ContractVersion => MergeCellWriteSynthesisContract.Version;
        public string PlanHash { get; }
        public string DecisionSetHash { get; }
        public string SynthesisHash { get; }
        public MergePlanRequiredValidation SatisfiedValidation { get; }
        public MergePlanRequiredValidation RemainingRequiredValidations { get; }
        public IReadOnlyList<MergeCellWrite> Writes { get; }
        public int WritesRequiringMasterProof { get; }
    }

    public static class MergeCellWriteSynthesizer
    {
        public static MergeCellWriteSynthesis Synthesize(
            MergePlan plan,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (plan == null) throw new ArgumentNullException(nameof(plan));
            cancellationToken.ThrowIfCancellationRequested();
            if (plan.ContractVersion != MergePlanContract.Version ||
                plan.Operations == null ||
                plan.Operations.Count == 0 ||
                !StableId.IsSha256(plan.PlanHash) ||
                !StableId.IsSha256(plan.DecisionSetHash) ||
                !plan.RequiredValidations.HasFlag(
                    MergePlanRequiredValidation.CellWriteSynthesis))
            {
                Fail(
                    MergeCellWriteSynthesisCodes.PlanInvalid,
                    "The merge plan is empty, malformed, or does not require this synthesis contract.",
                    "Rebuild the plan with the current merge-plan contract.");
            }
            if (plan.Operations.Count > MergeCellWriteSynthesisContract.MaxWrites)
            {
                Fail(
                    MergeCellWriteSynthesisCodes.QuotaExceeded,
                    "The merge plan exceeds the versioned cell-write synthesis quota.",
                    "Split the selected changes into smaller explicit merge transactions.");
            }

            var groups = new Dictionary<string, List<MergePlanOperation>>(
                StringComparer.OrdinalIgnoreCase);
            foreach (MergePlanOperation operation in plan.Operations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (operation == null || operation.Target == null)
                {
                    Fail(
                        MergeCellWriteSynthesisCodes.PlanInvalid,
                        "The merge plan contains a null operation or target.",
                        "Rebuild the immutable merge plan.");
                }
                string key = PhysicalTargetKey(operation.Target);
                if (!groups.TryGetValue(key, out List<MergePlanOperation> group))
                {
                    group = new List<MergePlanOperation>(2);
                    groups.Add(key, group);
                }
                group.Add(operation);
                if (group.Count > 2)
                {
                    Fail(
                        MergeCellWriteSynthesisCodes.PlanInvalid,
                        "More than two logical aspect operations address one physical cell.",
                        "Correct duplicate alignment or selection evidence.",
                        operation.Target.TargetId);
                }
            }

            var writes = new List<MergeCellWrite>(groups.Count);
            foreach (List<MergePlanOperation> group in groups.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                writes.Add(SynthesizeCell(group));
            }
            writes.Sort(CompareWrites);

            var components = new List<string>
            {
                MergeCellWriteSynthesisContract.Version.ToString(
                    CultureInfo.InvariantCulture),
                plan.PlanHash,
                plan.DecisionSetHash,
                ((int)(plan.RequiredValidations &
                       ~MergePlanRequiredValidation.CellWriteSynthesis)).ToString(
                    CultureInfo.InvariantCulture),
                writes.Count.ToString(CultureInfo.InvariantCulture),
            };
            components.AddRange(writes.Select(item => item.WriteId));
            string synthesisHash = StableId.Compute(
                "merge-cell-write-synthesis-v1",
                components.ToArray());
            return new MergeCellWriteSynthesis(plan, synthesisHash, writes);
        }

        private static MergeCellWrite SynthesizeCell(
            IReadOnlyList<MergePlanOperation> operations)
        {
            MergePlanOperation formula = null;
            MergePlanOperation value = null;
            foreach (MergePlanOperation operation in operations)
            {
                switch (operation.Aspect)
                {
                    case ComparisonAspect.Formula:
                        if (formula != null)
                            InvalidPlan(operation, "The cell repeats a Formula operation.");
                        formula = operation;
                        ValidateFormulaOperation(operation);
                        break;

                    case ComparisonAspect.Value:
                        if (value != null)
                            InvalidPlan(operation, "The cell repeats a Value operation.");
                        value = operation;
                        ValidateValueOperation(operation);
                        break;

                    default:
                        InvalidPlan(
                            operation,
                            "Only Formula and Value operations may enter cell synthesis.");
                        break;
                }
            }
            if (formula == null && value == null)
            {
                Fail(
                    MergeCellWriteSynthesisCodes.PlanInvalid,
                    "A physical target group contains no formula/value operation.",
                    "Rebuild the merge plan.");
            }
            MergePlanTarget target = formula?.Target ?? value.Target;
            if (formula != null && value != null)
            {
                EnsureSameTarget(formula.Target, value.Target);
                ValidateMasterCombination(formula, value);
                ValidateSelectedCombination(formula, value);
            }

            if (formula != null &&
                formula.SelectedOutput.ExactType == ObservationExactType.Formula)
            {
                MergeCellWriteProofRequirement proof =
                    value == null &&
                    formula.MasterPrecondition.ExactType == ObservationExactType.Blank
                        ? MergeCellWriteProofRequirement.MasterConstantValueIsBlank
                        : MergeCellWriteProofRequirement.None;
                return new MergeCellWrite(
                    target,
                    MergeCellWriteKind.SetFormula,
                    formula.SelectedOutput.Formula,
                    CanonicalScalar.Blank,
                    proof,
                    formula,
                    value);
            }

            if (value != null)
            {
                MergeCellWriteKind kind =
                    value.SelectedOutput.ExactType == ObservationExactType.Blank
                        ? MergeCellWriteKind.ClearContent
                        : MergeCellWriteKind.SetConstant;
                MergeCellWriteProofRequirement proof =
                    formula == null &&
                    value.MasterPrecondition.ExactType == ObservationExactType.Blank
                        ? MergeCellWriteProofRequirement.MasterHasNoFormula
                        : MergeCellWriteProofRequirement.None;
                return new MergeCellWrite(
                    target,
                    kind,
                    "",
                    value.SelectedOutput.Value,
                    proof,
                    formula,
                    value);
            }

            // The selected Formula state is Blank and the analyzed master Formula state is necessarily Formula,
            // because the plan builder rejected a no-op. Preserving the unselected master constant aspect means
            // preserving Blank, so the exact hybrid result is a clear-content operation.
            return new MergeCellWrite(
                target,
                MergeCellWriteKind.ClearContent,
                "",
                CanonicalScalar.Blank,
                MergeCellWriteProofRequirement.None,
                formula,
                null);
        }

        private static void ValidateFormulaOperation(MergePlanOperation operation)
        {
            if (!IsFormulaState(operation.MasterPrecondition) ||
                !IsFormulaState(operation.SelectedOutput) ||
                operation.MasterPrecondition.PayloadEquals(operation.SelectedOutput))
            {
                InvalidPlan(
                    operation,
                    "The Formula operation has an unsupported or unchanged exact state.");
            }
        }

        private static void ValidateValueOperation(MergePlanOperation operation)
        {
            if (!IsValueState(operation.MasterPrecondition) ||
                !IsValueState(operation.SelectedOutput) ||
                operation.MasterPrecondition.PayloadEquals(operation.SelectedOutput))
            {
                InvalidPlan(
                    operation,
                    "The Value operation has an unsupported or unchanged exact state.");
            }
        }

        private static void ValidateMasterCombination(
            MergePlanOperation formula,
            MergePlanOperation value)
        {
            if (formula.MasterPrecondition.ExactType == ObservationExactType.Formula &&
                value.MasterPrecondition.ExactType != ObservationExactType.Blank)
            {
                Fail(
                    MergeCellWriteSynthesisCodes.EvidenceInconsistent,
                    "The master evidence claims both a formula and a non-blank constant at " +
                    formula.Target.SheetName + "!" + formula.Target.Address + ".",
                    "Rebuild canonical cell evidence from the exact master bytes.",
                    formula.Target.TargetId);
            }
        }

        private static void ValidateSelectedCombination(
            MergePlanOperation formula,
            MergePlanOperation value)
        {
            if (formula.SelectedOutput.ExactType == ObservationExactType.Formula &&
                value.SelectedOutput.ExactType != ObservationExactType.Blank)
            {
                Fail(
                    MergeCellWriteSynthesisCodes.AspectConflict,
                    "The selected hybrid would require both a formula and a non-blank constant at " +
                    formula.Target.SheetName + "!" + formula.Target.Address + ".",
                    "Select a blank Value aspect with the formula, or select a no-formula state with the constant.",
                    formula.Target.TargetId);
            }
        }

        private static bool IsFormulaState(MergePlanCellState state) =>
            state != null &&
            (state.ExactType == ObservationExactType.Formula ||
             state.ExactType == ObservationExactType.Blank);

        private static bool IsValueState(MergePlanCellState state) =>
            state != null &&
            state.Value != null &&
            state.ExactType != ObservationExactType.Formula &&
            state.ExactType != ObservationExactType.Missing &&
            state.ExactType != ObservationExactType.Unsupported;

        private static void EnsureSameTarget(
            MergePlanTarget left,
            MergePlanTarget right)
        {
            if (!string.Equals(left.TargetId, right.TargetId, StringComparison.Ordinal) ||
                !string.Equals(
                    left.LogicalSheetId,
                    right.LogicalSheetId,
                    StringComparison.Ordinal) ||
                !string.Equals(left.SheetName, right.SheetName, StringComparison.Ordinal) ||
                !string.Equals(left.Address, right.Address, StringComparison.Ordinal) ||
                left.PhysicalRow != right.PhysicalRow ||
                left.PhysicalColumn != right.PhysicalColumn)
            {
                Fail(
                    MergeCellWriteSynthesisCodes.EvidenceInconsistent,
                    "Formula and Value operations for one physical cell do not retain the same exact target.",
                    "Correct the mapping and rebuild the merge plan.",
                    left?.TargetId ?? right?.TargetId ?? "");
            }
        }

        private static string PhysicalTargetKey(MergePlanTarget target) =>
            target.SheetName + "\u001F" +
            target.PhysicalRow.ToString(CultureInfo.InvariantCulture) + "\u001F" +
            target.PhysicalColumn.ToString(CultureInfo.InvariantCulture);

        private static int CompareWrites(MergeCellWrite left, MergeCellWrite right)
        {
            int value = StringComparer.Ordinal.Compare(
                left.Target.LogicalSheetId,
                right.Target.LogicalSheetId);
            if (value != 0) return value;
            value = left.Target.PhysicalRow.CompareTo(right.Target.PhysicalRow);
            if (value != 0) return value;
            value = left.Target.PhysicalColumn.CompareTo(right.Target.PhysicalColumn);
            return value != 0
                ? value
                : StringComparer.Ordinal.Compare(
                    left.Target.TargetId,
                    right.Target.TargetId);
        }

        private static void InvalidPlan(
            MergePlanOperation operation,
            string detail)
        {
            Fail(
                MergeCellWriteSynthesisCodes.PlanInvalid,
                detail,
                "Rebuild the immutable merge plan from exact current decisions.",
                operation?.Target?.TargetId ?? "");
        }

        private static void Fail(
            string code,
            string detail,
            string correctiveAction,
            string targetId = "")
        {
            throw new MergeCellWriteSynthesisException(
                code,
                detail,
                correctiveAction,
                targetId);
        }
    }
}
