using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    /// <summary>Versioned, bounded contract for a frozen formula/value merge plan.</summary>
    public static class MergePlanContract
    {
        public const int Version = 1;
        public const int MaxOperations = 100000;
        public const int MaxIdentityCharacters = 4096;
        public const int MaxCellTextCharacters = 32767;
        public const int MaxSheetNameCharacters = 31;
        public const int MaxRows = 1048576;
        public const int MaxColumns = 16384;
    }

    /// <summary>
    /// Checks that remain mandatory after the deterministic plan is frozen. A merge plan is selection authority,
    /// not write authority: a writer must consume a matching validation receipt rather than this flag set alone.
    /// </summary>
    [Flags]
    public enum MergePlanRequiredValidation
    {
        None = 0,
        SourceMasterByteIdentity = 1,
        ProjectAndArtifactFindings = 2,
        MasterTargetPreconditions = 4,
        CellWriteSynthesis = 8,
        WriterSafetyEnvelope = 16,
        FullOutputSemanticDiff = 32,
        AtomicPublication = 64,
        FormulaDependencies = 128,
        FormulaGroups = 256,
        FrozenDecisionAndMappingIdentity = 512,
    }

    public static class MergePlanBuildCodes
    {
        public const string EvidenceRequired = "CMP-MERGE-PLAN-EVIDENCE-REQUIRED";
        public const string ContextMismatch = "CMP-MERGE-PLAN-CONTEXT-MISMATCH";
        public const string SelectionInvalid = "CMP-MERGE-PLAN-SELECTION-INVALID";
        public const string SubjectNotMergeable = "CMP-MERGE-PLAN-SUBJECT-NOT-MERGEABLE";
        public const string EvidenceInconsistent = "CMP-MERGE-PLAN-EVIDENCE-INCONSISTENT";
        public const string NoActualChange = "CMP-MERGE-PLAN-NO-ACTUAL-CHANGE";
        public const string StateUnsupported = "CMP-MERGE-PLAN-STATE-UNSUPPORTED";
        public const string SourceInvalid = "CMP-MERGE-PLAN-SOURCE-INVALID";
        public const string TargetInvalid = "CMP-MERGE-PLAN-TARGET-INVALID";
        public const string TargetAmbiguous = "CMP-MERGE-PLAN-TARGET-AMBIGUOUS";
        public const string QuotaExceeded = "CMP-MERGE-PLAN-QUOTA-EXCEEDED";
    }

    /// <summary>
    /// Exact physical occurrence of the selected observation. It is path-free: SourceLabel is deliberately not
    /// retained.
    /// </summary>
    public sealed class MergePlanSourceOccurrence
    {
        internal MergePlanSourceOccurrence(ObservationProvenance provenance)
        {
            if (provenance == null) throw new ArgumentNullException(nameof(provenance));
            LogicalSheetId = MergePlanValidation.RequireIdentity(
                provenance.LogicalSheetId,
                nameof(provenance.LogicalSheetId));
            SheetName = MergePlanValidation.RequireSheetName(
                provenance.SheetName,
                nameof(provenance.SheetName));
            if (provenance.PhysicalRow < 1 ||
                provenance.PhysicalRow > MergePlanContract.MaxRows)
                throw new ArgumentOutOfRangeException(
                    nameof(provenance.PhysicalRow),
                    "The selected source row is outside the Excel worksheet grid.");
            if (provenance.PhysicalColumn < 1 ||
                provenance.PhysicalColumn > MergePlanContract.MaxColumns)
                throw new ArgumentOutOfRangeException(
                    nameof(provenance.PhysicalColumn),
                    "The selected source column is outside the Excel worksheet grid.");

            PhysicalRow = provenance.PhysicalRow;
            PhysicalColumn = provenance.PhysicalColumn;
            Address = MergePlanValidation.RequireCanonicalAddress(
                provenance.Address,
                PhysicalRow,
                PhysicalColumn,
                nameof(provenance.Address));
            OccurrenceId = StableId.Compute(
                "merge-plan-selected-source-occurrence-v1",
                LogicalSheetId,
                SheetName,
                Address,
                PhysicalRow.ToString(CultureInfo.InvariantCulture),
                PhysicalColumn.ToString(CultureInfo.InvariantCulture));
        }

        public string OccurrenceId { get; }
        public string LogicalSheetId { get; }
        public string SheetName { get; }
        public string Address { get; }
        public int PhysicalRow { get; }
        public int PhysicalColumn { get; }
    }

    /// <summary>
    /// Stable, structured failure suitable for verbose diagnostic projection. It deliberately carries no path,
    /// reviewer, timestamp, or source-label field.
    /// </summary>
    public sealed class MergePlanBuildException : InvalidOperationException
    {
        internal MergePlanBuildException(
            string code,
            string detail,
            string correctiveAction,
            string subjectId = "")
            : base(
                MergePlanValidation.RequireIdentity(code, nameof(code)) + ": " +
                (detail ?? "") +
                " Corrective action: " +
                (correctiveAction ?? "Rebuild the comparison evidence and review the selected change."))
        {
            Code = code;
            Detail = detail ?? "";
            CorrectiveAction = correctiveAction ?? "";
            SubjectId = subjectId ?? "";
        }

        public string Code { get; }
        public string Detail { get; }
        public string CorrectiveAction { get; }
        public string SubjectId { get; }
    }

    /// <summary>One explicitly selected effective reviewer decision.</summary>
    public sealed class MergePlanSelection
    {
        public MergePlanSelection(string subjectId, string effectiveDecisionEventId)
        {
            SubjectId = MergePlanValidation.RequireIdentity(subjectId, nameof(subjectId));
            EffectiveDecisionEventId = MergePlanValidation.RequireIdentity(
                effectiveDecisionEventId,
                nameof(effectiveDecisionEventId));
        }

        public string SubjectId { get; }
        public string EffectiveDecisionEventId { get; }
    }

    /// <summary>
    /// Optimistic-concurrency envelope captured before plan construction. The explicit event IDs make a selected
    /// subset stale as soon as any of its effective reviewer decisions changes.
    /// </summary>
    public sealed class MergePlanRequest
    {
        public MergePlanRequest(
            string projectId,
            string evidenceGenerationId,
            string mappingGenerationId,
            string catalogId,
            string expectedDecisionGenerationId,
            string masterRevisionId,
            IEnumerable<MergePlanSelection> selections)
        {
            ProjectId = MergePlanValidation.RequireIdentity(projectId, nameof(projectId));
            EvidenceGenerationId = MergePlanValidation.RequireIdentity(
                evidenceGenerationId,
                nameof(evidenceGenerationId));
            MappingGenerationId = MergePlanValidation.RequireIdentity(
                mappingGenerationId,
                nameof(mappingGenerationId));
            CatalogId = MergePlanValidation.RequireIdentity(catalogId, nameof(catalogId));
            ExpectedDecisionGenerationId = MergePlanValidation.RequireIdentity(
                expectedDecisionGenerationId,
                nameof(expectedDecisionGenerationId));
            MasterRevisionId = MergePlanValidation.RequireIdentity(
                masterRevisionId,
                nameof(masterRevisionId));
            Selections = new ReadOnlyCollection<MergePlanSelection>(
                FreezeSelections(selections));
        }

        public string ProjectId { get; }
        public string EvidenceGenerationId { get; }
        public string MappingGenerationId { get; }
        public string CatalogId { get; }
        public string ExpectedDecisionGenerationId { get; }
        public string MasterRevisionId { get; }
        public IReadOnlyList<MergePlanSelection> Selections { get; }

        public static MergePlanRequest Capture(
            ReviewDecisionGeneration generation,
            IEnumerable<string> selectedSubjectIds)
        {
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            if (selectedSubjectIds == null)
                throw new ArgumentNullException(nameof(selectedSubjectIds));

            var selections = new List<MergePlanSelection>();
            foreach (string subjectId in selectedSubjectIds)
            {
                if (selections.Count == MergePlanContract.MaxOperations)
                {
                    throw new MergePlanBuildException(
                        MergePlanBuildCodes.QuotaExceeded,
                        "The selected subject count exceeds the versioned merge-plan operation quota.",
                        "Split the reviewed changes into smaller explicit merge transactions.");
                }
                string required = MergePlanValidation.RequireIdentity(
                    subjectId,
                    nameof(selectedSubjectIds));
                ReviewDecisionState state = generation.GetSubjectState(required);
                if (state.EffectiveEventId.Length == 0)
                {
                    throw new MergePlanBuildException(
                        MergePlanBuildCodes.SelectionInvalid,
                        "Selected subject " + required +
                        " has no effective reviewed decision event.",
                        "Review the subject, explicitly choose a candidate revision, and capture the selection again.",
                        required);
                }
                selections.Add(new MergePlanSelection(required, state.EffectiveEventId));
            }

            return new MergePlanRequest(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.CatalogId,
                generation.GenerationId,
                generation.Catalog.MasterRevisionId,
                selections);
        }

        private static List<MergePlanSelection> FreezeSelections(
            IEnumerable<MergePlanSelection> selections)
        {
            if (selections == null) throw new ArgumentNullException(nameof(selections));
            var result = new List<MergePlanSelection>();
            var subjects = new HashSet<string>(StringComparer.Ordinal);
            var events = new HashSet<string>(StringComparer.Ordinal);
            foreach (MergePlanSelection selection in selections)
            {
                if (selection == null)
                    throw new ArgumentException(
                        "A merge-plan selection cannot be null.",
                        nameof(selections));
                if (result.Count == MergePlanContract.MaxOperations)
                    throw new ArgumentException(
                        "The merge-plan selection count exceeds the versioned operation quota.",
                        nameof(selections));
                if (!subjects.Add(selection.SubjectId))
                    throw new ArgumentException(
                        "A merge-plan request cannot repeat a subject.",
                        nameof(selections));
                if (!events.Add(selection.EffectiveDecisionEventId))
                    throw new ArgumentException(
                        "A merge-plan request cannot reuse one decision event for multiple subjects.",
                        nameof(selections));
                result.Add(selection);
            }
            if (result.Count == 0)
                throw new ArgumentException(
                    "A merge plan requires at least one explicit selected change.",
                    nameof(selections));
            result.Sort((left, right) =>
                StringComparer.Ordinal.Compare(left.SubjectId, right.SubjectId));
            return result;
        }
    }

    /// <summary>
    /// Exact typed aspect state copied from immutable evidence. Display text and formula cached results are
    /// intentionally excluded: neither is formula/value merge authority.
    /// </summary>
    public sealed class MergePlanCellState
    {
        internal MergePlanCellState(Observation observation)
        {
            if (observation == null) throw new ArgumentNullException(nameof(observation));
            if (!StableId.IsSha256(observation.ContentHash))
                throw new ArgumentException(
                    "A planned state requires a SHA-256 content identity.",
                    nameof(observation));
            if (observation.Formula.Length > MergePlanContract.MaxCellTextCharacters)
                throw new ArgumentException(
                    "Formula evidence exceeds the merge-plan text quota.",
                    nameof(observation));
            if (observation.Value != null &&
                observation.Value.CanonicalText.Length > MergePlanContract.MaxCellTextCharacters)
                throw new ArgumentException(
                    "Scalar evidence exceeds the merge-plan text quota.",
                    nameof(observation));

            ExactType = observation.ExactType;
            ContentHash = observation.ContentHash;
            Formula = observation.Formula;
            Value = observation.Value == null
                ? null
                : new CanonicalScalar(
                    observation.Value.Type,
                    observation.Value.CanonicalText,
                    observation.Value.DateSystem);

            StateId = StableId.Compute(
                "merge-plan-cell-state-v1",
                ((int)ExactType).ToString(CultureInfo.InvariantCulture),
                ContentHash,
                Formula,
                Value == null ? "0" : "1",
                Value == null
                    ? ""
                    : ((int)Value.Type).ToString(CultureInfo.InvariantCulture),
                Value?.CanonicalText ?? "",
                Value == null
                    ? ""
                    : ((int)Value.DateSystem).ToString(CultureInfo.InvariantCulture));
        }

        public string StateId { get; }
        public ObservationExactType ExactType { get; }
        public string ContentHash { get; }
        public string Formula { get; }
        public CanonicalScalar Value { get; }

        internal bool PayloadEquals(MergePlanCellState other) =>
            other != null &&
            ExactType == other.ExactType &&
            string.Equals(Formula, other.Formula, StringComparison.Ordinal) &&
            ScalarsEqual(Value, other.Value);

        private static bool ScalarsEqual(CanonicalScalar left, CanonicalScalar right)
        {
            if (ReferenceEquals(left, right)) return true;
            return left != null &&
                   right != null &&
                   left.Type == right.Type &&
                   left.DateSystem == right.DateSystem &&
                   string.Equals(
                       left.CanonicalText,
                       right.CanonicalText,
                       StringComparison.Ordinal);
        }
    }

    /// <summary>Exact physical occurrence in the analyzed master; SourceLabel is deliberately not retained.</summary>
    public sealed class MergePlanTarget
    {
        internal MergePlanTarget(ObservationProvenance provenance)
        {
            if (provenance == null) throw new ArgumentNullException(nameof(provenance));
            LogicalSheetId = MergePlanValidation.RequireIdentity(
                provenance.LogicalSheetId,
                nameof(provenance.LogicalSheetId));
            SheetName = MergePlanValidation.RequireSheetName(
                provenance.SheetName,
                nameof(provenance.SheetName));
            if (provenance.PhysicalRow < 1 ||
                provenance.PhysicalRow > MergePlanContract.MaxRows)
                throw new ArgumentOutOfRangeException(
                    nameof(provenance.PhysicalRow),
                    "The master target row is outside the Excel worksheet grid.");
            if (provenance.PhysicalColumn < 1 ||
                provenance.PhysicalColumn > MergePlanContract.MaxColumns)
                throw new ArgumentOutOfRangeException(
                    nameof(provenance.PhysicalColumn),
                    "The master target column is outside the Excel worksheet grid.");

            PhysicalRow = provenance.PhysicalRow;
            PhysicalColumn = provenance.PhysicalColumn;
            Address = MergePlanValidation.RequireCanonicalAddress(
                provenance.Address,
                PhysicalRow,
                PhysicalColumn,
                nameof(provenance.Address));
            TargetId = StableId.Compute(
                "merge-plan-target-v1",
                LogicalSheetId,
                SheetName,
                Address,
                PhysicalRow.ToString(CultureInfo.InvariantCulture),
                PhysicalColumn.ToString(CultureInfo.InvariantCulture));
        }

        public string TargetId { get; }
        public string LogicalSheetId { get; }
        public string SheetName { get; }
        public string Address { get; }
        public int PhysicalRow { get; }
        public int PhysicalColumn { get; }
    }

    /// <summary>
    /// One selected logical aspect change. It is not a sequential cell-write instruction: Formula and Value
    /// changes sharing a target must first pass the required cell-write synthesis step.
    /// </summary>
    public sealed class MergePlanOperation
    {
        internal MergePlanOperation(
            ReviewDecisionSubject subject,
            DecisionEvent decision,
            Observation selectedObservation,
            string selectedRevisionByteContentId,
            Observation masterObservation,
            MergePlanSourceOccurrence selectedSource,
            MergePlanTarget target,
            MergePlanCellState masterPrecondition,
            MergePlanCellState selectedOutput)
        {
            if (subject == null) throw new ArgumentNullException(nameof(subject));
            if (decision == null) throw new ArgumentNullException(nameof(decision));
            if (selectedObservation == null)
                throw new ArgumentNullException(nameof(selectedObservation));
            if (masterObservation == null)
                throw new ArgumentNullException(nameof(masterObservation));
            SubjectId = MergePlanValidation.RequireIdentity(
                subject.SubjectId,
                nameof(subject.SubjectId));
            Aspect = subject.Aspect;
            DecisionEventId = MergePlanValidation.RequireIdentity(
                decision.EventId,
                nameof(decision.EventId));
            SelectedObservationId = MergePlanValidation.RequireIdentity(
                selectedObservation.ObservationId,
                nameof(selectedObservation.ObservationId));
            SelectedRevisionId = MergePlanValidation.RequireIdentity(
                selectedObservation.RevisionId,
                nameof(selectedObservation.RevisionId));
            if (!StableId.IsSha256(selectedRevisionByteContentId))
                throw new ArgumentException(
                    "A planned selected revision requires a SHA-256 byte identity.",
                    nameof(selectedRevisionByteContentId));
            SelectedRevisionByteContentId = selectedRevisionByteContentId;
            MasterObservationId = MergePlanValidation.RequireIdentity(
                masterObservation.ObservationId,
                nameof(masterObservation.ObservationId));
            SelectedSource = selectedSource ??
                             throw new ArgumentNullException(nameof(selectedSource));
            Target = target ?? throw new ArgumentNullException(nameof(target));
            MasterPrecondition = masterPrecondition ??
                                 throw new ArgumentNullException(nameof(masterPrecondition));
            SelectedOutput = selectedOutput ??
                             throw new ArgumentNullException(nameof(selectedOutput));
            OperationId = StableId.Compute(
                "merge-plan-operation-v1",
                SubjectId,
                ((int)Aspect).ToString(CultureInfo.InvariantCulture),
                DecisionEventId,
                SelectedObservationId,
                SelectedRevisionId,
                SelectedRevisionByteContentId,
                MasterObservationId,
                SelectedSource.OccurrenceId,
                Target.TargetId,
                MasterPrecondition.StateId,
                SelectedOutput.StateId);
        }

        public string OperationId { get; }
        public string SubjectId { get; }
        public ComparisonAspect Aspect { get; }
        public string DecisionEventId { get; }
        public string SelectedObservationId { get; }
        public string SelectedRevisionId { get; }
        public string SelectedRevisionByteContentId { get; }
        public string MasterObservationId { get; }
        public MergePlanSourceOccurrence SelectedSource { get; }
        public MergePlanTarget Target { get; }
        public MergePlanCellState MasterPrecondition { get; }
        public MergePlanCellState SelectedOutput { get; }
    }

    /// <summary>Counts required by the future confirmation UI, frozen with the exact decision generation.</summary>
    public sealed class MergePlanReviewSummary
    {
        internal MergePlanReviewSummary(
            int totalSubjects,
            int unreviewed,
            int useMaster,
            int useCandidate,
            int equivalent,
            int needsInvestigation,
            int reviewedNoStructuralMerge,
            int selectedCandidate)
        {
            TotalSubjects = totalSubjects;
            Unreviewed = unreviewed;
            UseMaster = useMaster;
            UseCandidate = useCandidate;
            Equivalent = equivalent;
            NeedsInvestigation = needsInvestigation;
            ReviewedNoStructuralMerge = reviewedNoStructuralMerge;
            SelectedCandidate = selectedCandidate;
        }

        public int TotalSubjects { get; }
        public int Unreviewed { get; }
        public int UseMaster { get; }
        public int UseCandidate { get; }
        public int Equivalent { get; }
        public int NeedsInvestigation { get; }
        public int ReviewedNoStructuralMerge { get; }
        public int SelectedCandidate { get; }
        public int UnselectedCandidate => checked(UseCandidate - SelectedCandidate);

        internal void AddHashComponents(ICollection<string> target)
        {
            target.Add(TotalSubjects.ToString(CultureInfo.InvariantCulture));
            target.Add(Unreviewed.ToString(CultureInfo.InvariantCulture));
            target.Add(UseMaster.ToString(CultureInfo.InvariantCulture));
            target.Add(UseCandidate.ToString(CultureInfo.InvariantCulture));
            target.Add(Equivalent.ToString(CultureInfo.InvariantCulture));
            target.Add(NeedsInvestigation.ToString(CultureInfo.InvariantCulture));
            target.Add(ReviewedNoStructuralMerge.ToString(CultureInfo.InvariantCulture));
            target.Add(SelectedCandidate.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>
    /// Immutable deterministic selection plan. The required validation flags are intentionally non-empty; this
    /// object cannot authorize a workbook mutation until a later component proves every named check against the
    /// same PlanHash.
    /// </summary>
    public sealed class MergePlan
    {
        internal MergePlan(
            string projectId,
            string evidenceGenerationId,
            string mappingGenerationId,
            string catalogId,
            string decisionGenerationId,
            string masterRevisionId,
            string masterByteContentId,
            string decisionSetHash,
            string planHash,
            MergePlanRequiredValidation requiredValidations,
            MergePlanReviewSummary reviewSummary,
            IEnumerable<MergePlanOperation> operations)
        {
            ProjectId = projectId;
            EvidenceGenerationId = evidenceGenerationId;
            MappingGenerationId = mappingGenerationId;
            CatalogId = catalogId;
            DecisionGenerationId = decisionGenerationId;
            MasterRevisionId = masterRevisionId;
            MasterByteContentId = masterByteContentId;
            DecisionSetHash = decisionSetHash;
            PlanHash = planHash;
            RequiredValidations = requiredValidations;
            ReviewSummary = reviewSummary ?? throw new ArgumentNullException(nameof(reviewSummary));
            var frozen = new List<MergePlanOperation>(
                operations ?? throw new ArgumentNullException(nameof(operations)));
            Operations = new ReadOnlyCollection<MergePlanOperation>(frozen);
            SelectedDecisionIds = new ReadOnlyCollection<string>(
                frozen.Select(item => item.DecisionEventId).ToList());
        }

        public int ContractVersion => MergePlanContract.Version;
        public string ProjectId { get; }
        public string EvidenceGenerationId { get; }
        public string MappingGenerationId { get; }
        public string CatalogId { get; }
        public string DecisionGenerationId { get; }
        public string MasterRevisionId { get; }
        public string MasterByteContentId { get; }
        public string DecisionSetHash { get; }
        public string PlanHash { get; }
        public MergePlanRequiredValidation RequiredValidations { get; }
        public MergePlanReviewSummary ReviewSummary { get; }
        public IReadOnlyList<MergePlanOperation> Operations { get; }
        public IReadOnlyList<string> SelectedDecisionIds { get; }
    }

    internal static class MergePlanValidation
    {
        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

        internal static string RequireIdentity(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException(
                    "A non-blank merge-plan identity is required.",
                    parameterName);
            if (value.Length > MergePlanContract.MaxIdentityCharacters)
                throw new ArgumentException(
                    "A merge-plan identity exceeds the versioned character quota.",
                    parameterName);
            EnsureWellFormedUnicode(value, parameterName);
            return value;
        }

        internal static string RequireSheetName(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A master worksheet name is required.", parameterName);
            if (value.Length > MergePlanContract.MaxSheetNameCharacters)
                throw new ArgumentException(
                    "The master worksheet name exceeds the Excel limit.",
                    parameterName);
            if (value[0] == '\'' || value[value.Length - 1] == '\'')
                throw new ArgumentException(
                    "A worksheet name cannot begin or end with an apostrophe.",
                    parameterName);
            for (int index = 0; index < value.Length; index++)
            {
                char current = value[index];
                if (char.IsControl(current) ||
                    current == '[' ||
                    current == ']' ||
                    current == ':' ||
                    current == '*' ||
                    current == '?' ||
                    current == '/' ||
                    current == '\\')
                    throw new ArgumentException(
                        "The master worksheet name contains an invalid character.",
                        parameterName);
            }
            EnsureWellFormedUnicode(value, parameterName);
            return value;
        }

        internal static string RequireCanonicalAddress(
            string value,
            int row,
            int column,
            string parameterName)
        {
            string expected = ToA1(row, column);
            if (!string.Equals(value, expected, StringComparison.Ordinal))
                throw new ArgumentException(
                    "The master target address is missing, non-canonical, or disagrees with its coordinates.",
                    parameterName);
            return value;
        }

        private static string ToA1(int row, int column)
        {
            var letters = new char[3];
            int cursor = letters.Length;
            int remaining = column;
            while (remaining > 0)
            {
                remaining--;
                letters[--cursor] = (char)('A' + (remaining % 26));
                remaining /= 26;
            }
            return new string(letters, cursor, letters.Length - cursor) +
                   row.ToString(CultureInfo.InvariantCulture);
        }

        private static void EnsureWellFormedUnicode(string value, string parameterName)
        {
            try
            {
                StrictUtf8.GetByteCount(value);
            }
            catch (EncoderFallbackException exception)
            {
                throw new ArgumentException(
                    "Merge-plan text must be well-formed UTF-16.",
                    parameterName,
                    exception);
            }
        }
    }
}
