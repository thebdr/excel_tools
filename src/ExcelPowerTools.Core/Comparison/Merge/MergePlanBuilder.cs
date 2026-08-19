using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Merge
{
    /// <summary>
    /// Builds a deterministic formula/value selection plan from one exact evidence-bound decision generation.
    /// It performs no I/O and grants no write authority.
    /// </summary>
    public static class MergePlanBuilder
    {
        private const MergePlanRequiredValidation BaseRequiredValidations =
            MergePlanRequiredValidation.SourceMasterByteIdentity |
            MergePlanRequiredValidation.ProjectAndArtifactFindings |
            MergePlanRequiredValidation.MasterTargetPreconditions |
            MergePlanRequiredValidation.CellWriteSynthesis |
            MergePlanRequiredValidation.WriterSafetyEnvelope |
            MergePlanRequiredValidation.FullOutputSemanticDiff |
            MergePlanRequiredValidation.AtomicPublication |
            MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity;

        public static MergePlan Build(
            ReviewDecisionGeneration generation,
            MergePlanRequest request,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (generation == null) throw new ArgumentNullException(nameof(generation));
            if (request == null) throw new ArgumentNullException(nameof(request));
            cancellationToken.ThrowIfCancellationRequested();

            ReviewDecisionCatalog catalog = generation.Catalog;
            ReviewEvidenceGeneration evidence = catalog.EvidenceGeneration;
            if (evidence == null)
            {
                Fail(
                    MergePlanBuildCodes.EvidenceRequired,
                    "The decision catalog was built from opaque or legacy generation labels rather than exact revision bytes.",
                    "Reopen or rebuild the comparison so every revision occurrence has a verified byte identity.");
            }
            ValidateContext(generation, request);

            ReviewEvidenceRevisionIdentity masterEvidence = evidence.Revisions.SingleOrDefault(
                item => string.Equals(
                    item.RevisionId,
                    catalog.MasterRevisionId,
                    StringComparison.Ordinal));
            if (masterEvidence == null)
            {
                Fail(
                    MergePlanBuildCodes.EvidenceInconsistent,
                    "The evidence generation does not contain the catalog's master revision occurrence.",
                    "Rebuild the evidence generation and confirm the intended master.");
            }
            IReadOnlyDictionary<string, ReviewEvidenceRevisionIdentity> evidenceByRevision =
                new ReadOnlyDictionary<string, ReviewEvidenceRevisionIdentity>(
                    evidence.Revisions.ToDictionary(
                        item => item.RevisionId,
                        StringComparer.Ordinal));

            var operations = new List<MergePlanOperation>(request.Selections.Count);
            var physicalTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var logicalToPhysicalSheet = new Dictionary<string, string>(StringComparer.Ordinal);
            var physicalToLogicalSheet = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool hasFormulaOperation = false;

            foreach (MergePlanSelection selection in request.Selections)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!catalog.TryGetSubject(
                        selection.SubjectId,
                        out ReviewDecisionSubject subject))
                {
                    Fail(
                        MergePlanBuildCodes.SelectionInvalid,
                        "Selected subject " + selection.SubjectId +
                        " is outside the exact review catalog.",
                        "Refresh the review workspace and select the change again.",
                        selection.SubjectId);
                }
                if (!subject.CanMergeSelection ||
                    (subject.Aspect != ComparisonAspect.Formula &&
                     subject.Aspect != ComparisonAspect.Value))
                {
                    Fail(
                        MergePlanBuildCodes.SubjectNotMergeable,
                        "Subject " + selection.SubjectId +
                        " is not explicitly authorized for formula/value merge.",
                        "Keep the item as review evidence or choose a mergeable formula/value subject.",
                        selection.SubjectId);
                }

                ReviewDecisionState state = generation.GetSubjectState(selection.SubjectId);
                if (state.Status != ReviewDecisionStatus.UseCandidate ||
                    !string.Equals(
                        state.EffectiveEventId,
                        selection.EffectiveDecisionEventId,
                        StringComparison.Ordinal) ||
                    state.SelectedObservationId.Length == 0)
                {
                    Fail(
                        MergePlanBuildCodes.SelectionInvalid,
                        "Subject " + selection.SubjectId +
                        " is not an exact current UseCandidate decision.",
                        "Refresh the decision generation, explicitly choose the candidate, and rebuild the plan.",
                        selection.SubjectId);
                }
                if (!generation.TryGetActiveDecision(
                        selection.SubjectId,
                        out DecisionEvent decision) ||
                    !string.Equals(
                        decision.EventId,
                        selection.EffectiveDecisionEventId,
                        StringComparison.Ordinal))
                {
                    Fail(
                        MergePlanBuildCodes.SelectionInvalid,
                        "The effective event for subject " + selection.SubjectId +
                        " is absent or stale.",
                        "Replay the durable decision ledger and rebuild the plan.",
                        selection.SubjectId);
                }
                ValidateDecision(generation, state, decision, selection.SubjectId);

                if (!catalog.TryGetObservation(
                        state.SelectedObservationId,
                        out Observation selectedObservation) ||
                    !string.Equals(
                        selectedObservation.SubjectId,
                        selection.SubjectId,
                        StringComparison.Ordinal) ||
                    string.Equals(
                        selectedObservation.RevisionId,
                        catalog.MasterRevisionId,
                        StringComparison.Ordinal))
                {
                    Fail(
                        MergePlanBuildCodes.SelectionInvalid,
                        "The selected observation is absent, belongs to another subject, or names the master.",
                        "Review the candidate selection against freshly rebuilt evidence.",
                        selection.SubjectId);
                }
                if (!catalog.TryGetObservation(
                        selection.SubjectId,
                        catalog.MasterRevisionId,
                        out Observation masterObservation))
                {
                    Fail(
                        MergePlanBuildCodes.TargetInvalid,
                        "Subject " + selection.SubjectId +
                        " has no exact observation in the master revision.",
                        "Correct the alignment or select a master that contains an unambiguous target.",
                        selection.SubjectId);
                }
                if (!evidenceByRevision.TryGetValue(
                        selectedObservation.RevisionId,
                        out ReviewEvidenceRevisionIdentity selectedRevisionEvidence))
                {
                    Fail(
                        MergePlanBuildCodes.EvidenceInconsistent,
                        "The selected observation revision is outside the exact byte-evidence generation.",
                        "Rebuild the evidence generation and review the candidate selection again.",
                        selection.SubjectId);
                }

                ValidateDecisionSupport(decision, selectedObservation, selection.SubjectId);
                ValidateAspectState(subject, masterObservation, "master precondition");
                ValidateAspectState(subject, selectedObservation, "selected output");

                MergePlanCellState masterState;
                MergePlanCellState selectedState;
                try
                {
                    masterState = new MergePlanCellState(masterObservation);
                    selectedState = new MergePlanCellState(selectedObservation);
                }
                catch (Exception exception) when (
                    exception is ArgumentException ||
                    exception is OverflowException)
                {
                    Fail(
                        MergePlanBuildCodes.QuotaExceeded,
                        "Exact cell evidence for subject " + selection.SubjectId +
                        " is malformed or exceeds the plan contract: " + exception.Message,
                        "Rebuild bounded canonical evidence and inspect the source cell.",
                        selection.SubjectId);
                    throw;
                }

                bool contentIdentityEqual = string.Equals(
                    masterState.ContentHash,
                    selectedState.ContentHash,
                    StringComparison.Ordinal);
                bool payloadEqual = masterState.PayloadEquals(selectedState);
                if (contentIdentityEqual != payloadEqual)
                {
                    Fail(
                        MergePlanBuildCodes.EvidenceInconsistent,
                        "The exact payload and content identity disagree for subject " +
                        selection.SubjectId + ".",
                        "Discard the plan, rebuild canonical observations from the source bytes, and investigate the mismatch.",
                        selection.SubjectId);
                }
                if (payloadEqual)
                {
                    Fail(
                        MergePlanBuildCodes.NoActualChange,
                        "The selected candidate and master have the same exact aspect state for subject " +
                        selection.SubjectId + ".",
                        "Remove the redundant selection or refresh the comparison if a change was expected.",
                        selection.SubjectId);
                }

                MergePlanTarget target;
                MergePlanSourceOccurrence selectedSource;
                try
                {
                    selectedSource = new MergePlanSourceOccurrence(
                        selectedObservation.Provenance);
                }
                catch (Exception exception) when (
                    exception is ArgumentException ||
                    exception is OverflowException)
                {
                    Fail(
                        MergePlanBuildCodes.SourceInvalid,
                        "The selected source for subject " + selection.SubjectId +
                        " is incomplete or inconsistent: " + exception.Message,
                        "Correct the mapping and regenerate exact selected-revision provenance.",
                        selection.SubjectId);
                    throw;
                }
                try
                {
                    target = new MergePlanTarget(masterObservation.Provenance);
                }
                catch (Exception exception) when (
                    exception is ArgumentException ||
                    exception is OverflowException)
                {
                    Fail(
                        MergePlanBuildCodes.TargetInvalid,
                        "The master target for subject " + selection.SubjectId +
                        " is incomplete or inconsistent: " + exception.Message,
                        "Correct the mapping and regenerate exact master provenance.",
                        selection.SubjectId);
                    throw;
                }
                if (!string.Equals(
                        selectedSource.LogicalSheetId,
                        target.LogicalSheetId,
                        StringComparison.Ordinal))
                {
                    Fail(
                        MergePlanBuildCodes.EvidenceInconsistent,
                        "The selected source and master target for subject " +
                        selection.SubjectId +
                        " do not share one logical sheet identity.",
                        "Correct the sheet mapping and rebuild comparison evidence.",
                        selection.SubjectId);
                }

                ValidateSheetMapping(
                    target,
                    selection.SubjectId,
                    logicalToPhysicalSheet,
                    physicalToLogicalSheet);
                string physicalTargetKey =
                    target.SheetName + "\u001F" +
                    target.PhysicalRow.ToString(CultureInfo.InvariantCulture) + "\u001F" +
                    target.PhysicalColumn.ToString(CultureInfo.InvariantCulture) + "\u001F" +
                    ((int)subject.Aspect).ToString(CultureInfo.InvariantCulture);
                if (!physicalTargets.Add(physicalTargetKey))
                {
                    Fail(
                        MergePlanBuildCodes.TargetAmbiguous,
                        "Multiple selected subjects map the same aspect to master target " +
                        target.SheetName + "!" + target.Address + ".",
                        "Correct the logical alignment before building a merge plan.",
                        selection.SubjectId);
                }

                operations.Add(new MergePlanOperation(
                    subject,
                    decision,
                    selectedObservation,
                    selectedRevisionEvidence.ByteContentId,
                    masterObservation,
                    selectedSource,
                    target,
                    masterState,
                    selectedState));
                hasFormulaOperation |= subject.Aspect == ComparisonAspect.Formula;
            }

            operations.Sort(CompareOperations);
            IReadOnlyList<MergePlanOperation> frozenOperations =
                new ReadOnlyCollection<MergePlanOperation>(operations);
            MergePlanReviewSummary summary = BuildReviewSummary(
                generation,
                operations.Count,
                cancellationToken);
            string decisionSetHash = ComputeDecisionSetHash(
                generation,
                frozenOperations);
            MergePlanRequiredValidation requiredValidations = BaseRequiredValidations;
            if (hasFormulaOperation)
            {
                requiredValidations |=
                    MergePlanRequiredValidation.FormulaDependencies |
                    MergePlanRequiredValidation.FormulaGroups;
            }
            string planHash = ComputePlanHash(
                generation,
                masterEvidence.ByteContentId,
                decisionSetHash,
                requiredValidations,
                summary,
                frozenOperations);

            return new MergePlan(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.CatalogId,
                generation.GenerationId,
                generation.Catalog.MasterRevisionId,
                masterEvidence.ByteContentId,
                decisionSetHash,
                planHash,
                requiredValidations,
                summary,
                frozenOperations);
        }

        private static void ValidateContext(
            ReviewDecisionGeneration generation,
            MergePlanRequest request)
        {
            if (!string.Equals(
                    generation.ProjectId,
                    request.ProjectId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    generation.BaseGenerationId,
                    request.EvidenceGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    generation.MappingGenerationId,
                    request.MappingGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    generation.CatalogId,
                    request.CatalogId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    generation.GenerationId,
                    request.ExpectedDecisionGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    generation.Catalog.MasterRevisionId,
                    request.MasterRevisionId,
                    StringComparison.Ordinal))
            {
                Fail(
                    MergePlanBuildCodes.ContextMismatch,
                    "The request does not bind the exact current project, evidence, mapping, catalog, decision generation, and master.",
                    "Refresh the workspace and capture a new merge-plan request.");
            }
        }

        private static void ValidateDecision(
            ReviewDecisionGeneration generation,
            ReviewDecisionState state,
            DecisionEvent decision,
            string subjectId)
        {
            bool valid =
                decision.NewDecision == ReviewDecisionStatus.UseCandidate &&
                decision.SelectionScope == DecisionSelectionScope.SingleSubject &&
                string.Equals(decision.SubjectId, subjectId, StringComparison.Ordinal) &&
                decision.GroupId.Length == 0 &&
                decision.AffectedSubjectIds.Count == 1 &&
                string.Equals(
                    decision.AffectedSubjectIds[0],
                    subjectId,
                    StringComparison.Ordinal) &&
                string.Equals(
                    decision.SelectedObservationId,
                    state.SelectedObservationId,
                    StringComparison.Ordinal) &&
                decision.RequestId.Length != 0 &&
                string.Equals(
                    decision.ProjectId,
                    generation.ProjectId,
                    StringComparison.Ordinal) &&
                string.Equals(
                    decision.EvidenceGenerationId,
                    generation.BaseGenerationId,
                    StringComparison.Ordinal) &&
                string.Equals(
                    decision.MappingGenerationId,
                    generation.MappingGenerationId,
                    StringComparison.Ordinal);
            if (!valid)
            {
                Fail(
                    MergePlanBuildCodes.SelectionInvalid,
                    "The effective decision event for subject " + subjectId +
                    " is not a complete single-subject candidate selection bound to this generation.",
                    "Replay the decision ledger from its trusted root and rebuild the plan.",
                    subjectId);
            }
        }

        private static void ValidateDecisionSupport(
            DecisionEvent decision,
            Observation selectedObservation,
            string subjectId)
        {
            if (decision.SupportingRevisionIds.Count != 1 ||
                !string.Equals(
                    decision.SupportingRevisionIds[0],
                    selectedObservation.RevisionId,
                    StringComparison.Ordinal) ||
                decision.SupportingContentHashes.Count != 1 ||
                !string.Equals(
                    decision.SupportingContentHashes[0],
                    selectedObservation.ContentHash,
                    StringComparison.Ordinal))
            {
                Fail(
                    MergePlanBuildCodes.EvidenceInconsistent,
                    "The effective decision support does not identify the exact selected revision and content for subject " +
                    subjectId + ".",
                    "Replay the decision ledger and inspect the affected event before merging.",
                    subjectId);
            }
        }

        private static void ValidateAspectState(
            ReviewDecisionSubject subject,
            Observation observation,
            string role)
        {
            bool supported;
            if (subject.Aspect == ComparisonAspect.Formula)
            {
                supported =
                    observation.ExactType == ObservationExactType.Formula ||
                    observation.ExactType == ObservationExactType.Blank;
            }
            else
            {
                supported =
                    observation.ExactType != ObservationExactType.Formula &&
                    observation.ExactType != ObservationExactType.Missing &&
                    observation.ExactType != ObservationExactType.Unsupported;
            }
            if (!supported)
            {
                Fail(
                    MergePlanBuildCodes.StateUnsupported,
                    "The " + role + " for subject " + subject.SubjectId +
                    " has exact type " + observation.ExactType +
                    ", which has no unambiguous v1 formula/value write meaning.",
                    "Choose explicit blank or supported typed evidence; handle missing/unsupported states through review and structural correction.",
                    subject.SubjectId);
            }
        }

        private static void ValidateSheetMapping(
            MergePlanTarget target,
            string subjectId,
            IDictionary<string, string> logicalToPhysical,
            IDictionary<string, string> physicalToLogical)
        {
            if (logicalToPhysical.TryGetValue(
                    target.LogicalSheetId,
                    out string priorPhysical) &&
                !string.Equals(
                    priorPhysical,
                    target.SheetName,
                    StringComparison.Ordinal))
            {
                Fail(
                    MergePlanBuildCodes.TargetAmbiguous,
                    "Logical sheet " + target.LogicalSheetId +
                    " maps to multiple master worksheet names.",
                    "Correct the sheet mapping and rebuild the evidence generation.",
                    subjectId);
            }
            logicalToPhysical[target.LogicalSheetId] = target.SheetName;

            if (physicalToLogical.TryGetValue(
                    target.SheetName,
                    out string priorLogical) &&
                !string.Equals(
                    priorLogical,
                    target.LogicalSheetId,
                    StringComparison.Ordinal))
            {
                Fail(
                    MergePlanBuildCodes.TargetAmbiguous,
                    "Master worksheet " + target.SheetName +
                    " maps to multiple logical sheet identities.",
                    "Correct the sheet mapping and rebuild the evidence generation.",
                    subjectId);
            }
            physicalToLogical[target.SheetName] = target.LogicalSheetId;
        }

        private static MergePlanReviewSummary BuildReviewSummary(
            ReviewDecisionGeneration generation,
            int selectedCandidate,
            CancellationToken cancellationToken)
        {
            int unreviewed = 0;
            int useMaster = 0;
            int useCandidate = 0;
            int equivalent = 0;
            int needsInvestigation = 0;
            int reviewedNoStructuralMerge = 0;
            foreach (ReviewDecisionState state in generation.SubjectStates.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                switch (state.Status)
                {
                    case ReviewDecisionStatus.Unreviewed: unreviewed++; break;
                    case ReviewDecisionStatus.UseMaster: useMaster++; break;
                    case ReviewDecisionStatus.UseCandidate: useCandidate++; break;
                    case ReviewDecisionStatus.Equivalent: equivalent++; break;
                    case ReviewDecisionStatus.NeedsInvestigation: needsInvestigation++; break;
                    case ReviewDecisionStatus.ReviewedNoStructuralMerge:
                        reviewedNoStructuralMerge++;
                        break;
                    default:
                        Fail(
                            MergePlanBuildCodes.EvidenceInconsistent,
                            "The decision generation contains an unknown review status.",
                            "Replay the ledger with a supported contract.");
                        break;
                }
            }
            if (selectedCandidate > useCandidate)
            {
                Fail(
                    MergePlanBuildCodes.EvidenceInconsistent,
                    "The selected candidate count exceeds the generation's effective candidate decisions.",
                    "Replay the ledger and rebuild the merge-plan request.");
            }
            return new MergePlanReviewSummary(
                generation.SubjectStates.Count,
                unreviewed,
                useMaster,
                useCandidate,
                equivalent,
                needsInvestigation,
                reviewedNoStructuralMerge,
                selectedCandidate);
        }

        private static string ComputeDecisionSetHash(
            ReviewDecisionGeneration generation,
            IReadOnlyList<MergePlanOperation> operations)
        {
            var components = new List<string>
            {
                MergePlanContract.Version.ToString(CultureInfo.InvariantCulture),
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.CatalogId,
                generation.GenerationId,
                generation.Catalog.MasterRevisionId,
                operations.Count.ToString(CultureInfo.InvariantCulture),
            };
            foreach (MergePlanOperation operation in operations.OrderBy(
                         item => item.SubjectId,
                         StringComparer.Ordinal))
            {
                components.Add(operation.SubjectId);
                components.Add(operation.DecisionEventId);
                components.Add(operation.SelectedObservationId);
            }
            return StableId.Compute("merge-plan-decision-set-v1", components.ToArray());
        }

        private static string ComputePlanHash(
            ReviewDecisionGeneration generation,
            string masterByteContentId,
            string decisionSetHash,
            MergePlanRequiredValidation requiredValidations,
            MergePlanReviewSummary summary,
            IReadOnlyList<MergePlanOperation> operations)
        {
            var components = new List<string>
            {
                MergePlanContract.Version.ToString(CultureInfo.InvariantCulture),
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.CatalogId,
                generation.GenerationId,
                generation.Catalog.MasterRevisionId,
                masterByteContentId,
                decisionSetHash,
                ((int)requiredValidations).ToString(CultureInfo.InvariantCulture),
            };
            summary.AddHashComponents(components);
            components.Add(operations.Count.ToString(CultureInfo.InvariantCulture));
            components.AddRange(operations.Select(item => item.OperationId));
            return StableId.Compute("merge-plan-v1", components.ToArray());
        }

        private static int CompareOperations(
            MergePlanOperation left,
            MergePlanOperation right)
        {
            int value = StringComparer.Ordinal.Compare(
                left.Target.LogicalSheetId,
                right.Target.LogicalSheetId);
            if (value != 0) return value;
            value = left.Target.PhysicalRow.CompareTo(right.Target.PhysicalRow);
            if (value != 0) return value;
            value = left.Target.PhysicalColumn.CompareTo(right.Target.PhysicalColumn);
            if (value != 0) return value;
            value = left.Aspect.CompareTo(right.Aspect);
            return value != 0
                ? value
                : StringComparer.Ordinal.Compare(left.SubjectId, right.SubjectId);
        }

        private static void Fail(
            string code,
            string detail,
            string correctiveAction,
            string subjectId = "")
        {
            throw new MergePlanBuildException(
                code,
                detail,
                correctiveAction,
                subjectId);
        }
    }
}
