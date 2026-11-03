using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Lineage;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;
using ExcelPowerTools.UI;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Authoritative state for the Phase 2 comparison project. Original source locators are retained alongside
    /// exact content hashes so similarly named local copies remain unambiguous in the generated review.
    /// </summary>
    internal static class ComparisonProjectStateJson
    {
        internal static string Serialize(
            ComparisonProjectPreparationResult preparation,
            ComparisonProjectSetupResult setup,
            AlignmentCorrectionGeneration alignmentCorrections,
            OpenXmlComparisonProjectResult projectResult)
        {
            if (preparation == null) throw new ArgumentNullException(nameof(preparation));
            if (setup == null) throw new ArgumentNullException(nameof(setup));
            if (alignmentCorrections == null) throw new ArgumentNullException(nameof(alignmentCorrections));
            if (projectResult == null) throw new ArgumentNullException(nameof(projectResult));
            string expectedEffectiveMappingGenerationId = StableId.Compute(
                "ooxml-project-effective-mapping-generation-v1",
                setup.Mapping.GenerationId,
                alignmentCorrections.GenerationId);
            if (!string.Equals(preparation.ProjectId, projectResult.ProjectId, StringComparison.Ordinal) ||
                !string.Equals(setup.Mapping.GenerationId, projectResult.MappingGenerationId, StringComparison.Ordinal) ||
                !string.Equals(
                    setup.Mapping.GenerationId,
                    alignmentCorrections.BaseMappingGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    alignmentCorrections.GenerationId,
                    projectResult.AlignmentCorrectionGenerationId,
                    StringComparison.Ordinal) ||
                !string.Equals(
                    expectedEffectiveMappingGenerationId,
                    projectResult.EffectiveMappingGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("The project result does not match the confirmed project generation.");

            var artifacts = preparation.OrderedArtifacts.ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            var proposals = preparation.Inference.Proposals.ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            var json = new StringBuilder(64 * 1024);
            json.Append('{');
            // Alignment corrections are an additive extension to the v1 object. Existing v1 consumers can
            // continue to ignore unknown properties; changing the schema token would unnecessarily break them.
            Property(json, "schema", "comparison-project-phase2-v1", true);
            Property(json, "mode", "multi-revision-aligned-history-alpha", false);
            Property(json, "generatedUtc", DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture), false);
            Property(json, "projectId", preparation.ProjectId, false);
            Property(json, "inferenceProposalGenerationId", preparation.Inference.ProposalGenerationId, false);
            Property(json, "mappingGenerationId", setup.Mapping.GenerationId, false);
            Property(json, "alignmentCorrectionGenerationId", alignmentCorrections.GenerationId, false);
            Property(json, "effectiveMappingGenerationId", projectResult.EffectiveMappingGenerationId, false);
            BooleanProperty(json, "complete", projectResult.IsComplete, false);

            json.Append(",\"limitations\":[");
            string[] limitations =
            {
                "Multi-revision project input currently accepts .xlsx and .xlsm only",
                "Ambiguous sheet, row, or column identity blocks history instead of guessing",
                "User-facing review decisions, snapshots, dependency expansion, merge, archive export, and XLSB extraction are not implemented",
                "The embedded Phase 2 state preserves exact changed-subject history but not complete source package bytes",
                "Current in-process safety limits are 768 MiB of inputs; 2,000,000 physically stored cells per revision and 10,000,000 audited per project; 2,000,000 retained useful-range cells; and 10,000,000 bounded grid cells",
            };
            AppendStringArrayValues(json, limitations);
            json.Append(']');

            json.Append(",\"revisions\":[");
            for (int order = 0; order < setup.Mapping.OrderedRevisionIds.Count; order++)
            {
                if (order > 0) json.Append(',');
                string revisionId = setup.Mapping.OrderedRevisionIds[order];
                ComparisonProjectArtifactInfo artifact = artifacts[revisionId];
                LineageRevisionProposal proposal = proposals[revisionId];
                json.Append('{');
                Property(json, "revisionId", revisionId, true);
                NumberProperty(json, "order", order, false);
                Property(json, "label", setup.Mapping.GetLabel(revisionId), false);
                Property(json, "parentRevisionId", setup.Mapping.GetParent(revisionId) ?? "", false);
                BooleanProperty(json, "isMaster", string.Equals(
                    revisionId, setup.Mapping.MasterRevisionId, StringComparison.Ordinal), false);
                BooleanProperty(json, "isTrustedBaseline", string.Equals(
                    revisionId, setup.Mapping.TrustedBaselineRevisionId, StringComparison.Ordinal), false);
                Property(json, "sourceFileName", artifact.BaseName, false);
                Property(json, "sourcePath", artifact.Path, false);
                Property(json, "byteContentId", artifact.ByteContentId, false);
                NumberProperty(json, "byteLength", artifact.ByteLength, false);
                Property(json, "createdUtc", FormatTimestamp(artifact.CreatedUtc), false);
                Property(json, "modifiedUtc", FormatTimestamp(artifact.ModifiedUtc), false);
                Property(json, "format", artifact.FormatLabel, false);
                Property(json, "proposedParentRevisionId", proposal.ProposedParentRevisionId ?? "", false);
                json.Append(",\"orderConfidence\":");
                AppendLineageConfidence(json, proposal.OrderConfidence);
                json.Append(",\"parentConfidence\":");
                AppendLineageConfidence(json, proposal.ParentConfidence);
                json.Append(",\"orderReasons\":");
                AppendLineageReasons(json, proposal.OrderReasons);
                json.Append(",\"parentReasons\":");
                AppendLineageReasons(json, proposal.ParentReasons);
                json.Append('}');
            }
            json.Append(']');

            json.Append(",\"exactDuplicateGroups\":[");
            for (int index = 0; index < preparation.Inference.DuplicateGroups.Count; index++)
            {
                if (index > 0) json.Append(',');
                LineageDuplicateGroup group = preparation.Inference.DuplicateGroups[index];
                json.Append('{');
                Property(json, "groupId", group.GroupId, true);
                Property(json, "byteContentId", group.ByteContentId, false);
                json.Append(",\"revisionIds\":[");
                AppendStringArrayValues(json, group.RevisionIds);
                json.Append("]}");
            }
            json.Append(']');

            json.Append(",\"lineageCorrectionEvents\":[");
            for (int index = 0; index < setup.Mapping.Events.Count; index++)
            {
                if (index > 0) json.Append(',');
                LineageCorrectionEvent item = setup.Mapping.Events[index];
                json.Append('{');
                Property(json, "eventId", item.EventId, true);
                Property(json, "priorGenerationId", item.PriorGenerationId ?? "", false);
                Property(json, "newGenerationId", item.NewGenerationId, false);
                Property(json, "kind", item.Kind.ToString(), false);
                Property(json, "targetRevisionId", item.TargetRevisionId, false);
                Property(json, "previousValue", item.PreviousValue ?? "", false);
                Property(json, "newValue", item.NewValue ?? "", false);
                Property(json, "createdUtc", item.CreatedUtc.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture), false);
                Property(json, "note", item.Note, false);
                json.Append(",\"previousOrder\":[");
                AppendStringArrayValues(json, item.PreviousOrder);
                json.Append("],\"newOrder\":[");
                AppendStringArrayValues(json, item.NewOrder);
                json.Append("]}");
            }
            json.Append(']');

            json.Append(",\"alignmentCorrections\":");
            AppendAlignmentCorrections(
                json,
                alignmentCorrections,
                projectResult.EffectiveMappingGenerationId);

            json.Append(",\"edges\":[");
            for (int edgeIndex = 0; edgeIndex < projectResult.Edges.Count; edgeIndex++)
            {
                if (edgeIndex > 0) json.Append(',');
                AppendEdge(json, projectResult.Edges[edgeIndex]);
            }
            json.Append(']');

            json.Append(",\"alignedRevisions\":[");
            for (int revisionIndex = 0; revisionIndex < projectResult.Revisions.Count; revisionIndex++)
            {
                if (revisionIndex > 0) json.Append(',');
                AppendAlignedRevision(json, projectResult.Revisions[revisionIndex]);
            }
            json.Append(']');

            json.Append(",\"findings\":[");
            List<OpenXmlProjectFinding> findings = projectResult.Findings
                .OrderBy(item => item.FindingId, StringComparer.Ordinal)
                .ToList();
            for (int index = 0; index < findings.Count; index++)
            {
                if (index > 0) json.Append(',');
                AppendFinding(json, findings[index]);
            }
            json.Append(']');

            json.Append(",\"history\":");
            if (projectResult.History == null) json.Append("null");
            else AppendHistory(json, projectResult.History);
            json.Append('}');

            string result = json.ToString();
            int bytes = new UTF8Encoding(false, true).GetByteCount(result);
            if (bytes > ReviewWorkbookFile.MaxEmbeddedStateBytes)
                throw new InvalidDataException(
                    "The comparison project requires " + bytes.ToString("N0", CultureInfo.InvariantCulture) +
                    " embedded state bytes, exceeding the current Review.xlsx limit of " +
                    ReviewWorkbookFile.MaxEmbeddedStateBytes.ToString("N0", CultureInfo.InvariantCulture) +
                    ". Use a smaller revision set until archive-backed project storage is implemented.");
            return result;
        }

        private static void AppendAlignmentCorrections(
            StringBuilder json,
            AlignmentCorrectionGeneration corrections,
            string effectiveMappingGenerationId)
        {
            json.Append('{');
            Property(json, "baseMappingGenerationId", corrections.BaseMappingGenerationId, true);
            Property(json, "generationId", corrections.GenerationId, false);
            Property(json, "priorGenerationId", corrections.PriorGenerationId, false);
            Property(json, "effectiveMappingGenerationId", effectiveMappingGenerationId, false);

            json.Append(",\"activeSheetPairConfirmations\":[");
            for (int index = 0; index < corrections.SheetPairs.Count; index++)
            {
                if (index > 0) json.Append(',');
                AppendSheetPairConfirmation(json, corrections.SheetPairs[index]);
            }

            json.Append("],\"activeUnmatchedSheetConfirmations\":[");
            for (int index = 0; index < corrections.UnmatchedSheets.Count; index++)
            {
                if (index > 0) json.Append(',');
                AppendSheetUnmatchedConfirmation(json, corrections.UnmatchedSheets[index]);
            }

            json.Append("],\"activeHeaderConfirmations\":[");
            for (int index = 0; index < corrections.HeaderRows.Count; index++)
            {
                if (index > 0) json.Append(',');
                AppendHeaderRowConfirmation(json, corrections.HeaderRows[index]);
            }

            json.Append("],\"activeKeyColumnConfirmations\":[");
            for (int index = 0; index < corrections.KeyColumns.Count; index++)
            {
                if (index > 0) json.Append(',');
                AppendKeyColumnsConfirmation(json, corrections.KeyColumns[index]);
            }

            json.Append("],\"events\":[");
            for (int index = 0; index < corrections.Events.Count; index++)
            {
                if (index > 0) json.Append(',');
                AppendAlignmentCorrectionEvent(json, corrections.Events[index]);
            }
            json.Append("]}");
        }

        private static void AppendAlignmentCorrectionEvent(
            StringBuilder json,
            AlignmentCorrectionEvent item)
        {
            json.Append('{');
            Property(json, "eventId", item.EventId, true);
            Property(json, "priorGenerationId", item.PriorGenerationId, false);
            Property(json, "newGenerationId", item.NewGenerationId, false);
            Property(json, "kind", item.Kind.ToString(), false);
            Property(json, "requestId", item.RequestId, false);
            Property(json, "confirmationId", item.ConfirmationId, false);
            Property(json, "supersededConfirmationId", item.SupersededConfirmationId, false);
            Property(
                json,
                "createdUtc",
                item.CreatedUtc.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture),
                false);
            Property(json, "note", item.Note, false);
            json.Append(",\"reviewer\":");
            AppendReviewerSnapshot(json, item.Reviewer);
            json.Append(",\"scope\":");
            AppendAlignmentCorrectionScope(json, item.Scope);
            json.Append(",\"confirmation\":");
            switch (item.Kind)
            {
                case AlignmentCorrectionKind.ConfirmSheetPair:
                    AppendSheetPairConfirmation(json, item.SheetPair);
                    break;
                case AlignmentCorrectionKind.ConfirmSheetUnmatched:
                    AppendSheetUnmatchedConfirmation(json, item.SheetUnmatched);
                    break;
                case AlignmentCorrectionKind.ConfirmHeaderRow:
                    AppendHeaderRowConfirmation(json, item.HeaderRow);
                    break;
                case AlignmentCorrectionKind.ConfirmKeyColumns:
                    AppendKeyColumnsConfirmation(json, item.KeyColumns);
                    break;
                case AlignmentCorrectionKind.ClearSheetPair:
                case AlignmentCorrectionKind.ClearSheetUnmatched:
                case AlignmentCorrectionKind.ClearHeaderRow:
                case AlignmentCorrectionKind.ClearKeyColumns:
                    json.Append("null");
                    break;
                default:
                    throw new InvalidOperationException(
                        "The alignment correction ledger contains an unknown correction kind.");
            }
            json.Append('}');
        }

        private static void AppendAlignmentCorrectionScope(
            StringBuilder json,
            AlignmentCorrectionScope scope)
        {
            if (scope == null)
                throw new InvalidOperationException("An alignment correction event has no exact scope payload.");
            json.Append('{');
            Property(json, "parentRevisionId", scope.ParentRevisionId, true);
            Property(json, "childRevisionId", scope.ChildRevisionId, false);
            Property(json, "revisionId", scope.RevisionId, false);
            Property(json, "beforeSheetKey", scope.BeforeSheetKey, false);
            Property(json, "afterSheetKey", scope.AfterSheetKey, false);
            Property(json, "physicalSheetKey", scope.PhysicalSheetKey, false);
            json.Append('}');
        }

        private static void AppendReviewerSnapshot(StringBuilder json, ReviewerSnapshot reviewer)
        {
            reviewer = reviewer ?? ReviewerSnapshot.Empty;
            json.Append('{');
            Property(json, "confirmedDisplayName", reviewer.ConfirmedDisplayName, true);
            Property(json, "officeDisplayName", reviewer.OfficeDisplayName, false);
            json.Append(",\"identitySourceLabels\":[");
            AppendStringArrayValues(json, reviewer.IdentitySourceLabels);
            json.Append("]}");
        }

        private static void AppendSheetPairConfirmation(
            StringBuilder json,
            AlignmentSheetPairConfirmation confirmation)
        {
            if (confirmation == null)
                throw new InvalidOperationException("A sheet-pair correction event has no confirmation payload.");
            json.Append('{');
            Property(json, "confirmationId", confirmation.ConfirmationId, true);
            Property(json, "parentRevisionId", confirmation.ParentRevisionId, false);
            Property(json, "childRevisionId", confirmation.ChildRevisionId, false);
            Property(json, "beforeSheetKey", confirmation.BeforeSheetKey, false);
            Property(json, "afterSheetKey", confirmation.AfterSheetKey, false);
            json.Append('}');
        }

        private static void AppendSheetUnmatchedConfirmation(
            StringBuilder json,
            AlignmentSheetUnmatchedConfirmation confirmation)
        {
            if (confirmation == null)
                throw new InvalidOperationException("An unmatched-sheet correction event has no confirmation payload.");
            json.Append('{');
            Property(json, "confirmationId", confirmation.ConfirmationId, true);
            Property(json, "parentRevisionId", confirmation.ParentRevisionId, false);
            Property(json, "childRevisionId", confirmation.ChildRevisionId, false);
            Property(json, "beforeSheetKey", confirmation.BeforeSheetKey, false);
            json.Append('}');
        }

        private static void AppendHeaderRowConfirmation(
            StringBuilder json,
            AlignmentHeaderRowConfirmation confirmation)
        {
            if (confirmation == null)
                throw new InvalidOperationException("A header correction event has no confirmation payload.");
            json.Append('{');
            Property(json, "confirmationId", confirmation.ConfirmationId, true);
            Property(json, "revisionId", confirmation.RevisionId, false);
            Property(json, "physicalSheetKey", confirmation.PhysicalSheetKey, false);
            Property(json, "decision", confirmation.Decision.ToString(), false);
            json.Append(",\"headerRow\":");
            AppendRowReference(json, confirmation.Row);
            json.Append(",\"firstDataRow\":");
            AppendRowReference(json, confirmation.FirstDataRow);
            json.Append('}');
        }

        private static void AppendKeyColumnsConfirmation(
            StringBuilder json,
            AlignmentKeyColumnsConfirmation confirmation)
        {
            if (confirmation == null)
                throw new InvalidOperationException("A key-column correction event has no confirmation payload.");
            json.Append('{');
            Property(json, "confirmationId", confirmation.ConfirmationId, true);
            Property(json, "parentRevisionId", confirmation.ParentRevisionId, false);
            Property(json, "childRevisionId", confirmation.ChildRevisionId, false);
            Property(json, "beforeSheetKey", confirmation.BeforeSheetKey, false);
            Property(json, "afterSheetKey", confirmation.AfterSheetKey, false);
            json.Append(",\"beforeColumns\":[");
            AppendColumnReferences(json, confirmation.BeforeColumns);
            json.Append("],\"afterColumns\":[");
            AppendColumnReferences(json, confirmation.AfterColumns);
            json.Append("]}");
        }

        private static void AppendRowReference(
            StringBuilder json,
            AlignmentGridIndexReference reference)
        {
            if (reference == null)
            {
                json.Append("null");
                return;
            }
            json.Append('{');
            NumberProperty(json, "tableRowIndex", reference.TableIndex, true);
            NumberProperty(json, "oneBasedPhysicalRow", reference.PhysicalIndex, false);
            json.Append('}');
        }

        private static void AppendColumnReferences(
            StringBuilder json,
            IEnumerable<AlignmentGridIndexReference> references)
        {
            int index = 0;
            foreach (AlignmentGridIndexReference reference in references ??
                         Array.Empty<AlignmentGridIndexReference>())
            {
                if (reference == null)
                    throw new InvalidOperationException("A key-column confirmation contains a null coordinate.");
                if (index++ > 0) json.Append(',');
                json.Append('{');
                NumberProperty(json, "tableColumnIndex", reference.TableIndex, true);
                NumberProperty(json, "oneBasedPhysicalColumn", reference.PhysicalIndex, false);
                json.Append('}');
            }
        }

        private static void AppendEdge(StringBuilder json, OpenXmlProjectEdgeResult edge)
        {
            json.Append('{');
            Property(json, "parentRevisionId", edge.ParentRevisionId, true);
            Property(json, "childRevisionId", edge.ChildRevisionId, false);
            BooleanProperty(json, "complete", edge.IsComplete, false);
            json.Append(",\"resolvedSheetMappings\":[");
            for (int index = 0; index < edge.SheetMapping.Mapping.Mappings.Count; index++)
            {
                if (index > 0) json.Append(',');
                ResolvedSheetMapping mapping = edge.SheetMapping.Mapping.Mappings[index];
                json.Append('{');
                Property(json, "beforeSheetName", mapping.Before.Name, true);
                Property(json, "afterSheetName", mapping.After.Name, false);
                NumberProperty(json, "beforeSheetIndex", SheetIndex(
                    edge.SheetMapping.Before, mapping.Before.SheetKey), false);
                NumberProperty(json, "afterSheetIndex", SheetIndex(
                    edge.SheetMapping.After, mapping.After.SheetKey), false);
                Property(json, "kind", mapping.Kind.ToString(), false);
                DoubleProperty(json, "confidence", mapping.Confidence, false);
                json.Append(",\"reasonCodes\":[");
                AppendStringArrayValues(json, mapping.ReasonCodes);
                json.Append("]}");
            }
            json.Append("],\"unresolvedSheetMappings\":[");
            for (int index = 0; index < edge.SheetMapping.Mapping.UnresolvedBeforeSheets.Count; index++)
            {
                if (index > 0) json.Append(',');
                UnresolvedSheetMapping unresolved = edge.SheetMapping.Mapping.UnresolvedBeforeSheets[index];
                json.Append('{');
                Property(json, "beforeSheetName", unresolved.Before.Name, true);
                NumberProperty(json, "beforeSheetIndex", SheetIndex(
                    edge.SheetMapping.Before, unresolved.Before.SheetKey), false);
                json.Append(",\"reasonCodes\":[");
                AppendStringArrayValues(json, unresolved.ReasonCodes);
                json.Append("],\"candidates\":[");
                for (int candidate = 0; candidate < unresolved.Candidates.Count; candidate++)
                {
                    if (candidate > 0) json.Append(',');
                    SheetMappingCandidate item = unresolved.Candidates[candidate];
                    json.Append('{');
                    Property(json, "afterSheetName", item.After.Name, true);
                    NumberProperty(json, "afterSheetIndex", SheetIndex(
                        edge.SheetMapping.After, item.After.SheetKey), false);
                    DoubleProperty(json, "confidence", item.Confidence, false);
                    json.Append(",\"reasonCodes\":[");
                    AppendStringArrayValues(json, item.ReasonCodes);
                    json.Append("]}");
                }
                json.Append("]}");
            }
            json.Append("],\"unmatchedBeforeSheets\":[");
            AppendStringArrayValues(json, edge.SheetMapping.Mapping.UnmatchedBeforeSheets.Select(item => item.Name));
            json.Append("],\"unmatchedAfterSheets\":[");
            AppendStringArrayValues(json, edge.SheetMapping.Mapping.UnmatchedAfterSheets.Select(item => item.Name));
            json.Append("],\"sheetAlignments\":[");
            for (int index = 0; index < edge.SheetAlignments.Count; index++)
            {
                if (index > 0) json.Append(',');
                OpenXmlProjectSheetAlignment alignment = edge.SheetAlignments[index];
                json.Append('{');
                Property(json, "logicalSheetId", alignment.LogicalSheetId, true);
                Property(json, "beforeSheetName", alignment.BeforeSheetName, false);
                Property(json, "afterSheetName", alignment.AfterSheetName, false);
                BooleanProperty(json, "complete", alignment.IsComplete, false);
                Property(json, "beforeHeaderDecision", alignment.BeforeHeader?.Decision.ToString() ?? "", false);
                Property(json, "afterHeaderDecision", alignment.AfterHeader?.Decision.ToString() ?? "", false);
                Property(json, "keyDecision", alignment.KeyInference?.Decision.ToString() ?? "", false);
                json.Append(",\"rows\":");
                AppendAxisSummary(json, alignment.RowAlignment);
                json.Append(",\"columns\":");
                AppendAxisSummary(json, alignment.ColumnAlignment);
                json.Append('}');
            }
            json.Append("],\"sheetStructuralCoverage\":");
            AppendSheetStructuralCoverage(json, edge);
            json.Append('}');
        }

        /// <summary>
        /// Records only structure that follows from accepted sheet mappings or an explicit deleted/new choice.
        /// An unresolved Before sheet is not called a deletion because it may instead be a rename or replacement;
        /// the durable marker makes that refusal visible. Relative ordering is compared only across accepted one-to-one mappings, so inserted
        /// sheets cannot create a false reorder event merely by shifting numeric sheet positions.
        /// </summary>
        private static void AppendSheetStructuralCoverage(
            StringBuilder json,
            OpenXmlProjectEdgeResult edge)
        {
            List<ResolvedSheetMapping> byBefore = edge.SheetMapping.Mapping.Mappings
                .OrderBy(item => SheetIndex(edge.SheetMapping.Before, item.Before.SheetKey))
                .ThenBy(item => item.Before.SheetKey, StringComparer.Ordinal)
                .ToList();
            List<ResolvedSheetMapping> byAfter = edge.SheetMapping.Mapping.Mappings
                .OrderBy(item => SheetIndex(edge.SheetMapping.After, item.After.SheetKey))
                .ThenBy(item => item.After.SheetKey, StringComparer.Ordinal)
                .ToList();
            bool mappedOrderChanged = !byBefore.Select(item => item.Before.SheetKey)
                .SequenceEqual(byAfter.Select(item => item.Before.SheetKey), StringComparer.Ordinal);

            json.Append('{');
            BooleanProperty(json, "identityCoverageComplete",
                edge.SheetMapping.Mapping.UnresolvedBeforeSheets.Count == 0, true);
            BooleanProperty(json, "relativeMappedOrderChanged", mappedOrderChanged, false);
            json.Append(",\"beforeMappedOrder\":[");
            AppendStringArrayValues(json, byBefore.Select(item => item.Before.Name));
            json.Append("],\"afterMappedOrder\":[");
            AppendStringArrayValues(json, byAfter.Select(item => item.After.Name));
            json.Append("],\"markers\":[");

            int markerIndex = 0;
            if (mappedOrderChanged)
            {
                AppendSheetStructuralMarkerPrefix(
                    json,
                    ref markerIndex,
                    StableId.Compute(
                        "comparison-project-sheet-structural-marker-v1",
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        "CMP-PROJECT-SHEET-ORDER-CHANGED"),
                    "CMP-PROJECT-SHEET-ORDER-CHANGED",
                    "Observed",
                    "The relative order of accepted one-to-one sheet identities changed between these revisions.",
                    "Review the mapped Before and After order. This event records ordering only and does not guess which individual sheet was moved.");
                json.Append(",\"beforeMappedOrder\":[");
                AppendStringArrayValues(json, byBefore.Select(item => item.Before.Name));
                json.Append("],\"afterMappedOrder\":[");
                AppendStringArrayValues(json, byAfter.Select(item => item.After.Name));
                json.Append('}');
            }

            foreach (UnresolvedSheetMapping unresolved in edge.SheetMapping.Mapping.UnresolvedBeforeSheets
                         .OrderBy(item => SheetIndex(edge.SheetMapping.Before, item.Before.SheetKey))
                         .ThenBy(item => item.Before.SheetKey, StringComparer.Ordinal))
            {
                AppendSheetStructuralMarkerPrefix(
                    json,
                    ref markerIndex,
                    StableId.Compute(
                        "comparison-project-sheet-structural-marker-v1",
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        "CMP-PROJECT-SHEET-DELETION-UNRESOLVED",
                        unresolved.Before.SheetKey),
                    "CMP-PROJECT-SHEET-DELETION-UNRESOLVED",
                    "Unresolved",
                    "Before sheet '" + unresolved.Before.Name +
                    "' has no accepted identity in the child revision; deletion, rename, or replacement cannot be distinguished safely.",
                    "Confirm a rename target or confirm that the sheet was deleted before treating this as a logical deletion.");
                Property(json, "beforeSheetName", unresolved.Before.Name, false);
                NumberProperty(json, "beforeSheetIndex", SheetIndex(
                    edge.SheetMapping.Before, unresolved.Before.SheetKey), false);
                json.Append(",\"reasonCodes\":[");
                AppendStringArrayValues(json, unresolved.ReasonCodes);
                json.Append("],\"candidateSheetNames\":[");
                AppendStringArrayValues(json, unresolved.Candidates.Select(item => item.After.Name));
                json.Append("]}");
            }

            foreach (WorkbookSheetFingerprint before in edge.SheetMapping.Mapping.UnmatchedBeforeSheets
                         .OrderBy(item => SheetIndex(edge.SheetMapping.Before, item.SheetKey))
                         .ThenBy(item => item.SheetKey, StringComparer.Ordinal))
            {
                AppendSheetStructuralMarkerPrefix(
                    json,
                    ref markerIndex,
                    StableId.Compute(
                        "comparison-project-sheet-structural-marker-v1",
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        "CMP-PROJECT-SHEET-DELETED",
                        before.SheetKey),
                    "CMP-PROJECT-SHEET-DELETED",
                    "Confirmed",
                    "Before sheet '" + before.Name +
                    "' was explicitly confirmed as having no successor in the child revision.",
                    "The sheet remains a logical deletion. Unclaimed child sheets are recorded separately as insertions.");
                Property(json, "beforeSheetName", before.Name, false);
                NumberProperty(json, "beforeSheetIndex", SheetIndex(
                    edge.SheetMapping.Before, before.SheetKey), false);
                json.Append('}');
            }

            foreach (WorkbookSheetFingerprint after in edge.SheetMapping.Mapping.UnmatchedAfterSheets
                         .OrderBy(item => SheetIndex(edge.SheetMapping.After, item.SheetKey))
                         .ThenBy(item => item.SheetKey, StringComparer.Ordinal))
            {
                bool explicitRenameCandidate = edge.SheetMapping.Mapping.UnresolvedBeforeSheets.Any(unresolved =>
                    unresolved.Candidates.Any(candidate => string.Equals(
                        candidate.After.SheetKey,
                        after.SheetKey,
                        StringComparison.Ordinal)));
                bool identityUnresolved = edge.SheetMapping.Mapping.UnresolvedBeforeSheets.Count != 0;
                string code = identityUnresolved
                    ? "CMP-PROJECT-SHEET-ONLY-AFTER-UNRESOLVED"
                    : "CMP-PROJECT-SHEET-INSERTED";
                AppendSheetStructuralMarkerPrefix(
                    json,
                    ref markerIndex,
                    StableId.Compute(
                        "comparison-project-sheet-structural-marker-v1",
                        edge.ParentRevisionId,
                        edge.ChildRevisionId,
                        code,
                        after.SheetKey),
                    code,
                    identityUnresolved ? "Unresolved" : "Observed",
                    "After sheet '" + after.Name +
                    "' has no accepted identity in the parent revision.",
                    identityUnresolved
                        ? "Resolve every unmatched Before sheet before treating this occurrence as a logical insertion; it may be a rename or replacement."
                        : "The accepted one-to-one mapping leaves no possible predecessor, so this occurrence has a new logical sheet identity.");
                Property(json, "afterSheetName", after.Name, false);
                NumberProperty(json, "afterSheetIndex", SheetIndex(
                    edge.SheetMapping.After, after.SheetKey), false);
                BooleanProperty(json, "isExplicitRenameCandidate", explicitRenameCandidate, false);
                json.Append('}');
            }

            json.Append("]}");
        }

        private static void AppendSheetStructuralMarkerPrefix(
            StringBuilder json,
            ref int markerIndex,
            string markerId,
            string code,
            string status,
            string message,
            string suggestedAction)
        {
            if (markerIndex++ > 0) json.Append(',');
            json.Append('{');
            Property(json, "markerId", markerId, true);
            Property(json, "code", code, false);
            Property(json, "status", status, false);
            Property(json, "message", message, false);
            Property(json, "suggestedAction", suggestedAction, false);
            Property(json, "source", OpenXmlProjectReviewProjector.EngineSourceHint, false);
        }

        private static int SheetIndex(
            OpenXmlWorkbookSheetFingerprintSet workbook,
            string sheetKey)
        {
            for (int index = 0; index < workbook.Sheets.Count; index++)
                if (string.Equals(
                    workbook.Sheets[index].Fingerprint.SheetKey,
                    sheetKey,
                    StringComparison.Ordinal))
                    return index;
            return -1;
        }

        private static void AppendAlignedRevision(
            StringBuilder json,
            OpenXmlAlignedProjectRevision revision)
        {
            json.Append('{');
            Property(json, "revisionId", revision.Source.RevisionId, true);
            BooleanProperty(json, "complete", revision.IsComplete, false);
            json.Append(",\"sheets\":[");
            for (int index = 0; index < revision.Sheets.Count; index++)
            {
                if (index > 0) json.Append(',');
                OpenXmlAlignedSheetRevision sheet = revision.Sheets[index];
                json.Append('{');
                Property(json, "logicalSheetId", sheet.LogicalSheetId, true);
                Property(json, "physicalSheetName", sheet.PhysicalSheetName, false);
                Property(json, "physicalSheetKey", sheet.PhysicalSheetKey, false);
                json.Append(",\"canonicalGrid\":{");
                NumberProperty(json, "firstPhysicalRow", sheet.Grid.FirstPhysicalRow, true);
                NumberProperty(json, "firstPhysicalColumn", sheet.Grid.FirstPhysicalColumn, false);
                NumberProperty(json, "rowCount", sheet.Grid.Table.RowCount, false);
                NumberProperty(json, "columnCount", sheet.Grid.Table.ColumnCount, false);
                NumberProperty(json, "lastPhysicalRow", sheet.Grid.Table.RowCount == 0
                    ? 0
                    : checked(sheet.Grid.FirstPhysicalRow + sheet.Grid.Table.RowCount - 1), false);
                NumberProperty(json, "lastPhysicalColumn", sheet.Grid.Table.ColumnCount == 0
                    ? 0
                    : checked(sheet.Grid.FirstPhysicalColumn + sheet.Grid.Table.ColumnCount - 1), false);
                NumberProperty(json, "materializedRectangleCellCount", checked(
                    (long)sheet.Grid.Table.RowCount * sheet.Grid.Table.ColumnCount), false);
                NumberProperty(json, "storedCanonicalCellCount", sheet.Grid.Source.Cells.Count, false);
                json.Append('}');
                json.Append(",\"rowProjection\":");
                AppendAxisProjection(json, sheet.Rows);
                json.Append(",\"columnProjection\":");
                AppendAxisProjection(json, sheet.Columns);
                json.Append('}');
            }
            json.Append("]}");
        }

        private static void AppendAxisProjection(
            StringBuilder json,
            LogicalAxisIdentityProjection projection)
        {
            if (projection == null)
            {
                json.Append("null");
                return;
            }

            json.Append('{');
            Property(json, "projectionId", projection.ProjectionId, true);
            Property(json, "parentProjectionId", projection.ParentProjectionId, false);
            Property(json, "revisionId", projection.RevisionId, false);
            Property(json, "parentRevisionId", projection.ParentRevisionId, false);
            Property(json, "axis", projection.Axis.ToString(), false);
            NumberProperty(json, "physicalCount", projection.PhysicalCount, false);
            Property(json, "confidence", projection.Confidence.ToString(), false);
            json.Append(",\"assignments\":[");
            int assignmentIndex = 0;
            foreach (KeyValuePair<int, LogicalAxisIdentityAssignment> pair in
                     projection.AssignmentsByPhysicalIndex.OrderBy(item => item.Key))
            {
                if (assignmentIndex++ > 0) json.Append(',');
                LogicalAxisIdentityAssignment assignment = pair.Value;
                json.Append('{');
                Property(json, "assignmentId", assignment.AssignmentId, true);
                Property(json, "logicalIdentityId", assignment.LogicalIdentityId, false);
                NumberProperty(json, "physicalIndex", assignment.PhysicalIndex, false);
                Property(json, "source", assignment.Source.ToString(), false);
                Property(json, "originRevisionId", assignment.OriginRevisionId, false);
                NumberProperty(json, "originPhysicalIndex", assignment.OriginPhysicalIndex, false);
                NullableNumberProperty(json, "parentPhysicalIndex", assignment.ParentPhysicalIndex, false);
                Property(json, "confidence", assignment.Confidence.ToString(), false);
                Property(json, "strategy", assignment.Strategy.ToString(), false);
                json.Append(",\"reasons\":");
                AppendAlignmentReasons(json, assignment.Reasons);
                json.Append('}');
            }
            json.Append("],\"structuralEvents\":[");
            for (int index = 0; index < projection.StructuralEvents.Count; index++)
            {
                if (index > 0) json.Append(',');
                LogicalAxisStructuralEvent item = projection.StructuralEvents[index];
                json.Append('{');
                Property(json, "eventId", item.EventId, true);
                Property(json, "kind", item.Kind.ToString(), false);
                Property(json, "parentRevisionId", item.ParentRevisionId, false);
                Property(json, "childRevisionId", item.ChildRevisionId, false);
                Property(json, "logicalIdentityId", item.LogicalIdentityId, false);
                NullableNumberProperty(json, "beforeIndex", item.BeforeIndex, false);
                NullableNumberProperty(json, "afterIndex", item.AfterIndex, false);
                Property(json, "confidence", item.Confidence.ToString(), false);
                Property(json, "strategy", item.Strategy.ToString(), false);
                json.Append(",\"reasons\":");
                AppendAlignmentReasons(json, item.Reasons);
                json.Append('}');
            }
            json.Append("],\"ambiguities\":[");
            for (int index = 0; index < projection.Ambiguities.Count; index++)
            {
                if (index > 0) json.Append(',');
                LogicalAxisIdentityAmbiguity item = projection.Ambiguities[index];
                json.Append('{');
                Property(json, "ambiguityId", item.AmbiguityId, true);
                Property(json, "kind", item.Kind.ToString(), false);
                Property(json, "evidenceToken", item.EvidenceToken, false);
                Property(json, "confidence", item.Confidence.ToString(), false);
                Property(json, "strategy", item.Strategy.ToString(), false);
                json.Append(",\"beforeIndices\":[");
                AppendNumberArrayValues(json, item.BeforeIndices);
                json.Append("],\"afterIndices\":[");
                AppendNumberArrayValues(json, item.AfterIndices);
                json.Append("],\"knownParentIdentities\":[");
                int knownIndex = 0;
                foreach (KeyValuePair<int, string> known in item.KnownParentIdentities.OrderBy(entry => entry.Key))
                {
                    if (knownIndex++ > 0) json.Append(',');
                    json.Append('{');
                    NumberProperty(json, "physicalIndex", known.Key, true);
                    Property(json, "logicalIdentityId", known.Value, false);
                    json.Append('}');
                }
                json.Append("],\"reasons\":");
                AppendAlignmentReasons(json, item.Reasons);
                json.Append('}');
            }
            json.Append("],\"unmappedPhysicalIndices\":[");
            AppendNumberArrayValues(json, projection.UnmappedPhysicalIndices);
            json.Append("]}");
        }

        private static void AppendAlignmentReasons(
            StringBuilder json,
            IEnumerable<AlignmentReason> reasons)
        {
            json.Append('[');
            int index = 0;
            foreach (AlignmentReason reason in reasons ?? Array.Empty<AlignmentReason>())
            {
                if (index++ > 0) json.Append(',');
                json.Append('{');
                Property(json, "code", reason.Code, true);
                Property(json, "message", reason.Message, false);
                DoubleProperty(json, "scoreImpact", reason.ScoreImpact, false);
                json.Append('}');
            }
            json.Append(']');
        }

        private static void AppendAxisSummary(StringBuilder json, AxisAlignmentResult axis)
        {
            if (axis == null)
            {
                json.Append("null");
                return;
            }
            int matched = axis.Links.Count(item => item.Kind == AxisAlignmentKind.Matched);
            int inserted = axis.Links.Count(item => item.Kind == AxisAlignmentKind.OnlyAfter);
            int deleted = axis.Links.Count(item => item.Kind == AxisAlignmentKind.OnlyBefore);
            AlignmentConfidence confidence = matched == 0
                ? AlignmentConfidence.None
                : axis.Links.Where(item => item.Kind == AxisAlignmentKind.Matched)
                    .Min(item => item.Confidence);
            json.Append('{');
            Property(json, "strategy", axis.Strategy.ToString(), true);
            Property(json, "minimumMatchedConfidence", confidence.ToString(), false);
            NumberProperty(json, "matched", matched, false);
            NumberProperty(json, "inserted", inserted, false);
            NumberProperty(json, "deleted", deleted, false);
            NumberProperty(json, "reordered", axis.Links.Count(item => item.Reordered), false);
            NumberProperty(json, "ambiguityBlocks", axis.Ambiguities.Count, false);
            json.Append(",\"reasonCodes\":[");
            AppendStringArrayValues(json, axis.Reasons.Select(item => item.Code)
                .Concat(axis.Ambiguities.SelectMany(item => item.Reasons.Select(reason => reason.Code))));
            json.Append("]}");
        }

        private static void AppendFinding(StringBuilder json, OpenXmlProjectFinding finding)
        {
            json.Append('{');
            Property(json, "findingId", finding.FindingId, true);
            Property(json, "code", finding.Code, false);
            Property(json, "severity", finding.Severity.ToString(), false);
            Property(json, "beforeRevisionId", finding.BeforeRevisionId, false);
            Property(json, "afterRevisionId", finding.AfterRevisionId, false);
            Property(json, "logicalSheetId", finding.LogicalSheetId, false);
            Property(json, "sheetName", finding.SheetName, false);
            Property(json, "message", finding.Message, false);
            Property(json, "suggestedAction", finding.SuggestedAction, false);
            Property(json, "source", OpenXmlProjectReviewProjector.EngineSourceHint, false);
            json.Append(",\"reasonCodes\":[");
            AppendStringArrayValues(json, finding.ReasonCodes);
            json.Append("],\"candidateSheetNames\":[");
            AppendStringArrayValues(json, finding.CandidateSheetNames);
            json.Append("]}");
        }

        private static void AppendHistory(StringBuilder json, LogicalHistoryResult history)
        {
            json.Append('{');
            Property(json, "masterRevisionId", history.MasterRevisionId, true);
            Property(json, "trustedBaselineRevisionId", history.TrustedBaselineRevisionId, false);
            NumberProperty(json, "subjectCount", history.Subjects.Count, false);
            NumberProperty(json, "atomicTransitionCount", history.AtomicTransitionCount, false);
            json.Append(",\"groups\":[");
            for (int index = 0; index < history.Groups.Count; index++)
            {
                if (index > 0) json.Append(',');
                HistoryChangeGroup group = history.Groups[index];
                json.Append('{');
                Property(json, "groupId", group.GroupId, true);
                Property(json, "edgeId", group.EdgeId, false);
                Property(json, "aspect", group.Aspect.ToString(), false);
                Property(json, "beforeStateId", group.BeforeStateId, false);
                Property(json, "afterStateId", group.AfterStateId, false);
                json.Append(",\"subjectIds\":[");
                AppendStringArrayValues(json, group.SubjectIds);
                json.Append("]}");
            }
            json.Append("],\"subjects\":[");
            for (int subjectIndex = 0; subjectIndex < history.Subjects.Count; subjectIndex++)
            {
                if (subjectIndex > 0) json.Append(',');
                LogicalSubjectHistory subject = history.Subjects[subjectIndex];
                json.Append('{');
                Property(json, "subjectId", subject.SubjectId, true);
                Property(json, "logicalSubjectKey", subject.LogicalSubjectKey, false);
                Property(json, "aspect", subject.Aspect.ToString(), false);
                BooleanProperty(json, "isReverted", subject.IsReverted, false);
                BooleanProperty(json, "hasConflict", subject.HasConflict, false);
                BooleanProperty(json, "hasBranchOnlyChange", subject.HasBranchOnlyChange, false);
                json.Append(",\"candidateStates\":[");
                for (int stateIndex = 0; stateIndex < subject.CandidateStates.Count; stateIndex++)
                {
                    if (stateIndex > 0) json.Append(',');
                    HistoryCandidateState state = subject.CandidateStates[stateIndex];
                    json.Append('{');
                    Property(json, "stateId", state.StateId, true);
                    json.Append(",\"revisionIds\":[");
                    AppendStringArrayValues(json, state.RevisionIds);
                    json.Append("]}");
                }
                json.Append("],\"observations\":[");
                for (int observationIndex = 0; observationIndex < subject.Observations.Count; observationIndex++)
                {
                    if (observationIndex > 0) json.Append(',');
                    HistoryObservation observation = subject.Observations[observationIndex];
                    json.Append('{');
                    Property(json, "revisionId", observation.RevisionId, true);
                    NumberProperty(json, "revisionOrder", observation.RevisionOrder, false);
                    Property(json, "stateId", observation.StateId, false);
                    BooleanProperty(json, "isPhysicallyPresent", observation.IsPhysicallyPresent, false);
                    Property(json, "exactText", observation.ExactText ?? "", false);
                    json.Append(",\"exactValue\":");
                    AppendCanonicalValue(json, observation.ExactValue);
                    json.Append('}');
                }
                json.Append("],\"transitions\":[");
                for (int transitionIndex = 0; transitionIndex < subject.Transitions.Count; transitionIndex++)
                {
                    if (transitionIndex > 0) json.Append(',');
                    HistoryTransition transition = subject.Transitions[transitionIndex];
                    json.Append('{');
                    Property(json, "transitionId", transition.TransitionId, true);
                    Property(json, "edgeId", transition.EdgeId, false);
                    Property(json, "beforeRevisionId", transition.BeforeRevisionId, false);
                    Property(json, "afterRevisionId", transition.AfterRevisionId, false);
                    Property(json, "beforeStateId", transition.BeforeStateId, false);
                    Property(json, "afterStateId", transition.AfterStateId, false);
                    Property(json, "exactKind", transition.ExactKind.ToString(), false);
                    Property(json, "equivalenceTags", transition.EquivalenceTags.ToString(), false);
                    BooleanProperty(json, "isOnMasterLineage", transition.IsOnMasterLineage, false);
                    BooleanProperty(json, "isRevert", transition.IsRevert, false);
                    BooleanProperty(json, "isConflict", transition.IsConflict, false);
                    BooleanProperty(json, "isBranchOnly", transition.IsBranchOnly, false);
                    Property(json, "groupId", transition.GroupId, false);
                    json.Append('}');
                }
                json.Append("]}");
            }
            json.Append("]}");
        }

        private static void AppendCanonicalValue(StringBuilder json, CanonicalValue value)
        {
            if (value == null)
            {
                json.Append("null");
                return;
            }
            json.Append('{');
            Property(json, "kind", value.Kind.ToString(), true);
            Property(json, "exactText", value.ExactText, false);
            Property(json, "dateSystem", value.DateSystem.ToString(), false);
            json.Append('}');
        }

        private static void AppendLineageConfidence(StringBuilder json, LineageConfidence confidence)
        {
            json.Append('{');
            Property(json, "level", confidence.Level.ToString(), true);
            DoubleProperty(json, "score", confidence.Score, false);
            BooleanProperty(json, "requiresConfirmation", confidence.RequiresConfirmation, false);
            json.Append('}');
        }

        private static void AppendLineageReasons(StringBuilder json, IEnumerable<LineageReason> reasons)
        {
            json.Append('[');
            int index = 0;
            foreach (LineageReason reason in reasons ?? Array.Empty<LineageReason>())
            {
                if (index++ > 0) json.Append(',');
                json.Append('{');
                Property(json, "code", reason.Code.ToString(), true);
                DoubleProperty(json, "confidence", reason.Confidence.Score, false);
                Property(json, "message", reason.Message, false);
                json.Append(",\"relatedRevisionIds\":[");
                AppendStringArrayValues(json, reason.RelatedRevisionIds);
                json.Append("]}");
            }
            json.Append(']');
        }

        private static void AppendStringArrayValues(StringBuilder json, IEnumerable<string> values)
        {
            int index = 0;
            foreach (string value in values ?? Array.Empty<string>())
            {
                if (index++ > 0) json.Append(',');
                AppendString(json, value ?? "");
            }
        }

        private static void AppendNumberArrayValues(StringBuilder json, IEnumerable<int> values)
        {
            int index = 0;
            foreach (int value in values ?? Array.Empty<int>())
            {
                if (index++ > 0) json.Append(',');
                json.Append(value.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static string FormatTimestamp(DateTimeOffset? value) =>
            value.HasValue ? value.Value.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture) : "";

        private static void Property(StringBuilder json, string name, string value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(':');
            AppendString(json, value ?? "");
        }

        private static void BooleanProperty(StringBuilder json, string name, bool value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(value ? ":true" : ":false");
        }

        private static void NumberProperty(StringBuilder json, string name, long value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(':').Append(value.ToString(CultureInfo.InvariantCulture));
        }

        private static void NullableNumberProperty(
            StringBuilder json,
            string name,
            int? value,
            bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(':');
            if (value.HasValue) json.Append(value.Value.ToString(CultureInfo.InvariantCulture));
            else json.Append("null");
        }

        private static void DoubleProperty(StringBuilder json, string name, double value, bool first)
        {
            if (!first) json.Append(',');
            AppendString(json, name);
            json.Append(':').Append(value.ToString("R", CultureInfo.InvariantCulture));
        }

        private static void AppendString(StringBuilder json, string value)
        {
            json.Append('"');
            value = value ?? "";
            for (int index = 0; index < value.Length; index++)
            {
                char character = value[index];
                switch (character)
                {
                    case '"': json.Append("\\\""); break;
                    case '\\': json.Append("\\\\"); break;
                    case '\b': json.Append("\\b"); break;
                    case '\f': json.Append("\\f"); break;
                    case '\n': json.Append("\\n"); break;
                    case '\r': json.Append("\\r"); break;
                    case '\t': json.Append("\\t"); break;
                    default:
                        if (character < 0x20 || character == '\u2028' || character == '\u2029' ||
                            char.IsSurrogate(character))
                            json.Append("\\u").Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));
                        else json.Append(character);
                        break;
                }
            }
            json.Append('"');
        }
    }
}
