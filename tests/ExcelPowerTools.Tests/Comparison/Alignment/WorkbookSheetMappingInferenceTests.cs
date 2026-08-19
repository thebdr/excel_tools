using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison.Alignment
{
    public sealed class WorkbookSheetMappingInferenceTests
    {
        [Fact]
        public void Unique_exact_names_map_strongly_without_requiring_matching_fingerprints()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before-data", "Data", structure: "old-structure", content: "old-content") },
                new[] { Sheet("after-data", "data", structure: "new-structure", content: "new-content") });

            ResolvedSheetMapping mapping = Assert.Single(result.Mappings);
            Assert.Equal(SheetMappingKind.ExactName, mapping.Kind);
            Assert.Equal(1, mapping.Confidence);
            Assert.Equal(SheetMappingReasonCodes.ExactName, Assert.Single(mapping.ReasonCodes));
            Assert.False(result.RequiresUserMapping);
            Assert.Empty(result.UnmatchedAfterSheets);
        }

        [Fact]
        public void Unique_rename_with_exact_structure_and_content_is_inferred()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before", "Old Name", structure: "structure-1", content: "content-1") },
                new[] { Sheet("after", "New Name", structure: "structure-1", content: "content-1") });

            ResolvedSheetMapping mapping = Assert.Single(result.Mappings);
            Assert.Equal(SheetMappingKind.InferredRename, mapping.Kind);
            Assert.Equal("Old Name", mapping.Before.Name);
            Assert.Equal("New Name", mapping.After.Name);
            Assert.Equal(0.65, mapping.Confidence, 6);
            Assert.Contains(SheetMappingReasonCodes.StructureFingerprint, mapping.ReasonCodes);
            Assert.Contains(SheetMappingReasonCodes.ContentFingerprint, mapping.ReasonCodes);
            Assert.Empty(result.UnresolvedBeforeSheets);
            Assert.Empty(result.UnmatchedAfterSheets);
        }

        [Fact]
        public void Matching_code_name_alone_never_proves_a_rename()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before", "Old", codeName: "StableCode") },
                new[] { Sheet("after", "New", codeName: "StableCode") });

            Assert.Empty(result.Mappings);
            UnresolvedSheetMapping unresolved = Assert.Single(result.UnresolvedBeforeSheets);
            Assert.Empty(unresolved.Candidates);
            Assert.Contains(SheetMappingReasonCodes.NoSupportedRenameEvidence, unresolved.ReasonCodes);
            Assert.Equal("after", Assert.Single(result.UnmatchedAfterSheets).SheetKey);
        }

        [Theory]
        [InlineData("shared-structure", null)]
        [InlineData(null, "shared-content")]
        public void Rename_requires_both_structural_and_content_evidence(
            string? structure,
            string? content)
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before", "Old", structure: structure, content: content) },
                new[] { Sheet("after", "New", structure: structure, content: content) });

            Assert.Empty(result.Mappings);
            UnresolvedSheetMapping unresolved = Assert.Single(result.UnresolvedBeforeSheets);
            Assert.Empty(unresolved.Candidates);
            Assert.Contains(SheetMappingReasonCodes.NoSupportedRenameEvidence, unresolved.ReasonCodes);
        }

        [Fact]
        public void Weak_supported_candidate_is_exposed_with_confidence_and_reasons_but_not_chosen()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before", "Old", structure: "layout", header: "header") },
                new[] { Sheet("after", "New", structure: "layout", header: "header") });

            Assert.Empty(result.Mappings);
            UnresolvedSheetMapping unresolved = Assert.Single(result.UnresolvedBeforeSheets);
            SheetMappingCandidate candidate = Assert.Single(unresolved.Candidates);
            Assert.Equal("after", candidate.After.SheetKey);
            Assert.Equal(0.42, candidate.Confidence, 6);
            Assert.Equal(
                new[]
                {
                    SheetMappingReasonCodes.StructureFingerprint,
                    SheetMappingReasonCodes.HeaderFingerprint,
                },
                candidate.ReasonCodes);
            Assert.Contains(SheetMappingReasonCodes.BelowRenameThreshold, unresolved.ReasonCodes);
        }

        [Fact]
        public void Equal_target_candidates_remain_ambiguous_and_ranked_deterministically()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before", "Original", structure: "layout", content: "cells") },
                new[]
                {
                    Sheet("target-b", "Renamed B", structure: "layout", content: "cells"),
                    Sheet("target-a", "Renamed A", structure: "layout", content: "cells"),
                });

            Assert.Empty(result.Mappings);
            UnresolvedSheetMapping unresolved = Assert.Single(result.UnresolvedBeforeSheets);
            Assert.Equal(new[] { "target-a", "target-b" }, unresolved.Candidates.Select(c => c.After.SheetKey));
            Assert.All(unresolved.Candidates, candidate => Assert.Equal(0.65, candidate.Confidence, 6));
            Assert.Contains(SheetMappingReasonCodes.CompetingTargets, unresolved.ReasonCodes);
            Assert.Equal(2, result.UnmatchedAfterSheets.Count);
        }

        [Fact]
        public void Target_shared_by_equal_sources_remains_unresolved_in_both_directions()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[]
                {
                    Sheet("source-a", "Old A", structure: "layout", content: "cells"),
                    Sheet("source-b", "Old B", structure: "layout", content: "cells"),
                },
                new[] { Sheet("target", "New", structure: "layout", content: "cells") });

            Assert.Empty(result.Mappings);
            Assert.Equal(2, result.UnresolvedBeforeSheets.Count);
            Assert.All(result.UnresolvedBeforeSheets, unresolved =>
                Assert.Contains(SheetMappingReasonCodes.CompetingSources, unresolved.ReasonCodes));
            Assert.Single(result.UnmatchedAfterSheets);
        }

        [Fact]
        public void Known_incompatible_sheet_kinds_are_not_rename_candidates()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[]
                {
                    Sheet("before", "Old", WorkbookSheetKind.Worksheet,
                        structure: "layout", content: "cells")
                },
                new[]
                {
                    Sheet("after", "New", WorkbookSheetKind.ChartSheet,
                        structure: "layout", content: "cells")
                });

            Assert.Empty(result.Mappings);
            Assert.Empty(Assert.Single(result.UnresolvedBeforeSheets).Candidates);
        }

        [Fact]
        public void Exact_name_mapping_and_rename_inference_preserve_one_to_one_targets()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[]
                {
                    Sheet("before-data", "Data", structure: "layout-a", content: "cells-a"),
                    Sheet("before-old", "Old", structure: "layout-b", content: "cells-b"),
                },
                new[]
                {
                    Sheet("after-data", "Data", structure: "layout-b", content: "cells-b"),
                    Sheet("after-new", "New", structure: "layout-b", content: "cells-b"),
                });

            Assert.Collection(
                result.Mappings,
                mapping =>
                {
                    Assert.Equal("before-data", mapping.Before.SheetKey);
                    Assert.Equal("after-data", mapping.After.SheetKey);
                    Assert.Equal(SheetMappingKind.ExactName, mapping.Kind);
                },
                mapping =>
                {
                    Assert.Equal("before-old", mapping.Before.SheetKey);
                    Assert.Equal("after-new", mapping.After.SheetKey);
                    Assert.Equal(SheetMappingKind.InferredRename, mapping.Kind);
                });
        }

        [Fact]
        public void Results_and_nested_reason_and_candidate_collections_are_immutable_copies()
        {
            var before = new List<WorkbookSheetFingerprint>
            {
                Sheet("before", "Old", structure: "layout", content: "cells")
            };
            var after = new List<WorkbookSheetFingerprint>
            {
                Sheet("after-b", "New B", structure: "layout", content: "cells"),
                Sheet("after-a", "New A", structure: "layout", content: "cells"),
            };

            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(before, after);
            before.Clear();
            after.Clear();

            UnresolvedSheetMapping unresolved = Assert.Single(result.UnresolvedBeforeSheets);
            SheetMappingCandidate candidate = unresolved.Candidates[0];
            Assert.Throws<NotSupportedException>(() =>
                ((IList<UnresolvedSheetMapping>)result.UnresolvedBeforeSheets).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<SheetMappingCandidate>)unresolved.Candidates).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<string>)candidate.ReasonCodes).Clear());
            Assert.Equal(2, unresolved.Candidates.Count);
        }

        [Fact]
        public void Duplicate_workbook_local_sheet_keys_are_rejected()
        {
            Assert.Throws<ArgumentException>(() => WorkbookSheetMappingInference.Infer(
                new[] { Sheet("duplicate", "A"), Sheet("duplicate", "B") },
                Array.Empty<WorkbookSheetFingerprint>()));
        }

        [Fact]
        public void Missing_names_are_preserved_for_diagnostics_but_never_used_as_exact_identity()
        {
            WorkbookSheetMappingResult result = WorkbookSheetMappingInference.Infer(
                new[] { Sheet("before", "") },
                new[] { Sheet("after", "") });

            Assert.Empty(result.Mappings);
            Assert.Equal("", Assert.Single(result.UnresolvedBeforeSheets).Before.Name);
            Assert.Equal("", Assert.Single(result.UnmatchedAfterSheets).Name);
        }

        private static WorkbookSheetFingerprint Sheet(
            string key,
            string name,
            WorkbookSheetKind kind = WorkbookSheetKind.Worksheet,
            string? codeName = null,
            string? structure = null,
            string? content = null,
            string? header = null,
            string? formula = null,
            int rows = 0,
            int columns = 0,
            long cells = 0,
            string? addInIdentity = null) =>
            new WorkbookSheetFingerprint(
                key,
                name,
                kind,
                codeName,
                addInIdentity,
                structure,
                content,
                header,
                formula,
                rows,
                columns,
                cells);
    }
}
