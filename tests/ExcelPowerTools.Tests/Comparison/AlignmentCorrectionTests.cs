using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Project;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class AlignmentCorrectionTests
    {
        private static readonly ReviewerSnapshot Reviewer = new ReviewerSnapshot(
            "Confirmed Reviewer",
            "Office Reviewer",
            new[] { "explicit-test" });
        private static readonly DateTimeOffset When = new DateTimeOffset(
            2026, 7, 17, 8, 9, 10, TimeSpan.Zero);

        [Fact]
        public void Generations_are_append_only_deterministic_replace_same_scope_and_reject_collisions()
        {
            AlignmentCorrectionRequest sheet = AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", "before/sheet", "after/sheet", Reviewer, When, "pair");
            AlignmentCorrectionRequest header = AlignmentCorrectionRequest.ConfirmHeaderRow(
                "before", "before/sheet", 1, 2, 4, 5, Reviewer, When, "header and data");
            AlignmentCorrectionRequest key = AlignmentCorrectionRequest.ConfirmKeyColumns(
                "before",
                "after",
                "before/sheet",
                "after/sheet",
                new[]
                {
                    new AlignmentGridIndexReference(0, 1),
                    new AlignmentGridIndexReference(1, 2),
                    new AlignmentGridIndexReference(2, 3),
                },
                new[]
                {
                    new AlignmentGridIndexReference(0, 1),
                    new AlignmentGridIndexReference(1, 2),
                    new AlignmentGridIndexReference(2, 3),
                },
                Reviewer,
                When,
                "three-part domain key");

            AlignmentCorrectionGeneration first = ApplyAll("mapping-generation", sheet, header, key);
            AlignmentCorrectionGeneration second = ApplyAll("mapping-generation", sheet, header, key);

            Assert.Equal(first.GenerationId, second.GenerationId);
            Assert.Equal(first.Events.Select(item => item.EventId), second.Events.Select(item => item.EventId));
            Assert.Equal(3, first.Events.Count);
            Assert.Equal(first.Events[1].PriorGenerationId, first.Events[0].NewGenerationId);
            Assert.Equal(first.GenerationId, first.Events[2].NewGenerationId);
            Assert.Equal(3, Assert.Single(first.KeyColumns).BeforeColumns.Count);
            Assert.Equal(4, Assert.Single(first.HeaderRows).FirstDataRow!.TableIndex);

            AlignmentCorrectionGeneration root = AlignmentCorrectionEngine.Create("mapping-generation");
            AlignmentCorrectionGeneration withSheet = AlignmentCorrectionEngine.Apply(root, sheet);
            Assert.Throws<InvalidOperationException>(() => AlignmentCorrectionEngine.Apply(withSheet, sheet));
            AlignmentCorrectionGeneration replacedSheet = AlignmentCorrectionEngine.Apply(
                withSheet,
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    "before", "after", "before/sheet", "after/other", Reviewer, When));
            Assert.Equal("after/other", Assert.Single(replacedSheet.SheetPairs).AfterSheetKey);
            Assert.Equal(sheet.SheetPair!.ConfirmationId, replacedSheet.Events.Last().SupersededConfirmationId);
            AlignmentCorrectionGeneration twoSources = AlignmentCorrectionEngine.Apply(
                root,
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    "before", "after", "before/first", "after/claimed", Reviewer, When));
            Assert.Throws<InvalidOperationException>(() => AlignmentCorrectionEngine.Apply(
                twoSources,
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    "before", "after", "before/second", "after/claimed", Reviewer, When)));

            AlignmentCorrectionGeneration withHeader = AlignmentCorrectionEngine.Apply(root, header);
            AlignmentCorrectionGeneration replacedHeader = AlignmentCorrectionEngine.Apply(
                withHeader,
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "before", "before/sheet", 2, 3, Reviewer, When));
            Assert.Equal(2, Assert.Single(replacedHeader.HeaderRows).Row!.TableIndex);
            Assert.Equal(header.HeaderRow!.ConfirmationId, replacedHeader.Events.Last().SupersededConfirmationId);

            AlignmentCorrectionGeneration withKey = AlignmentCorrectionEngine.Apply(root, key);
            AlignmentCorrectionGeneration replacedKey = AlignmentCorrectionEngine.Apply(
                withKey,
                AlignmentCorrectionRequest.ConfirmKeyColumns(
                    "before",
                    "after",
                    "before/sheet",
                    "after/sheet",
                    new[] { new AlignmentGridIndexReference(3, 4) },
                    new[] { new AlignmentGridIndexReference(3, 4) },
                    Reviewer,
                    When));
            Assert.Single(Assert.Single(replacedKey.KeyColumns).BeforeColumns);
            Assert.Equal(key.KeyColumns!.ConfirmationId, replacedKey.Events.Last().SupersededConfirmationId);
            Assert.Throws<ArgumentException>(() => AlignmentCorrectionRequest.ConfirmKeyColumns(
                "before",
                "after",
                "before/sheet",
                "after/sheet",
                new[]
                {
                    new AlignmentGridIndexReference(0, 1),
                    new AlignmentGridIndexReference(0, 2),
                },
                new[]
                {
                    new AlignmentGridIndexReference(0, 1),
                    new AlignmentGridIndexReference(1, 2),
                },
                Reviewer,
                When));
        }

        [Fact]
        public void Deleted_new_confirmation_is_mutually_exclusive_with_a_pair_and_can_return_to_auto()
        {
            AlignmentCorrectionGeneration root = AlignmentCorrectionEngine.Create("mapping-generation");
            AlignmentCorrectionGeneration unmatched = AlignmentCorrectionEngine.Apply(
                root,
                AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                    "before", "after", "before/deleted", Reviewer, When, "not a rename"));

            AlignmentSheetUnmatchedConfirmation active = Assert.Single(unmatched.UnmatchedSheets);
            Assert.Equal("before/deleted", active.BeforeSheetKey);
            Assert.Empty(unmatched.SheetPairs);
            Assert.Equal(AlignmentCorrectionKind.ConfirmSheetUnmatched, unmatched.Events.Last().Kind);
            Assert.Throws<InvalidOperationException>(() => AlignmentCorrectionEngine.Apply(
                unmatched,
                AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                    "before", "after", "before/deleted", Reviewer, When)));

            AlignmentCorrectionGeneration paired = AlignmentCorrectionEngine.Apply(
                unmatched,
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    "before", "after", "before/deleted", "after/new", Reviewer, When));
            Assert.Empty(paired.UnmatchedSheets);
            Assert.Equal("after/new", Assert.Single(paired.SheetPairs).AfterSheetKey);
            Assert.Equal(active.ConfirmationId, paired.Events.Last().SupersededConfirmationId);

            AlignmentCorrectionGeneration restoredUnmatched = AlignmentCorrectionEngine.Apply(
                paired,
                AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                    "before", "after", "before/deleted", Reviewer, When.AddMinutes(1)));
            Assert.Empty(restoredUnmatched.SheetPairs);
            Assert.Single(restoredUnmatched.UnmatchedSheets);
            Assert.Equal(
                paired.Events.Last().ConfirmationId,
                restoredUnmatched.Events.Last().SupersededConfirmationId);

            AlignmentCorrectionGeneration automatic = AlignmentCorrectionEngine.Apply(
                restoredUnmatched,
                AlignmentCorrectionRequest.ClearSheetUnmatched(
                    "before", "after", "before/deleted", Reviewer, When.AddMinutes(2)));
            Assert.Empty(automatic.UnmatchedSheets);
            Assert.Equal(
                restoredUnmatched.UnmatchedSheets.Single().ConfirmationId,
                automatic.Events.Last().SupersededConfirmationId);
        }

        [Fact]
        public void Confirmed_deleted_and_new_sheets_complete_the_edge_without_comparing_them_as_a_rename()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Legacy", "LegacyCode",
                Text("A1", "Old ID"), Text("B1", "Old value"),
                Text("A2", "A"), Text("B2", "one")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Replacement", "ReplacementCode",
                Text("D5", "New ID"), Text("E5", "New amount"),
                Text("D6", "Z"), Text("E6", "900")));
            string deletedKey = SheetKey(before, "Legacy");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-1"),
                AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                    "before", "after", deletedKey, Reviewer, When, "deleted in revision 1.2"));

            OpenXmlComparisonProjectResult result = Build(
                "project-deleted-new",
                "mapping-1",
                before,
                after,
                corrections);

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            WorkbookSheetMappingResult mapping = Assert.Single(result.Edges).SheetMapping.Mapping;
            Assert.Empty(mapping.Mappings);
            Assert.Empty(mapping.UnresolvedBeforeSheets);
            Assert.Equal(deletedKey, Assert.Single(mapping.UnmatchedBeforeSheets).SheetKey);
            Assert.Equal("Replacement", Assert.Single(mapping.UnmatchedAfterSheets).Name);
            Assert.Single(result.Revisions.Single(item => item.Source.RevisionId == "after").Sheets);
        }

        [Fact]
        public void Clear_requests_return_exact_scopes_to_auto_and_keep_append_only_swap_history()
        {
            AlignmentCorrectionRequest pairA = AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", "before/A", "after/X", Reviewer, When);
            AlignmentCorrectionRequest pairB = AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", "before/B", "after/Y", Reviewer, When);
            AlignmentCorrectionRequest clearB = AlignmentCorrectionRequest.ClearSheetPair(
                "before", "after", "before/B", Reviewer, When, "free target Y");
            AlignmentCorrectionRequest moveA = AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", "before/A", "after/Y", Reviewer, When);
            AlignmentCorrectionRequest moveB = AlignmentCorrectionRequest.ConfirmSheetPair(
                "before", "after", "before/B", "after/X", Reviewer, When);
            AlignmentCorrectionRequest header = AlignmentCorrectionRequest.ConfirmNoHeader(
                "before", "before/A", Reviewer, When);
            AlignmentCorrectionRequest clearHeader = AlignmentCorrectionRequest.ClearHeaderRow(
                "before", "before/A", Reviewer, When, "return to detector");
            AlignmentCorrectionRequest key = AlignmentCorrectionRequest.ConfirmKeyColumns(
                "before",
                "after",
                "before/A",
                "after/Y",
                new[] { new AlignmentGridIndexReference(0, 1) },
                new[] { new AlignmentGridIndexReference(0, 1) },
                Reviewer,
                When);
            AlignmentCorrectionRequest clearKey = AlignmentCorrectionRequest.ClearKeyColumns(
                "before", "after", "before/A", "after/Y", Reviewer, When, "return to inference");

            var requests = new[]
            {
                pairA,
                pairB,
                clearB,
                moveA,
                moveB,
                header,
                clearHeader,
                key,
                clearKey,
            };
            AlignmentCorrectionGeneration result = ApplyAll("mapping-generation", requests);
            AlignmentCorrectionGeneration replay = ApplyAll("mapping-generation", requests);

            Assert.Equal(result.GenerationId, replay.GenerationId);
            Assert.Equal(result.Events.Select(item => item.EventId), replay.Events.Select(item => item.EventId));
            Assert.Equal(9, result.Events.Count);
            Assert.Collection(
                result.SheetPairs.OrderBy(item => item.BeforeSheetKey, StringComparer.Ordinal),
                item =>
                {
                    Assert.Equal("before/A", item.BeforeSheetKey);
                    Assert.Equal("after/Y", item.AfterSheetKey);
                },
                item =>
                {
                    Assert.Equal("before/B", item.BeforeSheetKey);
                    Assert.Equal("after/X", item.AfterSheetKey);
                });
            Assert.Empty(result.HeaderRows);
            Assert.Empty(result.KeyColumns);

            AlignmentCorrectionEvent clearedPair = result.Events[2];
            Assert.Equal(AlignmentCorrectionKind.ClearSheetPair, clearedPair.Kind);
            Assert.Equal(clearB.RequestId, clearedPair.RequestId);
            Assert.Equal("", clearedPair.ConfirmationId);
            Assert.Equal(pairB.SheetPair!.ConfirmationId, clearedPair.SupersededConfirmationId);
            Assert.Equal("before", clearedPair.Scope.ParentRevisionId);
            Assert.Equal("after", clearedPair.Scope.ChildRevisionId);
            Assert.Equal("before/B", clearedPair.Scope.BeforeSheetKey);
            Assert.Equal("", clearedPair.Scope.AfterSheetKey);
            Assert.Null(clearedPair.SheetPair);
            Assert.Null(clearedPair.HeaderRow);
            Assert.Null(clearedPair.KeyColumns);

            Assert.Equal(AlignmentCorrectionKind.ClearHeaderRow, result.Events[6].Kind);
            Assert.Equal(header.HeaderRow!.ConfirmationId, result.Events[6].SupersededConfirmationId);
            Assert.Equal("before", result.Events[6].Scope.RevisionId);
            Assert.Equal("before/A", result.Events[6].Scope.PhysicalSheetKey);
            Assert.Equal(AlignmentCorrectionKind.ClearKeyColumns, result.Events[8].Kind);
            Assert.Equal(key.KeyColumns!.ConfirmationId, result.Events[8].SupersededConfirmationId);
            Assert.Equal("after/Y", result.Events[8].Scope.AfterSheetKey);

            AlignmentCorrectionGeneration root = AlignmentCorrectionEngine.Create("mapping-generation");
            Assert.Throws<InvalidOperationException>(() => AlignmentCorrectionEngine.Apply(
                root,
                AlignmentCorrectionRequest.ClearSheetPair(
                    "before", "after", "before/missing", Reviewer, When)));
            Assert.Throws<InvalidOperationException>(() => AlignmentCorrectionEngine.Apply(
                root,
                AlignmentCorrectionRequest.ClearHeaderRow(
                    "before", "before/missing", Reviewer, When)));
            Assert.Throws<InvalidOperationException>(() => AlignmentCorrectionEngine.Apply(
                root,
                AlignmentCorrectionRequest.ClearKeyColumns(
                    "before", "after", "before/missing", "after/missing", Reviewer, When)));
        }

        [Fact]
        public void Confirmed_sheet_pair_resolves_ambiguity_and_participates_in_effective_mapping_identity()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Original", "StableCode",
                Text("A1", "ID"), Text("B1", "Value"), Text("A2", "A"), Text("B2", "one")));
            OpenXmlWorkbookSnapshot after = Workbook("after",
                Sheet("Candidate A", "StableCode",
                    Text("A1", "ID"), Text("B1", "Value"), Text("A2", "A"), Text("B2", "one")),
                Sheet("Candidate B", "StableCode",
                    Text("A1", "ID"), Text("B1", "Value"), Text("A2", "A"), Text("B2", "one")));
            string beforeKey = SheetKey(before, "Original");
            string afterKey = SheetKey(after, "Candidate A");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-1"),
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    "before", "after", beforeKey, afterKey, Reviewer, When));

            OpenXmlComparisonProjectResult result = Build(
                "project-confirmed-sheet",
                "mapping-1",
                before,
                after,
                corrections);

            Assert.True(result.IsComplete);
            Assert.Equal(corrections.GenerationId, result.AlignmentCorrectionGenerationId);
            Assert.NotEqual(result.MappingGenerationId, result.EffectiveMappingGenerationId);
            Assert.Equal(result.EffectiveMappingGenerationId, result.History!.MappingGenerationId);
            ResolvedSheetMapping mapping = Assert.Single(
                Assert.Single(result.Edges).SheetMapping.Mapping.Mappings,
                item => item.Kind == SheetMappingKind.UserConfirmed);
            Assert.Equal(beforeKey, mapping.Before.SheetKey);
            Assert.Equal(afterKey, mapping.After.SheetKey);
            Assert.Contains(SheetMappingReasonCodes.UserConfirmed, mapping.ReasonCodes);
        }

        [Fact]
        public void Cleared_header_uses_auto_detection_but_the_audit_generation_remains_in_result_identity()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "1"),
                Text("A3", "B"), Text("B3", "2")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "1"),
                Text("A3", "B"), Text("B3", "3")));
            string beforeKey = SheetKey(before, "Data");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmNoHeader(
                    "before", beforeKey, Reviewer, When, "temporary review choice"));
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ClearHeaderRow(
                    "before", beforeKey, Reviewer, When, "use detector again"));

            OpenXmlComparisonProjectResult result = Build(
                "project-cleared-header",
                "mapping-1",
                before,
                after,
                corrections);

            OpenXmlProjectSheetAlignment alignment = Assert.Single(Assert.Single(result.Edges).SheetAlignments);
            Assert.NotEqual(HeaderDecision.ConfirmedHeader, alignment.BeforeHeader!.Decision);
            Assert.NotEqual(HeaderDecision.ConfirmedNoHeader, alignment.BeforeHeader.Decision);
            Assert.Empty(corrections.HeaderRows);
            Assert.Equal(2, corrections.Events.Count);
            Assert.Equal(corrections.GenerationId, result.AlignmentCorrectionGenerationId);
            Assert.NotEqual(result.MappingGenerationId, result.EffectiveMappingGenerationId);
        }

        [Fact]
        public void Confirmed_nonleading_header_data_start_and_three_column_key_preserve_predata_changes()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "draft note"), Text("B1", "metadata"),
                Text("A2", "Group"), Text("B2", "Item"), Text("C2", "Channel"), Text("D2", "Value"),
                Text("A3", "unit"), Text("B3", "unit"), Text("C3", "unit"), Text("D3", "unit"),
                Text("A4", "---"), Text("B4", "---"), Text("C4", "---"), Text("D4", "---"),
                Text("A5", "G1"), Text("B5", "I1"), Text("C5", "C1"), Text("D5", "10"),
                Text("A6", "G1"), Text("B6", "I1"), Text("C6", "C2"), Text("D6", "20")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "reviewed note"), Text("B1", "metadata"),
                Text("A2", "Group"), Text("B2", "Item"), Text("C2", "Channel"), Text("D2", "Value"),
                Text("A3", "unit"), Text("B3", "unit"), Text("C3", "unit"), Text("D3", "unit"),
                Text("A4", "---"), Text("B4", "---"), Text("C4", "---"), Text("D4", "---"),
                Text("A5", "G1"), Text("B5", "I1"), Text("C5", "C1"), Text("D5", "10"),
                Text("A6", "G1"), Text("B6", "I1"), Text("C6", "C2"), Text("D6", "25")));
            string beforeKey = SheetKey(before, "Data");
            string afterKey = SheetKey(after, "Data");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "before", beforeKey, 1, 2, 4, 5, Reviewer, When));
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "after", afterKey, 1, 2, 4, 5, Reviewer, When));
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmKeyColumns(
                    "before",
                    "after",
                    beforeKey,
                    afterKey,
                    new[]
                    {
                        new AlignmentGridIndexReference(0, 1),
                        new AlignmentGridIndexReference(1, 2),
                        new AlignmentGridIndexReference(2, 3),
                    },
                    new[]
                    {
                        new AlignmentGridIndexReference(0, 1),
                        new AlignmentGridIndexReference(1, 2),
                        new AlignmentGridIndexReference(2, 3),
                    },
                    Reviewer,
                    When));

            OpenXmlComparisonProjectResult result = Build(
                "project-confirmed-region",
                "mapping-1",
                before,
                after,
                corrections);

            Assert.True(result.IsComplete);
            OpenXmlProjectSheetAlignment alignment = Assert.Single(Assert.Single(result.Edges).SheetAlignments);
            Assert.Equal(HeaderDecision.ConfirmedHeader, alignment.BeforeHeader!.Decision);
            Assert.Equal(4, alignment.BeforeHeader.DataStartRowIndex);
            Assert.Equal(KeyInferenceDecision.ConfirmedKey, alignment.KeyInference!.Decision);
            Assert.Equal(3, alignment.KeyInference.SelectedCandidate!.BeforeColumns.Count);
            Assert.All(alignment.RowAlignment!.Links.Where(item => item.BeforeIndex < 4), item =>
                Assert.Equal(AlignmentConfidence.Confirmed, item.Confidence));

            OpenXmlAlignedSheetRevision rootSheet = result.Revisions.Single(item =>
                item.Source.RevisionId == "before").Sheets.Single();
            string preambleSubject = StableId.Compute(
                "aligned-cell-subject-v1",
                rootSheet.LogicalSheetId,
                rootSheet.Rows.IdentityByPhysicalIndex[0],
                rootSheet.Columns.IdentityByPhysicalIndex[0]);
            var preambleHistory = Assert.Single(result.History!.Subjects, item =>
                item.LogicalSubjectKey == preambleSubject && item.Aspect == CellChangeAspect.Value);
            Assert.Equal("draft note", preambleHistory.Observations[0].ExactValue!.ExactText);
            Assert.Equal("reviewed note", preambleHistory.Observations[1].ExactValue!.ExactText);
        }

        [Fact]
        public void Unequal_confirmed_predata_shapes_remain_correctable_instead_of_being_guessed()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "note"), Text("A2", "ID"), Text("B2", "Value"),
                Text("A3", "unit"), Text("B3", "unit"), Text("A4", "---"), Text("B4", "---"),
                Text("A5", "A"), Text("B5", "1")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "note"), Text("A2", "ID"), Text("B2", "Value"),
                Text("A3", "unit"), Text("B3", "unit"), Text("A4", "A"), Text("B4", "1")));
            string beforeKey = SheetKey(before, "Data");
            string afterKey = SheetKey(after, "Data");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "before", beforeKey, 1, 2, 4, 5, Reviewer, When));
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "after", afterKey, 1, 2, 3, 4, Reviewer, When));

            OpenXmlComparisonProjectResult result = Build(
                "project-unequal-predata",
                "mapping-1",
                before,
                after,
                corrections);

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            Assert.Contains(result.Findings, item => item.Code == OpenXmlProjectFindingCodes.NonLeadingHeader);
        }

        [Fact]
        public void Explicit_no_header_decision_is_audited_and_treats_every_row_as_data()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "1")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "2")));
            string beforeKey = SheetKey(before, "Data");
            string afterKey = SheetKey(after, "Data");
            AlignmentCorrectionGeneration corrections = AlignmentCorrectionEngine.Create("mapping-1");
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "before", beforeKey, 0, 1, Reviewer, When));
            string superseded = corrections.HeaderRows.Single().ConfirmationId;
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmNoHeader(
                    "before", beforeKey, Reviewer, When, "false positive"));
            corrections = AlignmentCorrectionEngine.Apply(corrections,
                AlignmentCorrectionRequest.ConfirmNoHeader(
                    "after", afterKey, Reviewer, When, "false positive"));

            OpenXmlComparisonProjectResult result = Build(
                "project-no-header",
                "mapping-1",
                before,
                after,
                corrections);

            OpenXmlProjectSheetAlignment alignment = Assert.Single(Assert.Single(result.Edges).SheetAlignments);
            Assert.Equal(HeaderDecision.ConfirmedNoHeader, alignment.BeforeHeader!.Decision);
            Assert.Equal(HeaderDecision.ConfirmedNoHeader, alignment.AfterHeader!.Decision);
            Assert.Equal(0, alignment.BeforeHeader.DataStartRowIndex);
            Assert.Equal(superseded, corrections.Events[1].SupersededConfirmationId);
            Assert.Equal(AlignmentHeaderConfirmationDecision.NoHeader,
                corrections.HeaderRows.Single(item => item.RevisionId == "before").Decision);
        }

        [Fact]
        public void Engine_rejects_stale_sheet_keys_indices_provenance_and_generation_scope()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("B3", "ID"), Text("C3", "Value"), Text("B4", "A"), Text("C4", "1")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("B3", "ID"), Text("C3", "Value"), Text("B4", "A"), Text("C4", "2")));
            string beforeKey = SheetKey(before, "Data");
            string afterKey = SheetKey(after, "Data");

            AlignmentCorrectionGeneration staleSheet = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-1"),
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    "before", "after", beforeKey + "/stale", afterKey, Reviewer, When));
            Assert.Throws<InvalidOperationException>(() => Build(
                "project-stale-sheet", "mapping-1", before, after, staleSheet));

            AlignmentCorrectionGeneration outOfRange = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-1"),
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "before", beforeKey, 99, 102, Reviewer, When));
            Assert.Throws<InvalidOperationException>(() => Build(
                "project-stale-row", "mapping-1", before, after, outOfRange));

            AlignmentCorrectionGeneration stalePhysical = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-1"),
                AlignmentCorrectionRequest.ConfirmHeaderRow(
                    "before", beforeKey, 0, 99, Reviewer, When));
            Assert.Throws<InvalidOperationException>(() => Build(
                "project-stale-physical", "mapping-1", before, after, stalePhysical));

            AlignmentCorrectionGeneration wrongGeneration = AlignmentCorrectionEngine.Create("mapping-other");
            Assert.Throws<InvalidOperationException>(() => Build(
                "project-stale-generation", "mapping-1", before, after, wrongGeneration));

            AlignmentCorrectionGeneration staleKey = AlignmentCorrectionEngine.Apply(
                AlignmentCorrectionEngine.Create("mapping-1"),
                AlignmentCorrectionRequest.ConfirmKeyColumns(
                    "before",
                    "after",
                    beforeKey,
                    afterKey,
                    new[] { new AlignmentGridIndexReference(5, 7) },
                    new[] { new AlignmentGridIndexReference(0, 2) },
                    Reviewer,
                    When));
            Assert.Throws<InvalidOperationException>(() => Build(
                "project-stale-key", "mapping-1", before, after, staleKey));
        }

        private static AlignmentCorrectionGeneration ApplyAll(
            string mappingGeneration,
            params AlignmentCorrectionRequest[] requests)
        {
            AlignmentCorrectionGeneration result = AlignmentCorrectionEngine.Create(mappingGeneration);
            foreach (AlignmentCorrectionRequest request in requests)
                result = AlignmentCorrectionEngine.Apply(result, request);
            return result;
        }

        private static OpenXmlComparisonProjectResult Build(
            string projectId,
            string mappingGeneration,
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after,
            AlignmentCorrectionGeneration corrections) =>
            OpenXmlComparisonProjectEngine.Build(
                projectId,
                mappingGeneration,
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                },
                new OpenXmlComparisonProjectOptions { AlignmentCorrections = corrections });

        private static string SheetKey(OpenXmlWorkbookSnapshot workbook, string name) =>
            OpenXmlSheetFingerprintBuilder.Build(workbook).Sheets.Single(item =>
                item.Fingerprint.Name == name).Fingerprint.SheetKey;

        private static OpenXmlWorkbookSnapshot Workbook(
            string revisionId,
            params OpenXmlWorkbookFixture.SheetDefinition[] sheets) =>
            OpenXmlWorkbookSnapshotAdapter.FromBytes(revisionId, OpenXmlWorkbookFixture.Create(sheets));

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(
            string name,
            string codeName,
            params Cell[] cells) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                CodeName = codeName,
                Cells = cells,
            };

        private static Cell Text(string reference, string value) =>
            new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new SpreadsheetText(value)),
            };
    }
}
