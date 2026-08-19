using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.Project;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public class OpenXmlComparisonProjectEngineTests
    {
        [Fact]
        public void Builds_linear_history_and_keeps_row_identity_through_an_insertion()
        {
            OpenXmlWorkbookSnapshot r1 = Workbook("r1", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot r2 = Workbook("r2", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "20"),
                Text("A5", "C"), Text("B5", "30")));
            OpenXmlWorkbookSnapshot r3 = Workbook("r3", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "25"),
                Text("A5", "C"), Text("B5", "30")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-linear",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("r1", "", 0, r1, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("r2", "r1", 1, r2),
                    new OpenXmlProjectRevision("r3", "r2", 2, r3, isMaster: true),
                });

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            Assert.Equal(2, result.Findings.Count(item =>
                item.Code == OpenXmlProjectFindingCodes.NonFormulaValueContentUnmodeled &&
                item.Severity == DiagnosticSeverity.Warning));
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.Equal(2, result.Edges.Count);

            OpenXmlAlignedSheetRevision r1Sheet = Revision(result, "r1").Sheets.Single();
            OpenXmlAlignedSheetRevision r2Sheet = Revision(result, "r2").Sheets.Single();
            OpenXmlAlignedSheetRevision r3Sheet = Revision(result, "r3").Sheets.Single();
            string logicalB = r1Sheet.Rows.IdentityByPhysicalIndex[2];
            Assert.Equal(logicalB, r2Sheet.Rows.IdentityByPhysicalIndex[3]);
            Assert.Equal(logicalB, r3Sheet.Rows.IdentityByPhysicalIndex[3]);
            Assert.True(result.History!.AtomicTransitionCount > 0);
        }

        [Fact]
        public void Exact_byte_difference_warns_without_blocking_zero_formula_value_changes()
        {
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[]
                {
                    Sheet("Data", "DataCode",
                        Text("A1", "ID"), Text("B1", "Amount"),
                        Text("A2", "A"), Text("B2", "10"),
                        Text("A3", "B"), Text("B3", "20")),
                }));
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[]
                {
                    Sheet("Data", "DataCode",
                        Text("A1", "ID"), Text("B1", "Amount"),
                        Text("A2", "A"), Text("B2", "10"),
                        Text("A3", "B"), Text("B3", "20")),
                }, includeDateStyle: true));

            OpenXmlComparisonProjectResult first = OpenXmlComparisonProjectEngine.Build(
                "project-package-warning",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });
            OpenXmlComparisonProjectResult second = OpenXmlComparisonProjectEngine.Build(
                "project-package-warning",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.NotEqual(before.ByteContentId, after.ByteContentId);
            Assert.True(first.IsComplete);
            Assert.True(Assert.Single(first.Edges).IsComplete);
            Assert.NotNull(first.History);
            Assert.Equal(0, first.History!.AtomicTransitionCount);
            OpenXmlProjectFinding warning = Assert.Single(first.Findings);
            Assert.Equal(OpenXmlProjectFindingCodes.NonFormulaValueContentUnmodeled, warning.Code);
            Assert.Equal(DiagnosticSeverity.Warning, warning.Severity);
            Assert.Contains("styles", warning.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("comments", warning.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(warning.FindingId, Assert.Single(second.Findings).FindingId);
        }

        [Fact]
        public void Propagates_logical_sheet_identity_across_a_supported_rename()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Old name", "StableCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "one"),
                Text("A3", "B"), Text("B3", "two")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("New name", "StableCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "one"),
                Text("A3", "B"), Text("B3", "two")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-rename",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.True(result.IsComplete);
            Assert.Equal(
                Revision(result, "before").Sheets.Single().LogicalSheetId,
                Revision(result, "after").Sheets.Single().LogicalSheetId);
            Assert.Equal("New name", result.Edges.Single().SheetAlignments.Single().AfterSheetName);
        }

        [Fact]
        public void Competing_sheet_targets_are_correctable_and_never_fabricated()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Original", "StableCode",
                Text("A1", "ID"), Text("B1", "Value"),
                Text("A2", "A"), Text("B2", "one")));
            OpenXmlWorkbookSnapshot after = Workbook("after",
                Sheet("Candidate A", "StableCode",
                    Text("A1", "ID"), Text("B1", "Value"),
                    Text("A2", "A"), Text("B2", "one")),
                Sheet("Candidate B", "StableCode",
                    Text("A1", "ID"), Text("B1", "Value"),
                    Text("A2", "A"), Text("B2", "one")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-ambiguous-sheets",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            OpenXmlProjectFinding finding = Assert.Single(
                result.Findings,
                item => item.Code == OpenXmlProjectFindingCodes.SheetMappingAmbiguous);
            Assert.Equal(DiagnosticSeverity.Correctable, finding.Severity);
            Assert.Equal(new[] { "Candidate A", "Candidate B" }, finding.CandidateSheetNames);
            Assert.Empty(result.Edges.Single().SheetAlignments);
            Assert.Null(Revision(result, "after").AlignedSnapshot);
        }

        [Fact]
        public void Duplicate_rows_remain_an_explicit_axis_ambiguity_and_block_history()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "DUP"), Text("B2", "10"),
                Text("A3", "DUP"), Text("B3", "10")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "DUP"), Text("B2", "10"),
                Text("A3", "DUP"), Text("B3", "10")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-duplicate-rows",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            Assert.Contains(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.RowAlignmentAmbiguous &&
                item.Severity == DiagnosticSeverity.Correctable);
            OpenXmlProjectSheetAlignment alignment = result.Edges.Single().SheetAlignments.Single();
            Assert.True(alignment.RowAlignment!.HasAmbiguity);
            Assert.False(alignment.IsComplete);
            Assert.Null(Revision(result, "after").AlignedSnapshot);
        }

        [Fact]
        public void Propagates_root_identities_into_sibling_branches_and_exposes_the_conflict()
        {
            OpenXmlWorkbookSnapshot root = Workbook("root", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "25"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot branch = Workbook("branch", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "27"),
                Text("A4", "C"), Text("B4", "30")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-branches",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("root", "", 0, root, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("master", "root", 1, master, isMaster: true),
                    new OpenXmlProjectRevision("branch", "root", 2, branch),
                });

            Assert.True(result.IsComplete);
            string rootB = Revision(result, "root").Sheets.Single().Rows.IdentityByPhysicalIndex[2];
            Assert.Equal(rootB, Revision(result, "master").Sheets.Single().Rows.IdentityByPhysicalIndex[2]);
            Assert.Equal(rootB, Revision(result, "branch").Sheets.Single().Rows.IdentityByPhysicalIndex[2]);
            Assert.Contains(result.History!.Subjects, subject => subject.HasConflict);
        }

        [Fact]
        public void Independently_inserted_sibling_rows_with_the_same_strong_key_block_history()
        {
            OpenXmlWorkbookSnapshot root = Workbook("root", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "20"),
                Text("A5", "C"), Text("B5", "30")));
            OpenXmlWorkbookSnapshot branch = Workbook("branch", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "17"),
                Text("A4", "B"), Text("B4", "20"),
                Text("A5", "C"), Text("B5", "30")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-sibling-row-insertion",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("root", "", 0, root, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("master", "root", 1, master, isMaster: true),
                    new OpenXmlProjectRevision("branch", "root", 2, branch),
                });

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            Assert.Equal(2, result.Findings.Count(item =>
                item.Code == OpenXmlProjectFindingCodes.SiblingRowInsertionAmbiguous &&
                item.Severity == DiagnosticSeverity.Correctable));
            Assert.All(result.Edges, edge => Assert.False(edge.IsComplete));
            Assert.Null(Revision(result, "master").AlignedSnapshot);
            Assert.Null(Revision(result, "branch").AlignedSnapshot);
        }

        [Fact]
        public void Independently_inserted_sibling_columns_with_the_same_unique_header_block_history()
        {
            OpenXmlWorkbookSnapshot root = Workbook("root", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"), Text("C1", "Notes"),
                Text("A2", "A"), Text("B2", "10"), Text("C2", "alpha"),
                Text("A3", "B"), Text("B3", "20"), Text("C3", "beta"),
                Text("A4", "C"), Text("B4", "30"), Text("C4", "gamma")));
            OpenXmlWorkbookSnapshot branch = Workbook("branch", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"), Text("C1", "Notes"),
                Text("A2", "A"), Text("B2", "10"), Text("C2", "one"),
                Text("A3", "B"), Text("B3", "20"), Text("C3", "two"),
                Text("A4", "C"), Text("B4", "30"), Text("C4", "three")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-sibling-column-insertion",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("root", "", 0, root, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("master", "root", 1, master, isMaster: true),
                    new OpenXmlProjectRevision("branch", "root", 2, branch),
                });

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            Assert.Equal(2, result.Findings.Count(item =>
                item.Code == OpenXmlProjectFindingCodes.SiblingColumnInsertionAmbiguous &&
                item.Severity == DiagnosticSeverity.Correctable));
            Assert.All(result.Edges, edge => Assert.False(edge.IsComplete));
        }

        [Fact]
        public void Same_strong_key_inserted_on_cousin_branches_at_different_depths_blocks_history()
        {
            OpenXmlWorkbookSnapshot root = Workbook("root", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot left = Workbook("left", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "15"),
                Text("A4", "B"), Text("B4", "20"),
                Text("A5", "C"), Text("B5", "30")));
            OpenXmlWorkbookSnapshot leftDescendant = Workbook("left-descendant", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "16"),
                Text("A4", "B"), Text("B4", "20"),
                Text("A5", "C"), Text("B5", "30")));
            OpenXmlWorkbookSnapshot right = Workbook("right", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot rightDeep = Workbook("right-deep", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "X"), Text("B3", "17"),
                Text("A4", "B"), Text("B4", "20"),
                Text("A5", "C"), Text("B5", "30")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-cousin-row-insertion",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("root", "", 0, root, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("left", "root", 1, left),
                    new OpenXmlProjectRevision("left-descendant", "left", 2, leftDescendant, isMaster: true),
                    new OpenXmlProjectRevision("right", "root", 3, right),
                    new OpenXmlProjectRevision("right-deep", "right", 4, rightDeep),
                });

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            OpenXmlProjectFinding[] findings = result.Findings
                .Where(item => item.Code == OpenXmlProjectFindingCodes.SiblingRowInsertionAmbiguous)
                .ToArray();
            Assert.Equal(2, findings.Length);
            Assert.Equal(new[] { "root", "right" },
                findings.OrderBy(item => item.AfterRevisionId, StringComparer.Ordinal)
                    .Select(item => item.BeforeRevisionId));
            Assert.All(findings, finding =>
            {
                Assert.Equal(DiagnosticSeverity.Correctable, finding.Severity);
                Assert.Contains("nearest common ancestor is 'root'", finding.Message, StringComparison.Ordinal);
            });
            Assert.False(result.Edges.Single(item => item.ChildRevisionId == "left").IsComplete);
            Assert.False(result.Edges.Single(item => item.ChildRevisionId == "left-descendant").IsComplete);
            Assert.True(result.Edges.Single(item => item.ChildRevisionId == "right").IsComplete);
            Assert.False(result.Edges.Single(item => item.ChildRevisionId == "right-deep").IsComplete);
            Assert.Null(Revision(result, "left-descendant").AlignedSnapshot);
        }

        [Fact]
        public void Same_name_sheets_created_independently_on_non_ancestor_branches_block_history()
        {
            OpenXmlWorkbookSnapshot root = Workbook("root", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10")));
            OpenXmlWorkbookSnapshot left = Workbook("left",
                Sheet("Data", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10")),
                Sheet("Notes", "LeftNotesCode",
                    Text("A1", "Topic"), Text("B1", "Text"),
                    Text("A2", "A"), Text("B2", "left")));
            OpenXmlWorkbookSnapshot right = Workbook("right",
                Sheet("Data", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10")),
                Sheet("notes", "RightNotesCode",
                    Text("A1", "Topic"), Text("B1", "Text"),
                    Text("A2", "A"), Text("B2", "right")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-independent-sheets",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("root", "", 0, root, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("left", "root", 1, left, isMaster: true),
                    new OpenXmlProjectRevision("right", "root", 2, right),
                });

            Assert.False(result.IsComplete);
            Assert.Null(result.History);
            OpenXmlProjectFinding[] findings = result.Findings
                .Where(item => item.Code == OpenXmlProjectFindingCodes.IndependentSheetInsertionAmbiguous)
                .ToArray();
            Assert.Equal(2, findings.Length);
            Assert.All(findings, finding =>
            {
                Assert.Equal(DiagnosticSeverity.Correctable, finding.Severity);
                Assert.Equal("root", finding.BeforeRevisionId);
                Assert.Contains("same case-insensitive sheet name", finding.Message, StringComparison.Ordinal);
            });
            Assert.All(result.Edges, edge => Assert.False(edge.IsComplete));
        }

        [Fact]
        public void Row_and_sheet_insertions_propagated_to_descendants_do_not_conflict()
        {
            OpenXmlWorkbookSnapshot root = Workbook("root", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20"),
                Text("A4", "C"), Text("B4", "30")));
            OpenXmlWorkbookSnapshot inserted = Workbook("inserted",
                Sheet("Data", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10"),
                    Text("A3", "X"), Text("B3", "15"),
                    Text("A4", "B"), Text("B4", "20"),
                    Text("A5", "C"), Text("B5", "30")),
                Sheet("Notes", "NotesCode",
                    Text("A1", "Topic"), Text("B1", "Text"),
                    Text("A2", "A"), Text("B2", "draft")));
            OpenXmlWorkbookSnapshot descendant = Workbook("descendant",
                Sheet("Data", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10"),
                    Text("A3", "X"), Text("B3", "16"),
                    Text("A4", "B"), Text("B4", "20"),
                    Text("A5", "C"), Text("B5", "30")),
                Sheet("Notes", "NotesCode",
                    Text("A1", "Topic"), Text("B1", "Text"),
                    Text("A2", "A"), Text("B2", "reviewed")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-descendant-propagation",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("root", "", 0, root, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("inserted", "root", 1, inserted),
                    new OpenXmlProjectRevision("descendant", "inserted", 2, descendant, isMaster: true),
                });

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            Assert.DoesNotContain(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.SiblingRowInsertionAmbiguous ||
                item.Code == OpenXmlProjectFindingCodes.SiblingColumnInsertionAmbiguous ||
                item.Code == OpenXmlProjectFindingCodes.IndependentSheetInsertionAmbiguous);
            string insertedNotes = Revision(result, "inserted").Sheets
                .Single(item => item.PhysicalSheetName == "Notes").LogicalSheetId;
            Assert.Equal(insertedNotes, Revision(result, "descendant").Sheets
                .Single(item => item.PhysicalSheetName == "Notes").LogicalSheetId);
        }

        [Fact]
        public void Validates_the_explicit_revision_tree_before_alignment()
        {
            OpenXmlWorkbookSnapshot r1 = Workbook("r1", Sheet("Data", "DataCode", Text("A1", "x")));
            OpenXmlWorkbookSnapshot r2 = Workbook("r2", Sheet("Data", "DataCode", Text("A1", "x")));

            Assert.Throws<ArgumentException>(() => OpenXmlComparisonProjectEngine.Build(
                "project-invalid",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("r1", "", 1, r1, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("r2", "r1", 0, r2, isMaster: true),
                }));
        }

        [Fact]
        public void Reports_bounded_sheet_progress_with_semantic_alignment_stages()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "11"),
                Text("A3", "B"), Text("B3", "20")));
            var updates = new List<OpenXmlComparisonProjectProgress>();

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-progress",
                "mapping-progress",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                },
                progress: new InlineProgress<OpenXmlComparisonProjectProgress>(updates.Add));

            Assert.True(result.IsComplete);
            Assert.NotEmpty(updates);
            Assert.Contains(updates, item => item.Stage == OpenXmlComparisonProjectProgressStage.MappingSheets);
            Assert.Contains(updates, item => item.Stage == OpenXmlComparisonProjectProgressStage.DetectingHeaders);
            Assert.Contains(updates, item => item.Stage == OpenXmlComparisonProjectProgressStage.InferringKeys);
            Assert.Contains(updates, item => item.Stage == OpenXmlComparisonProjectProgressStage.AligningRows);
            Assert.Contains(updates, item => item.Stage == OpenXmlComparisonProjectProgressStage.AligningColumns);
            Assert.Contains(updates, item => item.Stage == OpenXmlComparisonProjectProgressStage.BuildingHistory);
            Assert.Equal(OpenXmlComparisonProjectProgressStage.Complete, updates.Last().Stage);
            Assert.All(updates, item => Assert.Equal(2, item.TotalWorkItems));
            Assert.Equal(2, updates.Last().CompletedWorkItems);
            Assert.True(updates.Zip(updates.Skip(1), (left, right) =>
                left.CompletedWorkItems <= right.CompletedWorkItems).All(value => value));
        }

        [Fact]
        public void Cancellation_from_alignment_progress_stops_before_more_expensive_work()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "11"),
                Text("A3", "B"), Text("B3", "20")));
            using var cancellation = new CancellationTokenSource();
            var observedStages = new List<OpenXmlComparisonProjectProgressStage>();
            var progress = new InlineProgress<OpenXmlComparisonProjectProgress>(update =>
            {
                observedStages.Add(update.Stage);
                if (update.Stage == OpenXmlComparisonProjectProgressStage.AligningRows)
                    cancellation.Cancel();
            });

            Assert.Throws<OperationCanceledException>(() => OpenXmlComparisonProjectEngine.Build(
                "project-cancel-progress",
                "mapping-cancel-progress",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                },
                cancellationToken: cancellation.Token,
                progress: progress));

            Assert.Contains(OpenXmlComparisonProjectProgressStage.AligningRows, observedStages);
            Assert.DoesNotContain(OpenXmlComparisonProjectProgressStage.AligningColumns, observedStages);
            Assert.DoesNotContain(OpenXmlComparisonProjectProgressStage.Complete, observedStages);
        }

        private static OpenXmlAlignedProjectRevision Revision(
            OpenXmlComparisonProjectResult result,
            string revisionId) =>
            result.Revisions.Single(item => item.Source.RevisionId == revisionId);

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

        private sealed class InlineProgress<T> : IProgress<T>
        {
            private readonly Action<T> _report;

            internal InlineProgress(Action<T> report)
            {
                _report = report;
            }

            public void Report(T value) => _report(value);
        }
    }
}
