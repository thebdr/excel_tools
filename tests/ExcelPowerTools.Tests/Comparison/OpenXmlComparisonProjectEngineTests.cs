using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;
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
            Assert.DoesNotContain(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.NonFormulaValueContentUnmodeled);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.Equal(2, result.Edges.Count);

            OpenXmlAlignedSheetRevision r1Sheet = Revision(result, "r1").Sheets.Single();
            OpenXmlAlignedSheetRevision r2Sheet = Revision(result, "r2").Sheets.Single();
            OpenXmlAlignedSheetRevision r3Sheet = Revision(result, "r3").Sheets.Single();
            string logicalB = r1Sheet.Rows.IdentityByPhysicalIndex[2];
            Assert.Equal(logicalB, r2Sheet.Rows.IdentityByPhysicalIndex[3]);
            Assert.Equal(logicalB, r3Sheet.Rows.IdentityByPhysicalIndex[3]);
            string logicalX = r2Sheet.Rows.IdentityByPhysicalIndex[2];
            Assert.DoesNotContain(logicalX, r1Sheet.Rows.IdentityByPhysicalIndex.Values);
            Assert.Equal(logicalX, r3Sheet.Rows.IdentityByPhysicalIndex[2]);
            LogicalAxisStructuralEvent inserted = Assert.Single(
                r2Sheet.Rows.StructuralEvents,
                item => item.Kind == LogicalAxisStructuralEventKind.Insertion);
            Assert.Null(inserted.BeforeIndex);
            Assert.Equal(2, inserted.AfterIndex);
            Assert.Empty(r2Sheet.Columns.StructuralEvents);
            Assert.True(result.History!.AtomicTransitionCount > 0);
        }

        [Fact]
        public void Builds_complete_history_and_keeps_column_identity_through_an_insertion()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Name"), Text("C1", "Amount"),
                Text("A2", "A"), Text("B2", "Alpha"), Text("C2", "10"),
                Text("A3", "B"), Text("B3", "Beta"), Text("C3", "20")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Notes"), Text("C1", "Name"), Text("D1", "Amount"),
                Text("A2", "A"), Text("B2", "first"), Text("C2", "Alpha"), Text("D2", "10"),
                Text("A3", "B"), Text("B3", "second"), Text("C3", "Beta"), Text("D3", "20")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-column-insertion",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            Assert.True(Assert.Single(result.Edges).IsComplete);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);

            OpenXmlAlignedSheetRevision beforeSheet = Revision(result, "before").Sheets.Single();
            OpenXmlAlignedSheetRevision afterSheet = Revision(result, "after").Sheets.Single();
            Assert.Equal(
                beforeSheet.Columns.IdentityByPhysicalIndex[0],
                afterSheet.Columns.IdentityByPhysicalIndex[0]);
            Assert.Equal(
                beforeSheet.Columns.IdentityByPhysicalIndex[1],
                afterSheet.Columns.IdentityByPhysicalIndex[2]);
            Assert.Equal(
                beforeSheet.Columns.IdentityByPhysicalIndex[2],
                afterSheet.Columns.IdentityByPhysicalIndex[3]);
            Assert.Equal(
                LogicalAxisIdentitySource.Inserted,
                afterSheet.Columns.AssignmentsByPhysicalIndex[1].Source);

            LogicalAxisStructuralEvent inserted = Assert.Single(
                afterSheet.Columns.StructuralEvents,
                item => item.Kind == LogicalAxisStructuralEventKind.Insertion);
            Assert.Null(inserted.BeforeIndex);
            Assert.Equal(1, inserted.AfterIndex);
            Assert.Empty(afterSheet.Rows.StructuralEvents);
            Assert.Equal(
                beforeSheet.Rows.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value),
                afterSheet.Rows.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value));
            Assert.True(result.History!.AtomicTransitionCount > 0);
        }

        [Fact]
        public void Identical_tables_complete_with_stable_axes_and_zero_history_transitions()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet("Data", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10"),
                    Text("A3", "B"), Text("B3", "20")),
            });
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes("before", bytes);
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes("after", bytes);

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-identical",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.Equal(before.ByteContentId, after.ByteContentId);
            Assert.True(result.IsComplete);
            Assert.Empty(result.Findings);
            Assert.True(Assert.Single(result.Edges).IsComplete);
            Assert.NotNull(result.History);
            Assert.Equal(0, result.History!.AtomicTransitionCount);

            OpenXmlAlignedSheetRevision beforeSheet = Revision(result, "before").Sheets.Single();
            OpenXmlAlignedSheetRevision afterSheet = Revision(result, "after").Sheets.Single();
            Assert.Empty(afterSheet.Rows.StructuralEvents);
            Assert.Empty(afterSheet.Columns.StructuralEvents);
            Assert.Equal(
                beforeSheet.Rows.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value),
                afterSheet.Rows.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value));
            Assert.Equal(
                beforeSheet.Columns.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value),
                afterSheet.Columns.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value));
        }

        [Fact]
        public void Stored_constant_value_change_is_reported_as_one_value_transition()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Status"),
                Text("A2", "ITEM-001"), Text("B2", "Draft"),
                Text("A3", "ITEM-002"), Text("B3", "Review"),
                Text("A4", "ITEM-003"), Text("B4", "Approved")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Status"),
                Text("A2", "ITEM-001"), Text("B2", "Draft"),
                Text("A3", "ITEM-002"), Text("B3", "Approved"),
                Text("A4", "ITEM-003"), Text("B4", "Approved")));

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-stored-value-change",
                before,
                after);

            Assert.True(result.IsComplete);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.NotNull(result.History);
            Assert.Equal(1, result.History!.AtomicTransitionCount);
            var subject = Assert.Single(result.History.Subjects);
            Assert.Equal(CellChangeAspect.Value, subject.Aspect);
            Assert.Equal(CellChangeKind.Modified, Assert.Single(subject.Transitions).ExactKind);
            Assert.Equal(
                "Review",
                subject.Observations.Single(item => item.RevisionId == "before").ExactValue.ExactText);
            Assert.Equal(
                "Approved",
                subject.Observations.Single(item => item.RevisionId == "after").ExactValue.ExactText);
        }

        [Fact]
        public void Stored_formula_text_change_is_reported_as_one_formula_transition()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Result"),
                Text("A2", "ITEM-001"), Formula("B2", "1+1", "2"),
                Text("A3", "ITEM-002"), Formula("B3", "2+2", "4"),
                Text("A4", "ITEM-003"), Formula("B4", "3+3", "6")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Result"),
                Text("A2", "ITEM-001"), Formula("B2", "1+1", "2"),
                Text("A3", "ITEM-002"), Formula("B3", "2+3", "4"),
                Text("A4", "ITEM-003"), Formula("B4", "3+3", "6")));

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-formula-text-change",
                before,
                after);

            Assert.True(result.IsComplete);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.NotNull(result.History);
            Assert.Equal(1, result.History!.AtomicTransitionCount);
            var subject = Assert.Single(result.History.Subjects);
            Assert.Equal(CellChangeAspect.Formula, subject.Aspect);
            Assert.Equal(CellChangeKind.Modified, Assert.Single(subject.Transitions).ExactKind);
            Assert.Equal(
                "2+2",
                subject.Observations.Single(item => item.RevisionId == "before").ExactText);
            Assert.Equal(
                "2+3",
                subject.Observations.Single(item => item.RevisionId == "after").ExactText);
        }

        [Fact]
        public void Stored_cached_result_change_is_reported_as_one_cached_result_transition()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Result"),
                Text("A2", "ITEM-001"), Formula("B2", "1+1", "2"),
                Text("A3", "ITEM-002"), Formula("B3", "2+2", "4"),
                Text("A4", "ITEM-003"), Formula("B4", "3+3", "6")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Result"),
                Text("A2", "ITEM-001"), Formula("B2", "1+1", "2"),
                Text("A3", "ITEM-002"), Formula("B3", "2+2", "5"),
                Text("A4", "ITEM-003"), Formula("B4", "3+3", "6")));

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-cached-result-change",
                before,
                after);

            Assert.True(result.IsComplete);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.NotNull(result.History);
            Assert.Equal(1, result.History!.AtomicTransitionCount);
            var subject = Assert.Single(result.History.Subjects);
            Assert.Equal(CellChangeAspect.CachedResult, subject.Aspect);
            Assert.Equal(CellChangeKind.Modified, Assert.Single(subject.Transitions).ExactKind);
            Assert.Equal(
                "4",
                subject.Observations.Single(item => item.RevisionId == "before").ExactValue.ExactText);
            Assert.Equal(
                "5",
                subject.Observations.Single(item => item.RevisionId == "after").ExactValue.ExactText);
        }

        [Fact]
        public void Style_and_package_only_byte_difference_is_silent_with_zero_history_transitions()
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

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-style-package-only",
                before,
                after);

            Assert.NotEqual(before.ByteContentId, after.ByteContentId);
            Assert.True(result.IsComplete);
            Assert.True(Assert.Single(result.Edges).IsComplete);
            Assert.Empty(result.Findings);
            Assert.NotNull(result.History);
            Assert.Equal(0, result.History!.AtomicTransitionCount);
            Assert.Empty(result.History.Subjects);
        }

        [Fact]
        public void Readable_sheet_and_numeric_payload_ignore_noncontent_sheet_style_and_date_metadata()
        {
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[]
                {
                    Sheet("Data", "DataCode",
                        Text("A1", "ID"), Text("B1", "Amount"),
                        Text("A2", "A"), Number("B2", "45000")),
                }));
            Cell styledNumber = Number("B2", "45000");
            styledNumber.StyleIndex = 99U;
            OpenXmlWorkbookFixture.SheetDefinition afterDefinition = Sheet(
                "Data",
                "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), styledNumber);
            afterDefinition.OmitSheetId = true;
            afterDefinition.RawState = "future-visibility";
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[] { afterDefinition }, date1904: true));

            Assert.True(after.CanAssertExactCellEquality);
            Assert.Contains(after.Sheets[0].CoverageMessages, item =>
                item.Code == OpenXmlCoverageCodes.InvalidCellStyleIndex &&
                !item.AffectsExactCellComparison);
            Assert.Contains(after.Sheets[0].CoverageMessages, item =>
                item.Code == OpenXmlCoverageCodes.MissingSheetId &&
                !item.AffectsExactCellComparison);
            Assert.Contains(after.Sheets[0].CoverageMessages, item =>
                item.Code == OpenXmlCoverageCodes.UnknownSheetVisibility &&
                !item.AffectsExactCellComparison);

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-noncontent-metadata",
                before,
                after);

            Assert.True(result.IsComplete);
            Assert.Empty(result.Findings);
            Assert.NotNull(result.History);
            Assert.Equal(0, result.History!.AtomicTransitionCount);
        }

        [Fact]
        public void Cell_evidence_blocker_excludes_irrelevant_package_metadata_reason_codes()
        {
            OpenXmlWorkbookSnapshot before = WorkbookWithCellAndPackageCoverage("before");
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "readable")));

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-filtered-cell-coverage",
                before,
                after);

            Assert.False(result.IsComplete);
            OpenXmlProjectFinding blocker = Assert.Single(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.IncompleteCellEvidence &&
                item.AfterRevisionId == "before");
            Assert.Equal(DiagnosticSeverity.ArtifactBlocking, blocker.Severity);
            Assert.Contains(OpenXmlCoverageCodes.MalformedCellPayload, blocker.ReasonCodes);
            Assert.DoesNotContain(OpenXmlCoverageCodes.UnsupportedPackagePart, blocker.ReasonCodes);
            Assert.DoesNotContain(OpenXmlCoverageCodes.UnmodeledWorkbookMetadata, blocker.ReasonCodes);
        }

        [Fact]
        public void Unresolved_sheet_remains_blocked_with_its_relevant_reason_code()
        {
            OpenXmlWorkbookSnapshot before = WorkbookWithUnresolvedSheetCoverage("before");
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "readable")));

            OpenXmlComparisonProjectResult result = BuildTwoRevisionProject(
                "project-unresolved-sheet-coverage",
                before,
                after);

            Assert.False(result.IsComplete);
            OpenXmlProjectFinding blocker = Assert.Single(result.Findings, item =>
                item.Code == OpenXmlProjectFindingCodes.UnsupportedSheet &&
                item.AfterRevisionId == "before");
            Assert.Equal(DiagnosticSeverity.ArtifactBlocking, blocker.Severity);
            Assert.Contains(OpenXmlCoverageCodes.UnresolvedSheetRelationship, blocker.ReasonCodes);
            Assert.DoesNotContain(OpenXmlCoverageCodes.UnsupportedPackagePart, blocker.ReasonCodes);
        }

        [Fact]
        public void Reviewable_dynamic_array_coverage_warns_without_blocking_project_history()
        {
            OpenXmlWorkbookSnapshot before = DynamicArrayWorkbook("before");
            OpenXmlWorkbookSnapshot after = DynamicArrayWorkbook("after");

            foreach (OpenXmlWorkbookSnapshot snapshot in new[] { before, after })
            {
                OpenXmlWorksheetSnapshot sheet = Assert.Single(snapshot.Sheets);
                Assert.False(snapshot.CanAssertExactCellEquality);
                Assert.False(sheet.CanAssertExactCellEquality);
                Assert.True(sheet.CanProjectExactStoredFormulaValueEvidence);
                OpenXmlFormulaObservation formula = Assert.Single(sheet.FormulaObservations);
                Assert.Equal(OpenXmlFormulaForm.DynamicArray, formula.Form);
                Assert.Equal("_xlfn.XLOOKUP(A2,A:A,B:B)", formula.StoredFormulaText);
                Assert.Equal(formula.StoredFormulaText, formula.EffectiveFormulaText);
                Assert.Equal("B2", formula.FormulaRange);
                Assert.Equal(CanonicalValue.Number(7D), formula.CachedResult);
                Assert.Contains(
                    sheet.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.DynamicArrayNotExpanded &&
                            item.AffectsExactCellComparison);
            }

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-reviewable-dynamic-array",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.True(result.IsComplete);
            Assert.True(Assert.Single(result.Edges).IsComplete);
            Assert.NotNull(result.History);
            Assert.Equal(0, result.History!.AtomicTransitionCount);
            OpenXmlProjectFinding[] coverageWarnings = result.Findings
                .Where(item => item.Code == OpenXmlProjectFindingCodes.IncompleteCellEvidence)
                .ToArray();
            Assert.NotEmpty(coverageWarnings);
            Assert.All(coverageWarnings, finding =>
            {
                Assert.Equal(DiagnosticSeverity.Warning, finding.Severity);
                Assert.Contains(OpenXmlCoverageCodes.DynamicArrayNotExpanded, finding.ReasonCodes);
                Assert.Contains("Merge remains blocked", finding.SuggestedAction, StringComparison.Ordinal);
            });
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.All(result.Revisions, revision =>
                Assert.False(revision.Source.Snapshot.CanAssertExactCellEquality));
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
        public void Repeated_rows_align_monotonically_and_emit_only_the_surplus_insertion()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "DUP"), Text("B2", "10"),
                Text("A3", "DUP"), Text("B3", "10")));
            OpenXmlWorkbookSnapshot after = Workbook("after", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "DUP"), Text("B2", "10"),
                Text("A3", "DUP"), Text("B3", "10"),
                Text("A4", "DUP"), Text("B4", "10")));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-duplicate-rows",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            Assert.True(result.History!.AtomicTransitionCount > 0);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            OpenXmlProjectSheetAlignment alignment = result.Edges.Single().SheetAlignments.Single();
            Assert.False(alignment.RowAlignment.HasAmbiguity);
            Assert.True(alignment.IsComplete);
            Assert.True(alignment.RowAlignment.TryMapBefore(1, out int firstDuplicateAfter));
            Assert.Equal(1, firstDuplicateAfter);
            Assert.True(alignment.RowAlignment.TryMapBefore(2, out int secondDuplicateAfter));
            Assert.Equal(2, secondDuplicateAfter);
            AxisAlignmentLink surplus = Assert.Single(
                alignment.RowAlignment.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter);
            Assert.Equal(3, surplus.AfterIndex);
            Assert.DoesNotContain(
                alignment.RowAlignment.Links,
                link => link.Kind == AxisAlignmentKind.OnlyBefore);

            OpenXmlAlignedSheetRevision beforeSheet = Revision(result, "before").Sheets.Single();
            OpenXmlAlignedSheetRevision afterSheet = Revision(result, "after").Sheets.Single();
            Assert.Equal(
                beforeSheet.Rows.IdentityByPhysicalIndex[1],
                afterSheet.Rows.IdentityByPhysicalIndex[1]);
            Assert.Equal(
                beforeSheet.Rows.IdentityByPhysicalIndex[2],
                afterSheet.Rows.IdentityByPhysicalIndex[2]);
            LogicalAxisStructuralEvent insertion = Assert.Single(
                afterSheet.Rows.StructuralEvents,
                item => item.Kind == LogicalAxisStructuralEventKind.Insertion);
            Assert.Null(insertion.BeforeIndex);
            Assert.Equal(3, insertion.AfterIndex);
        }

        [Fact]
        public void Identical_duplicate_rows_and_headers_use_stable_positional_identities()
        {
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[]
                {
                    Sheet("Data", "DataCode",
                        Text("A1", "ID"), Text("B1", "Value"), Text("C1", "Value"),
                        Text("A2", "DUP"), Text("B2", "same"), Text("C2", "same"),
                        Text("A3", "DUP"), Text("B3", "same"), Text("C3", "same")),
                }));
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[]
                {
                    Sheet("Data", "DataCode",
                        Text("A1", "ID"), Text("B1", "Value"), Text("C1", "Value"),
                        Text("A2", "DUP"), Text("B2", "same"), Text("C2", "same"),
                        Text("A3", "DUP"), Text("B3", "same"), Text("C3", "same")),
                }, includeDateStyle: true));

            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-identical-duplicate-table",
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

            Assert.NotEqual(before.ByteContentId, after.ByteContentId);
            Assert.True(result.IsComplete);
            Assert.True(Assert.Single(result.Edges).IsComplete);
            Assert.DoesNotContain(result.Findings, item => item.Severity >= DiagnosticSeverity.Correctable);
            Assert.NotNull(result.History);
            Assert.Equal(0, result.History!.AtomicTransitionCount);

            OpenXmlProjectSheetAlignment alignment = result.Edges.Single().SheetAlignments.Single();
            Assert.False(alignment.RowAlignment.HasAmbiguity);
            Assert.False(alignment.ColumnAlignment.HasAmbiguity);
            OpenXmlAlignedSheetRevision beforeSheet = Revision(result, "before").Sheets.Single();
            OpenXmlAlignedSheetRevision afterSheet = Revision(result, "after").Sheets.Single();
            Assert.Empty(afterSheet.Rows.StructuralEvents);
            Assert.Empty(afterSheet.Columns.StructuralEvents);
            Assert.Equal(
                beforeSheet.Rows.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value),
                afterSheet.Rows.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value));
            Assert.Equal(
                beforeSheet.Columns.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value),
                afterSheet.Columns.IdentityByPhysicalIndex.OrderBy(item => item.Key).Select(item => item.Value));
            Assert.All(afterSheet.Rows.AssignmentsByPhysicalIndex, pair =>
            {
                Assert.Equal(LogicalAxisIdentitySource.Inherited, pair.Value.Source);
                Assert.Equal<int?>(pair.Key, pair.Value.ParentPhysicalIndex);
            });
            Assert.All(afterSheet.Columns.AssignmentsByPhysicalIndex, pair =>
            {
                Assert.Equal(LogicalAxisIdentitySource.Inherited, pair.Value.Source);
                Assert.Equal<int?>(pair.Key, pair.Value.ParentPhysicalIndex);
            });
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

        private static OpenXmlComparisonProjectResult BuildTwoRevisionProject(
            string projectId,
            OpenXmlWorkbookSnapshot before,
            OpenXmlWorkbookSnapshot after) =>
            OpenXmlComparisonProjectEngine.Build(
                projectId,
                "mapping-1",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("after", "before", 1, after, isMaster: true),
                });

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

        private static OpenXmlWorkbookSnapshot DynamicArrayWorkbook(string revisionId)
        {
            const string formulaText = "_xlfn.XLOOKUP(A2,A:A,B:B)";
            const string exactMetadata = "dynamic-array-metadata";
            var formulaObservation = new OpenXmlFormulaObservation(
                "B2",
                OpenXmlFormulaForm.DynamicArray,
                formulaText,
                formulaText,
                "B2",
                null,
                false,
                exactMetadata,
                CanonicalValue.Number(7D));
            var coverage = new OpenXmlCoverageMessage(
                OpenXmlCoverageCodes.DynamicArrayNotExpanded,
                "The exact stored dynamic-array formula is reviewable, but spill semantics are not modeled.",
                true,
                true,
                "Data",
                "B2",
                "/xl/worksheets/sheet1.xml");
            var sheet = new OpenXmlWorksheetSnapshot(
                "Data",
                1U,
                "rId1",
                "/xl/worksheets/sheet1.xml",
                "DataCode",
                OpenXmlSheetVisibility.Visible,
                OpenXmlPhysicalSheetKind.Worksheet,
                new RevisionSnapshot(revisionId, new[]
                {
                    CanonicalCellState.Constant("A1", CanonicalValue.Text("ID")),
                    CanonicalCellState.Constant("B1", CanonicalValue.Text("Result")),
                    CanonicalCellState.Constant("A2", CanonicalValue.Text("A")),
                    CanonicalCellState.Formula(
                        "B2",
                        formulaText,
                        CanonicalValue.Number(7D),
                        exactFormulaMetadata: exactMetadata),
                    CanonicalCellState.Constant("A3", CanonicalValue.Text("B")),
                    CanonicalCellState.Constant("B3", CanonicalValue.Number(8D)),
                }),
                new List<OpenXmlFormulaObservation> { formulaObservation },
                new List<OpenXmlTableFormulaObservation>(),
                new List<OpenXmlCoverageMessage> { coverage });
            return new OpenXmlWorkbookSnapshot(
                revisionId,
                "same-dynamic-array-package",
                ExcelDateSystem.Date1900,
                false,
                new List<OpenXmlWorksheetSnapshot> { sheet },
                new List<OpenXmlCoverageMessage>());
        }

        private static OpenXmlWorkbookSnapshot WorkbookWithCellAndPackageCoverage(string revisionId)
        {
            var cellCoverage = new OpenXmlCoverageMessage(
                OpenXmlCoverageCodes.MalformedCellPayload,
                "The stored cell payload is malformed.",
                true,
                true,
                "Data",
                "A1",
                "/xl/worksheets/sheet1.xml");
            var sheet = new OpenXmlWorksheetSnapshot(
                "Data",
                1U,
                "rId1",
                "/xl/worksheets/sheet1.xml",
                "DataCode",
                OpenXmlSheetVisibility.Visible,
                OpenXmlPhysicalSheetKind.Worksheet,
                new RevisionSnapshot(revisionId, new[]
                {
                    CanonicalCellState.Constant("A1", CanonicalValue.Text("#malformed")),
                }),
                new List<OpenXmlFormulaObservation>(),
                new List<OpenXmlTableFormulaObservation>(),
                new List<OpenXmlCoverageMessage> { cellCoverage });
            return new OpenXmlWorkbookSnapshot(
                revisionId,
                "package-" + revisionId,
                ExcelDateSystem.Date1900,
                false,
                new List<OpenXmlWorksheetSnapshot> { sheet },
                new List<OpenXmlCoverageMessage>
                {
                    new OpenXmlCoverageMessage(
                        OpenXmlCoverageCodes.UnsupportedPackagePart,
                        "A drawing or comment part is outside the three content channels.",
                        true,
                        false,
                        null,
                        null,
                        null),
                    new OpenXmlCoverageMessage(
                        OpenXmlCoverageCodes.UnmodeledWorkbookMetadata,
                        "Workbook package metadata is outside the three content channels.",
                        true,
                        false,
                        null,
                        null,
                        null),
                });
        }

        private static OpenXmlWorkbookSnapshot WorkbookWithUnresolvedSheetCoverage(string revisionId)
        {
            var unresolved = new OpenXmlCoverageMessage(
                OpenXmlCoverageCodes.UnresolvedSheetRelationship,
                "The sheet relationship could not be resolved.",
                true,
                true,
                "Data",
                null,
                "/xl/workbook.xml");
            var sheet = new OpenXmlWorksheetSnapshot(
                "Data",
                1U,
                "missingRelationship",
                "",
                null,
                OpenXmlSheetVisibility.Visible,
                OpenXmlPhysicalSheetKind.Unsupported,
                new RevisionSnapshot(revisionId, null),
                new List<OpenXmlFormulaObservation>(),
                new List<OpenXmlTableFormulaObservation>(),
                new List<OpenXmlCoverageMessage> { unresolved });
            return new OpenXmlWorkbookSnapshot(
                revisionId,
                "package-" + revisionId,
                ExcelDateSystem.Date1900,
                false,
                new List<OpenXmlWorksheetSnapshot> { sheet },
                new List<OpenXmlCoverageMessage>
                {
                    new OpenXmlCoverageMessage(
                        OpenXmlCoverageCodes.UnsupportedPackagePart,
                        "A drawing or comment part is outside the three content channels.",
                        true,
                        false,
                        null,
                        null,
                        null),
                });
        }

        private static Cell Text(string reference, string value) =>
            new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new SpreadsheetText(value)),
            };

        private static Cell Formula(string reference, string formula, string cachedValue) =>
            new Cell(new CellFormula(formula), new CellValue(cachedValue))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
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
