using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Merge;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Project;
using Xunit;
using SpreadsheetText = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewDecisionCatalogProjectorTests
    {
        [Fact]
        public void Projects_every_revision_occurrence_with_deterministic_evidence_and_locations()
        {
            OpenXmlComparisonProjectResult first = InsertedRowProject();

            ReviewDecisionCatalog left = ReviewDecisionCatalogProjector.Project(
                first,
                Evidence(first).Reverse());
            ReviewDecisionCatalog right = ReviewDecisionCatalogProjector.Project(
                first,
                Evidence(first));

            Assert.Equal(left.CatalogId, right.CatalogId);
            Assert.Equal(left.EvidenceGenerationId, right.EvidenceGenerationId);
            Assert.NotNull(left.EvidenceGeneration);
            var expectedEvidence = new ReviewEvidenceGeneration(
                first.ProjectId,
                first.EffectiveMappingGenerationId,
                first.Revisions.Select(item => new ReviewEvidenceRevisionIdentity(
                    item.Source.RevisionId,
                    item.Source.Order,
                    item.Source.Snapshot.ByteContentId)));
            Assert.Equal(expectedEvidence.GenerationId, left.EvidenceGenerationId);
            Assert.Equal(
                new[] { "r1", "r2", "master" },
                left.EvidenceGeneration!.Revisions.Select(item => item.RevisionId));
            Assert.Equal(
                first.Revisions.OrderBy(item => item.Source.Order)
                    .Select(item => item.Source.Snapshot.ByteContentId),
                left.EvidenceGeneration.Revisions.Select(item => item.ByteContentId));
            Assert.Equal(
                left.Subjects.Count * first.Revisions.Count,
                left.Observations.Count);
            Assert.All(left.Subjects, subject => Assert.Equal(
                new[] { "master", "r1", "r2" },
                left.Observations.Where(item => item.SubjectId == subject.SubjectId)
                    .Select(item => item.RevisionId)
                    .OrderBy(item => item, StringComparer.Ordinal)));

            ReviewDecisionSubject changedAmount = Assert.Single(left.Subjects, subject =>
                subject.Aspect == ComparisonAspect.Value &&
                left.Observations.Any(item =>
                    item.SubjectId == subject.SubjectId &&
                    item.DisplayText == "25" &&
                    item.Provenance.Address == "B4"));
            Dictionary<string, Observation> amountByRevision = left.Observations
                .Where(item => item.SubjectId == changedAmount.SubjectId)
                .ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            Assert.Equal("B3", amountByRevision["r1"].Provenance.Address);
            Assert.Equal("B4", amountByRevision["r2"].Provenance.Address);
            Assert.Equal("B4", amountByRevision["master"].Provenance.Address);
            Assert.Equal("r1.xlsx", amountByRevision["r1"].Provenance.SourceLabel);
            Assert.NotEmpty(amountByRevision["r1"].Provenance.LogicalSheetId);
            Assert.Equal(
                amountByRevision["r1"].Provenance.LogicalSheetId,
                amountByRevision["master"].Provenance.LogicalSheetId);
            Assert.Equal(3, amountByRevision["r1"].Provenance.PhysicalRow);
            Assert.Equal(2, amountByRevision["r1"].Provenance.PhysicalColumn);
            var historyAmount = first.History!.Subjects.Single(item =>
                item.SubjectId == changedAmount.SubjectId);
            Assert.Equal(
                historyAmount.Observations.Single(item => item.RevisionId == "r1").StateId,
                amountByRevision["r1"].ContentHash);
            Assert.Equal(
                ComparisonIds.Observation(
                    changedAmount.SubjectId,
                    "r1",
                    amountByRevision["r1"].ExactType,
                    amountByRevision["r1"].ContentHash),
                amountByRevision["r1"].ObservationId);

            Assert.NotEmpty(first.History.Groups);
            Assert.Equal(first.History.Groups.Select(item => item.GroupId).OrderBy(item => item),
                left.Groups.Select(item => item.GroupId).OrderBy(item => item));
            foreach (ReviewDecisionGroup group in left.Groups)
            {
                Assert.Equal(
                    first.History.Groups.Single(item => item.GroupId == group.GroupId)
                        .SubjectIds.OrderBy(item => item),
                    group.SubjectIds);
            }
        }

        [Fact]
        public void Formula_value_and_cached_evidence_keep_exact_types_and_explicit_policies()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Formula"), Text("C1", "Amount"),
                Text("A2", "A"), Formula("B2", "1+1", "2"), Number("C2", "10"),
                Text("A3", "B"), Formula("B3", "2+2", "4"), Number("C3", "20")));
            OpenXmlWorkbookSnapshot middle = Workbook("middle", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Formula"), Text("C1", "Amount"),
                Text("A2", "A"), Formula("B2", "1+2", "3"), Number("C2", "11"),
                Text("A3", "B"), Formula("B3", "2+2", "4"), Number("C3", "20")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Formula"), Text("C1", "Amount"),
                Text("A2", "A"), Number("B2", "3"), Number("C2", "12"),
                Text("A3", "B"), Formula("B3", "2+2", "4"), Number("C3", "20")));
            OpenXmlComparisonProjectResult result = Build(
                "project-formula-value",
                before,
                middle,
                master);

            ReviewDecisionCatalog catalog = ReviewDecisionCatalogProjector.Project(result, Evidence(result));

            ReviewDecisionSubject formula = Assert.Single(catalog.Subjects, item =>
                item.Aspect == ComparisonAspect.Formula);
            Assert.Equal(ReviewDecisionSubjectPolicy.SelectableAndMergeable, formula.Policy);
            Dictionary<string, Observation> formulas = catalog.Observations
                .Where(item => item.SubjectId == formula.SubjectId)
                .ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            Assert.Equal(ObservationExactType.Formula, formulas["before"].ExactType);
            Assert.Equal("1+1", formulas["before"].Formula);
            Assert.Equal(ObservationExactType.Formula, formulas["middle"].ExactType);
            Assert.Equal("1+2", formulas["middle"].Formula);
            Assert.Equal(ObservationExactType.Blank, formulas["master"].ExactType);
            Assert.Equal("B2", formulas["master"].Provenance.Address);

            ReviewDecisionSubject amount = Assert.Single(catalog.Subjects, item =>
                item.Aspect == ComparisonAspect.Value &&
                catalog.Observations.Any(observation =>
                    observation.SubjectId == item.SubjectId &&
                    observation.Value?.CanonicalText == "12"));
            Assert.Equal(ReviewDecisionSubjectPolicy.SelectableAndMergeable, amount.Policy);
            Assert.All(catalog.Observations.Where(item => item.SubjectId == amount.SubjectId), item =>
                Assert.Equal(ObservationExactType.Number, item.ExactType));
            Assert.Equal(
                new[] { "10", "11", "12" },
                catalog.Observations.Where(item => item.SubjectId == amount.SubjectId)
                    .OrderBy(item => result.Revisions.Single(revision =>
                        revision.Source.RevisionId == item.RevisionId).Source.Order)
                    .Select(item => item.Value!.CanonicalText));

            ReviewDecisionSubject cached = Assert.Single(catalog.Subjects, item =>
                item.Aspect == ComparisonAspect.CachedResult);
            Assert.Equal(ReviewDecisionSubjectPolicy.ReviewOnly, cached.Policy);
            Assert.False(cached.CanSelectObservation);
            Assert.False(cached.CanMergeSelection);
            ReviewDecisionSubject formulaCellValue = Assert.Single(catalog.Subjects, item =>
                item.Aspect == ComparisonAspect.Value &&
                catalog.Observations.Any(observation =>
                    observation.SubjectId == item.SubjectId &&
                    observation.RevisionId == "master" &&
                    observation.Provenance.Address == "B2" &&
                    observation.Value?.CanonicalText == "3"));
            Assert.Equal(formula.LogicalSubjectKey, cached.LogicalSubjectKey);
            Assert.Equal(formula.LogicalSubjectKey, formulaCellValue.LogicalSubjectKey);
        }

        [Fact]
        public void Projected_evidence_flows_through_decisions_into_exact_merge_plan()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Formula"), Text("C1", "Amount"),
                Text("A2", "A"), Formula("B2", "1+1", "2"), Number("C2", "10")));
            OpenXmlWorkbookSnapshot middle = Workbook("middle", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Formula"), Text("C1", "Amount"),
                Text("A2", "A"), Formula("B2", "1+2", "3"), Number("C2", "11")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Formula"), Text("C1", "Amount"),
                Text("A2", "A"), Number("B2", "3"), Number("C2", "12")));
            OpenXmlComparisonProjectResult result = Build(
                "project-merge-plan",
                before,
                middle,
                master);
            ReviewDecisionCatalog catalog = ReviewDecisionCatalogProjector.Project(
                result,
                Evidence(result));
            ReviewDecisionSubject formula = Assert.Single(
                catalog.Subjects,
                item => item.Aspect == ComparisonAspect.Formula);
            ReviewDecisionSubject amount = Assert.Single(
                catalog.Subjects,
                item => item.Aspect == ComparisonAspect.Value &&
                        catalog.Observations.Any(observation =>
                            observation.SubjectId == item.SubjectId &&
                            observation.RevisionId == "master" &&
                            observation.Value?.CanonicalText == "12"));
            Observation selectedFormula = catalog.Observations.Single(
                item => item.SubjectId == formula.SubjectId &&
                        item.RevisionId == "before");
            Observation selectedAmount = catalog.Observations.Single(
                item => item.SubjectId == amount.SubjectId &&
                        item.RevisionId == "middle");

            ReviewDecisionGeneration generation = ReviewDecisionEngine.Create(catalog);
            generation = ReviewDecisionEngine.Apply(
                generation,
                ReviewDecisionRequest.ForSubject(
                    generation.ProjectId,
                    generation.BaseGenerationId,
                    generation.MappingGenerationId,
                    generation.GenerationId,
                    formula.SubjectId,
                    ReviewDecisionStatus.UseCandidate,
                    ReviewerSnapshot.Empty,
                    new DateTimeOffset(2026, 7, 23, 12, 0, 0, TimeSpan.Zero),
                    selectedFormula.ObservationId));
            generation = ReviewDecisionEngine.Apply(
                generation,
                ReviewDecisionRequest.ForSubject(
                    generation.ProjectId,
                    generation.BaseGenerationId,
                    generation.MappingGenerationId,
                    generation.GenerationId,
                    amount.SubjectId,
                    ReviewDecisionStatus.UseCandidate,
                    ReviewerSnapshot.Empty,
                    new DateTimeOffset(2026, 7, 23, 12, 1, 0, TimeSpan.Zero),
                    selectedAmount.ObservationId));

            MergePlan plan = MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(
                    generation,
                    new[] { amount.SubjectId, formula.SubjectId }));

            MergePlanOperation formulaOperation = Assert.Single(
                plan.Operations,
                item => item.SubjectId == formula.SubjectId);
            Assert.Equal(ObservationExactType.Blank, formulaOperation.MasterPrecondition.ExactType);
            Assert.Equal(ObservationExactType.Formula, formulaOperation.SelectedOutput.ExactType);
            Assert.Equal("1+1", formulaOperation.SelectedOutput.Formula);
            Assert.Equal("B2", formulaOperation.Target.Address);
            Assert.Equal("B2", formulaOperation.SelectedSource.Address);
            Assert.Equal("Data", formulaOperation.SelectedSource.SheetName);
            Assert.Equal(
                result.Revisions.Single(item => item.Source.RevisionId == "before")
                    .Source.Snapshot.ByteContentId,
                formulaOperation.SelectedRevisionByteContentId);

            MergePlanOperation valueOperation = Assert.Single(
                plan.Operations,
                item => item.SubjectId == amount.SubjectId);
            Assert.Equal("12", valueOperation.MasterPrecondition.Value.CanonicalText);
            Assert.Equal("11", valueOperation.SelectedOutput.Value.CanonicalText);
            Assert.Equal("C2", valueOperation.Target.Address);

            MergeCellWriteSynthesis synthesis =
                MergeCellWriteSynthesizer.Synthesize(plan);
            MergeCellWrite formulaWrite = Assert.Single(
                synthesis.Writes,
                item => item.Target.Address == "B2");
            Assert.Equal(MergeCellWriteKind.SetFormula, formulaWrite.Kind);
            Assert.Equal(
                MergeCellWriteProofRequirement.MasterConstantValueIsBlank,
                formulaWrite.ProofRequirements);
            MergeCellWrite valueWrite = Assert.Single(
                synthesis.Writes,
                item => item.Target.Address == "C2");
            Assert.Equal(MergeCellWriteKind.SetConstant, valueWrite.Kind);
            Assert.Equal("11", valueWrite.OutputValue.CanonicalText);
        }

        [Fact]
        public void Missing_master_observation_keeps_logical_provenance_without_inventing_a_location()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10"),
                Text("A3", "B"), Text("B3", "20")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10")));
            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                "project-deletion",
                "mapping-deletion",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("master", "before", 1, master, isMaster: true),
                });
            Assert.True(result.IsComplete);

            ReviewDecisionCatalog catalog = ReviewDecisionCatalogProjector.Project(result, Evidence(result));
            ReviewDecisionSubject removedValue = Assert.Single(catalog.Subjects, subject =>
                subject.Aspect == ComparisonAspect.Value &&
                catalog.Observations.Any(item =>
                    item.SubjectId == subject.SubjectId && item.RevisionId == "before" &&
                    item.DisplayText == "20"));
            Assert.Equal(ReviewDecisionSubjectPolicy.Selectable, removedValue.Policy);
            Assert.True(removedValue.CanSelectObservation);
            Assert.False(removedValue.CanMergeSelection);

            Observation beforeObservation = Assert.Single(catalog.Observations, item =>
                item.SubjectId == removedValue.SubjectId && item.RevisionId == "before");
            Observation missing = Assert.Single(catalog.Observations, item =>
                item.SubjectId == removedValue.SubjectId && item.RevisionId == "master");
            Assert.Equal("B3", beforeObservation.Provenance.Address);
            Assert.Equal(ObservationExactType.Missing, missing.ExactType);
            Assert.Equal("master.xlsx", missing.Provenance.SourceLabel);
            Assert.Equal(beforeObservation.Provenance.LogicalSheetId, missing.Provenance.LogicalSheetId);
            Assert.NotEmpty(missing.Provenance.LogicalSheetId);
            Assert.Empty(missing.Provenance.SheetName);
            Assert.Empty(missing.Provenance.Address);
            Assert.Equal(0, missing.Provenance.PhysicalRow);
            Assert.Equal(0, missing.Provenance.PhysicalColumn);
        }

        [Fact]
        public void Rejects_incomplete_projects_and_missing_or_inconsistent_revision_evidence()
        {
            OpenXmlWorkbookSnapshot before = Workbook("before", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"),
                Text("A2", "A"), Text("B2", "10")));
            OpenXmlWorkbookSnapshot ambiguous = Workbook("master",
                Sheet("Candidate A", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10")),
                Sheet("Candidate B", "DataCode",
                    Text("A1", "ID"), Text("B1", "Amount"),
                    Text("A2", "A"), Text("B2", "10")));
            OpenXmlComparisonProjectResult blocked = OpenXmlComparisonProjectEngine.Build(
                "project-blocked",
                "mapping-blocked",
                new[]
                {
                    new OpenXmlProjectRevision("before", "", 0, before, isTrustedBaseline: true),
                    new OpenXmlProjectRevision("master", "before", 1, ambiguous, isMaster: true),
                });
            Assert.False(blocked.IsComplete);
            Assert.Throws<InvalidOperationException>(() =>
                ReviewDecisionCatalogProjector.Project(blocked, Evidence(blocked)));

            OpenXmlComparisonProjectResult valid = InsertedRowProject();
            Assert.Throws<ArgumentException>(() => ReviewDecisionCatalogProjector.Project(
                valid,
                Evidence(valid).Take(valid.Revisions.Count - 1)));

            List<ReviewDecisionRevisionEvidence> inconsistent = Evidence(valid).ToList();
            inconsistent[0] = new ReviewDecisionRevisionEvidence(
                inconsistent[0].RevisionId,
                new string('A', 64),
                inconsistent[0].SourceLabel);
            Assert.Throws<InvalidOperationException>(() =>
                ReviewDecisionCatalogProjector.Project(valid, inconsistent));
        }

        [Fact]
        public void Revision_evidence_trims_a_portable_basename()
        {
            var evidence = new ReviewDecisionRevisionEvidence(
                "revision",
                new string('A', 64),
                "  Reviewed copy r2.xlsx  ");

            Assert.Equal("Reviewed copy r2.xlsx", evidence.SourceLabel);
        }

        [Theory]
        [InlineData(@"C:\private\report.xlsx")]
        [InlineData(@"C:report.xlsx")]
        [InlineData("/srv/private/report.xlsx")]
        [InlineData("https://user:secret@example.com/report.xlsx")]
        [InlineData("report.xlsx?sig=secret")]
        [InlineData("report.xlsx#private")]
        [InlineData("report\nprivate.xlsx")]
        [InlineData("report*.xlsx")]
        [InlineData(".")]
        [InlineData("..")]
        public void Revision_evidence_rejects_paths_uris_controls_and_ambiguous_labels(
            string sourceLabel)
        {
            Assert.Throws<ArgumentException>(() => new ReviewDecisionRevisionEvidence(
                "revision",
                new string('A', 64),
                sourceLabel));
        }

        [Fact]
        public void Projector_honors_cancellation_before_and_during_history_validation()
        {
            OpenXmlComparisonProjectResult result = InsertedRowProject();
            using (var canceled = new CancellationTokenSource())
            {
                canceled.Cancel();
                Assert.Throws<OperationCanceledException>(() =>
                    ReviewDecisionCatalogProjector.Project(
                        result,
                        Evidence(result),
                        canceled.Token));
            }

            using (var canceledAfterEvidence = new CancellationTokenSource())
            {
                IEnumerable<ReviewDecisionRevisionEvidence> EvidenceThenCancel()
                {
                    foreach (ReviewDecisionRevisionEvidence item in Evidence(result))
                        yield return item;
                    canceledAfterEvidence.Cancel();
                }

                Assert.Throws<OperationCanceledException>(() =>
                    ReviewDecisionCatalogProjector.Project(
                        result,
                        EvidenceThenCancel(),
                        canceledAfterEvidence.Token));
            }
        }

        private static OpenXmlComparisonProjectResult InsertedRowProject()
        {
            OpenXmlWorkbookSnapshot r1 = Workbook("r1", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"), Text("C1", "Mirror"),
                Text("A2", "A"), Text("B2", "10"), Text("C2", "10"),
                Text("A3", "B"), Text("B3", "20"), Text("C3", "20")));
            OpenXmlWorkbookSnapshot r2 = Workbook("r2", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"), Text("C1", "Mirror"),
                Text("A2", "A"), Text("B2", "10"), Text("C2", "10"),
                Text("A3", "X"), Text("B3", "15"), Text("C3", "15"),
                Text("A4", "B"), Text("B4", "20"), Text("C4", "20")));
            OpenXmlWorkbookSnapshot master = Workbook("master", Sheet("Data", "DataCode",
                Text("A1", "ID"), Text("B1", "Amount"), Text("C1", "Mirror"),
                Text("A2", "A"), Text("B2", "10"), Text("C2", "10"),
                Text("A3", "X"), Text("B3", "15"), Text("C3", "15"),
                Text("A4", "B"), Text("B4", "25"), Text("C4", "25")));
            return Build("project-all-occurrences", r1, r2, master, "r1", "r2");
        }

        private static OpenXmlComparisonProjectResult Build(
            string projectId,
            OpenXmlWorkbookSnapshot first,
            OpenXmlWorkbookSnapshot second,
            OpenXmlWorkbookSnapshot master,
            string firstId = "before",
            string secondId = "middle")
        {
            OpenXmlComparisonProjectResult result = OpenXmlComparisonProjectEngine.Build(
                projectId,
                "mapping-" + projectId,
                new[]
                {
                    new OpenXmlProjectRevision(
                        firstId,
                        "",
                        0,
                        first,
                        isTrustedBaseline: true),
                    new OpenXmlProjectRevision(secondId, firstId, 1, second),
                    new OpenXmlProjectRevision("master", secondId, 2, master, isMaster: true),
                },
                new OpenXmlComparisonProjectOptions
                {
                    ExactComparison = new ExactComparisonOptions
                    {
                        CompareCachedResults = true,
                        CompareCachedResultsWhenFormulaChanges = true,
                    },
                });
            Assert.True(result.IsComplete);
            Assert.NotNull(result.History);
            return result;
        }

        private static IEnumerable<ReviewDecisionRevisionEvidence> Evidence(
            OpenXmlComparisonProjectResult result) =>
            result.Revisions.Select(item => new ReviewDecisionRevisionEvidence(
                item.Source.RevisionId,
                item.Source.Snapshot.ByteContentId,
                item.Source.RevisionId + ".xlsx"));

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

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell Formula(string reference, string formula, string cachedValue) =>
            new Cell(new CellFormula(formula), new CellValue(cachedValue))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };
    }
}
