using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Merge;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class MergeOpenXmlMasterPreflightTests : IDisposable
    {
        private readonly string _root;

        public MergeOpenXmlMasterPreflightTests()
        {
            _root = Path.Combine(
                Path.GetTempPath(),
                "ExcelPowerTools-master-preflight-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);
        }

        [Fact]
        public void Retains_exact_master_proves_value_precondition_and_copies_owned_bytes()
        {
            string path = WorkbookPath("master.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            byte[] exactBytes = File.ReadAllBytes(path);
            PlanFixture fixture = ValuePlan(
                StableId.ForBytes(exactBytes),
                "A1",
                ObservationExactType.Number,
                "1",
                ObservationExactType.Number,
                "2");

            using MergeOpenXmlMasterPreflight preflight =
                MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                    path,
                    fixture.Plan,
                    fixture.Synthesis);

            Assert.True(preflight.IsLeaseActive);
            Assert.Equal(Path.GetFullPath(path), preflight.SourcePath);
            Assert.Equal(exactBytes.Length, preflight.ByteLength);
            Assert.Equal(fixture.Plan.MasterByteContentId, preflight.ByteContentId);
            Assert.Equal(
                MergePlanRequiredValidation.SourceMasterByteIdentity,
                preflight.SourceIdentityProof.Check);
            Assert.Equal(
                MergePlanRequiredValidation.MasterTargetPreconditions,
                preflight.TargetPreconditionsProof.Check);
            MergeMasterTargetObservation observed =
                Assert.Single(preflight.TargetObservations);
            Assert.False(observed.HasFormula);
            Assert.Equal(ObservationExactType.Number, observed.ConstantValue.Type);
            Assert.Equal("1", observed.ConstantValue.CanonicalText);

            using var copied = new MemoryStream();
            preflight.CopyMasterBytesTo(copied);
            Assert.Equal(exactBytes, copied.ToArray());

            Assert.Throws<IOException>(() =>
            {
                using var writer = new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Write,
                    FileShare.ReadWrite | FileShare.Delete);
            });
        }

        [Fact]
        public void Retained_formula_precondition_is_exact()
        {
            string path = WorkbookPath("formula.xlsx");
            CreateWorkbook(path, CellFormula("A1", "OLD()", "1"));
            PlanFixture fixture = FormulaPlan(
                HashFile(path),
                "A1",
                ObservationExactType.Formula,
                "OLD()",
                ObservationExactType.Formula,
                "NEW()");

            using MergeOpenXmlMasterPreflight preflight =
                MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                    path,
                    fixture.Plan,
                    fixture.Synthesis);

            MergeMasterTargetObservation observed =
                Assert.Single(preflight.TargetObservations);
            Assert.True(observed.HasFormula);
            Assert.Equal("OLD()", observed.Formula);
            Assert.Equal(ObservationExactType.Blank, observed.ConstantValue.Type);
        }

        [Fact]
        public void Disposed_lease_cannot_copy_base_bytes()
        {
            string path = WorkbookPath("disposed.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            PlanFixture fixture = ValuePlan(
                HashFile(path),
                "A1",
                ObservationExactType.Number,
                "1",
                ObservationExactType.Number,
                "2");
            MergeOpenXmlMasterPreflight preflight =
                MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                    path,
                    fixture.Plan,
                    fixture.Synthesis);

            preflight.Dispose();

            Assert.False(preflight.IsLeaseActive);
            MergeOpenXmlMasterPreflightException exception =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    preflight.CopyMasterBytesTo(new MemoryStream()));
            Assert.Equal(MergeOpenXmlMasterPreflightCodes.LeaseDisposed, exception.Code);
        }

        [Fact]
        public void Rejects_master_with_wrong_byte_identity()
        {
            string path = WorkbookPath("changed.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            PlanFixture fixture = ValuePlan(
                Hash("different-master-bytes"),
                "A1",
                ObservationExactType.Number,
                "1",
                ObservationExactType.Number,
                "2");

            MergeOpenXmlMasterPreflightException exception =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                        path,
                        fixture.Plan,
                        fixture.Synthesis));

            Assert.Equal(
                MergeOpenXmlMasterPreflightCodes.ByteIdentityMismatch,
                exception.Code);
            Assert.Equal(Path.GetFullPath(path), exception.SourcePath);
        }

        [Fact]
        public void Rejects_exact_bytes_when_planned_value_precondition_is_forged()
        {
            string path = WorkbookPath("precondition.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            PlanFixture fixture = ValuePlan(
                HashFile(path),
                "A1",
                ObservationExactType.Number,
                "99",
                ObservationExactType.Number,
                "2");

            MergeOpenXmlMasterPreflightException exception =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                        path,
                        fixture.Plan,
                        fixture.Synthesis));

            Assert.Equal(
                MergeOpenXmlMasterPreflightCodes.PreconditionMismatch,
                exception.Code);
            Assert.NotEmpty(exception.TargetId);
        }

        [Fact]
        public void Rejects_formula_addition_that_would_discard_unselected_constant()
        {
            string path = WorkbookPath("constant-to-formula.xlsx");
            CreateWorkbook(path, CellNumber("A1", "10"));
            PlanFixture fixture = FormulaPlan(
                HashFile(path),
                "A1",
                ObservationExactType.Blank,
                "",
                ObservationExactType.Formula,
                "CHOSEN()");
            MergeCellWrite write = Assert.Single(fixture.Synthesis.Writes);
            Assert.Equal(
                MergeCellWriteProofRequirement.MasterConstantValueIsBlank,
                write.ProofRequirements);

            MergeOpenXmlMasterPreflightException exception =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                        path,
                        fixture.Plan,
                        fixture.Synthesis));

            Assert.Equal(
                MergeOpenXmlMasterPreflightCodes.ProofRequirementFailed,
                exception.Code);
        }

        [Fact]
        public void Rejects_value_only_write_that_would_remove_unselected_formula()
        {
            string path = WorkbookPath("formula-to-value.xlsx");
            CreateWorkbook(path, CellFormula("A1", "OLD()", "1"));
            PlanFixture fixture = ValuePlan(
                HashFile(path),
                "A1",
                ObservationExactType.Blank,
                "",
                ObservationExactType.Number,
                "2");
            MergeCellWrite write = Assert.Single(fixture.Synthesis.Writes);
            Assert.Equal(
                MergeCellWriteProofRequirement.MasterHasNoFormula,
                write.ProofRequirements);

            MergeOpenXmlMasterPreflightException exception =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                        path,
                        fixture.Plan,
                        fixture.Synthesis));

            Assert.Equal(
                MergeOpenXmlMasterPreflightCodes.ProofRequirementFailed,
                exception.Code);
        }

        [Fact]
        public void Rejects_missing_physical_target_even_when_bytes_match_plan()
        {
            string path = WorkbookPath("missing-target.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            PlanFixture fixture = ValuePlan(
                HashFile(path),
                "A2",
                ObservationExactType.Number,
                "1",
                ObservationExactType.Number,
                "2");

            MergeOpenXmlMasterPreflightException exception =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                        path,
                        fixture.Plan,
                        fixture.Synthesis));

            Assert.Equal(MergeOpenXmlMasterPreflightCodes.TargetMissing, exception.Code);
        }

        [Fact]
        public void Source_and_target_proofs_complete_their_exact_receipt_slots()
        {
            string path = WorkbookPath("receipt.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            PlanFixture fixture = ValuePlan(
                HashFile(path),
                "A1",
                ObservationExactType.Number,
                "1",
                ObservationExactType.Number,
                "2");
            using MergeOpenXmlMasterPreflight preflight =
                MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                    path,
                    fixture.Plan,
                    fixture.Synthesis);
            var proofs = new List<MergePlanValidationProof>();
            foreach (MergePlanRequiredValidation check in
                     MergeValidationReceiptPolicy.EnumerateChecks(
                         fixture.Plan.RequiredValidations))
            {
                switch (check)
                {
                    case MergePlanRequiredValidation.SourceMasterByteIdentity:
                        proofs.Add(preflight.SourceIdentityProof);
                        break;
                    case MergePlanRequiredValidation.MasterTargetPreconditions:
                        proofs.Add(preflight.TargetPreconditionsProof);
                        break;
                    case MergePlanRequiredValidation.CellWriteSynthesis:
                        proofs.Add(MergePlanValidationVerifier.ProveCellWriteSynthesis(
                            fixture.Plan,
                            fixture.Synthesis));
                        break;
                    case MergePlanRequiredValidation.FrozenDecisionAndMappingIdentity:
                        proofs.Add(MergePlanValidationVerifier.ProveFrozenContext(
                            fixture.Plan,
                            fixture.Synthesis,
                            fixture.Generation));
                        break;
                    default:
                        proofs.Add(MergePlanValidationProof.CreateTrusted(
                            fixture.Plan,
                            fixture.Synthesis,
                            check,
                            "test-" + check + "-v1",
                            Hash("proof-" + check)));
                        break;
                }
            }

            MergePlanValidationReceipt receipt =
                MergePlanValidationReceipt.Assemble(
                    fixture.Plan,
                    fixture.Synthesis,
                    proofs);

            Assert.Equal(receipt.RequiredValidations, receipt.CoveredValidations);
            Assert.Contains(
                receipt.Proofs,
                item => item.ProofHash ==
                        preflight.SourceIdentityProof.ProofHash);
            Assert.Contains(
                receipt.Proofs,
                item => item.ProofHash ==
                        preflight.TargetPreconditionsProof.ProofHash);
        }

        [Fact]
        public void Invalid_path_and_cancellation_fail_before_a_lease_escapes()
        {
            string path = WorkbookPath("valid.xlsx");
            CreateWorkbook(path, CellNumber("A1", "1"));
            PlanFixture fixture = ValuePlan(
                HashFile(path),
                "A1",
                ObservationExactType.Number,
                "1",
                ObservationExactType.Number,
                "2");

            MergeOpenXmlMasterPreflightException invalid =
                Assert.Throws<MergeOpenXmlMasterPreflightException>(() =>
                    MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                        "",
                        fixture.Plan,
                        fixture.Synthesis));
            Assert.Equal(MergeOpenXmlMasterPreflightCodes.InvalidPath, invalid.Code);

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() =>
                MergeOpenXmlMasterPreflightVerifier.OpenAndProve(
                    path,
                    fixture.Plan,
                    fixture.Synthesis,
                    cancellation.Token));
        }

        public void Dispose()
        {
            if (Directory.Exists(_root))
                Directory.Delete(_root, recursive: true);
        }

        private PlanFixture ValuePlan(
            string masterByteContentId,
            string address,
            ObservationExactType masterType,
            string masterText,
            ObservationExactType candidateType,
            string candidateText) =>
            Plan(
                masterByteContentId,
                address,
                ComparisonAspect.Value,
                masterType,
                masterText,
                candidateType,
                candidateText);

        private PlanFixture FormulaPlan(
            string masterByteContentId,
            string address,
            ObservationExactType masterType,
            string masterText,
            ObservationExactType candidateType,
            string candidateText) =>
            Plan(
                masterByteContentId,
                address,
                ComparisonAspect.Formula,
                masterType,
                masterText,
                candidateType,
                candidateText);

        private PlanFixture Plan(
            string masterByteContentId,
            string address,
            ComparisonAspect aspect,
            ObservationExactType masterType,
            string masterText,
            ObservationExactType candidateType,
            string candidateText)
        {
            var subject = new ReviewDecisionSubject(
                "subject",
                "logical-cell",
                aspect,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            Observation master = Observation(
                aspect,
                "master",
                masterType,
                masterText,
                address);
            Observation candidate = Observation(
                aspect,
                "candidate",
                candidateType,
                candidateText,
                address);
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
                        masterByteContentId),
                });
            var catalog = new ReviewDecisionCatalog(
                evidence,
                "master",
                new[] { subject },
                new[] { candidate, master });
            ReviewDecisionGeneration generation =
                ReviewDecisionEngine.Create(catalog);
            generation = ReviewDecisionEngine.Apply(
                generation,
                ReviewDecisionRequest.ForSubject(
                    generation.ProjectId,
                    generation.BaseGenerationId,
                    generation.MappingGenerationId,
                    generation.GenerationId,
                    subject.SubjectId,
                    ReviewDecisionStatus.UseCandidate,
                    ReviewerSnapshot.Empty,
                    new DateTimeOffset(2026, 7, 23, 15, 0, 0, TimeSpan.Zero),
                    candidate.ObservationId));
            MergePlan plan = MergePlanBuilder.Build(
                generation,
                MergePlanRequest.Capture(generation, new[] { subject.SubjectId }));
            return new PlanFixture(
                generation,
                plan,
                MergeCellWriteSynthesizer.Synthesize(plan));
        }

        private static Observation Observation(
            ComparisonAspect aspect,
            string revisionId,
            ObservationExactType type,
            string text,
            string address)
        {
            string contentHash = Hash(aspect + ":" + type + ":" + text);
            CanonicalScalar? value = null;
            string? formula = null;
            if (type == ObservationExactType.Formula)
                formula = text;
            else if (type == ObservationExactType.Blank)
                value = CanonicalScalar.Blank;
            else
                value = new CanonicalScalar(type, text);
            return new Observation(
                ComparisonIds.Observation(
                    "subject",
                    revisionId,
                    type,
                    contentHash),
                "subject",
                revisionId,
                type,
                contentHash,
                formula,
                value,
                provenance: new ObservationProvenance(
                    revisionId + ".xlsx",
                    "Data",
                    address,
                    ParseRow(address),
                    ParseColumn(address),
                    "logical-sheet"));
        }

        private static void CreateWorkbook(string path, Cell cell)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Create(
                path,
                SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            string address = cell.CellReference?.Value ??
                             throw new ArgumentException(
                                 "The test cell requires a reference.",
                                 nameof(cell));
            var row = new Row { RowIndex = (uint)ParseRow(address) };
            row.Append(cell);
            worksheetPart.Worksheet = new Worksheet(new SheetData(row));
            var sheets = new Sheets();
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Data",
            });
            workbookPart.Workbook.Append(sheets);
            workbookPart.Workbook.Save();
            worksheetPart.Worksheet.Save();
        }

        private static Cell CellNumber(string address, string value) =>
            new Cell
            {
                CellReference = address,
                CellValue = new CellValue(value),
            };

        private static Cell CellFormula(
            string address,
            string formula,
            string cachedValue) =>
            new Cell
            {
                CellReference = address,
                CellFormula = new CellFormula(formula),
                CellValue = new CellValue(cachedValue),
            };

        private string WorkbookPath(string leaf) => Path.Combine(_root, leaf);

        private static string HashFile(string path) =>
            StableId.ForBytes(File.ReadAllBytes(path));

        private static string Hash(string value) =>
            StableId.ForBytes(Encoding.UTF8.GetBytes(value));

        private static int ParseRow(string address) =>
            int.Parse(new string(address.Where(char.IsDigit).ToArray()));

        private static int ParseColumn(string address)
        {
            int column = 0;
            foreach (char character in address.TakeWhile(char.IsLetter))
                column = checked(column * 26 + (char.ToUpperInvariant(character) - 'A' + 1));
            return column;
        }

        private sealed class PlanFixture
        {
            internal PlanFixture(
                ReviewDecisionGeneration generation,
                MergePlan plan,
                MergeCellWriteSynthesis synthesis)
            {
                Generation = generation;
                Plan = plan;
                Synthesis = synthesis;
            }

            internal ReviewDecisionGeneration Generation { get; }
            internal MergePlan Plan { get; }
            internal MergeCellWriteSynthesis Synthesis { get; }
        }
    }
}
