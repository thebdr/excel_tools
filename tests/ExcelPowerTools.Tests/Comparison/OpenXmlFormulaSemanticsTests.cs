using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Excel.DynamicArray;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using Xunit;
using SpreadsheetExtension = DocumentFormat.OpenXml.Spreadsheet.Extension;
using SpreadsheetExtensionList = DocumentFormat.OpenXml.Spreadsheet.ExtensionList;

namespace ExcelPowerTools.Tests.Comparison
{
    public class OpenXmlFormulaSemanticsTests
    {
        [Fact]
        public void Shared_A1_follower_reconstruction_preserves_lexical_text_and_translates_proven_references()
        {
            const string masterText =
                "A1+$A1+A$1+$A$1+SUM(A:C)+SUM($D:$E)+SUM(1:3)+SUM($4:$5)+\"A1\"+'Local Sheet'!A1";
            const string expectedFollower =
                "B2+$A2+B$1+$A$1+SUM(B:D)+SUM($D:$E)+SUM(2:4)+SUM($4:$5)+\"A1\"+'Local Sheet'!B2";
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("B2", 7U, "B2:C3", masterText, "1"),
                    SharedFollower("C2", 7U, "2"),
                    SharedFollower("B3", 7U, "2"),
                    SharedFollower("C3", 7U, "2")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);

            Assert.Equal(masterText, sheet.Snapshot.Cells["B2"].FormulaText);
            Assert.Equal(expectedFollower, sheet.Snapshot.Cells["C3"].FormulaText);
            Assert.DoesNotContain(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
            Assert.Equal(4, sheet.FormulaObservations.Count);
            OpenXmlFormulaObservation master =
                Assert.Single(sheet.FormulaObservations, item => item.CellReference == "B2");
            Assert.Equal(OpenXmlFormulaForm.SharedMaster, master.Form);
            Assert.Equal(masterText, master.StoredFormulaText);
            Assert.Equal(masterText, master.EffectiveFormulaText);
            Assert.False(master.WasReconstructed);
            Assert.Equal("B2:C3", master.FormulaRange);
            Assert.Equal(7U, master.SharedIndex);

            OpenXmlFormulaObservation follower =
                Assert.Single(sheet.FormulaObservations, item => item.CellReference == "C3");
            Assert.Equal(OpenXmlFormulaForm.SharedFollower, follower.Form);
            Assert.Equal("", follower.StoredFormulaText);
            Assert.Equal(expectedFollower, follower.EffectiveFormulaText);
            Assert.True(follower.WasReconstructed);
            Assert.NotNull(follower.ExactMetadata);
        }

        [Fact]
        public void Shared_translation_does_not_rewrite_A1_shaped_text_inside_Unicode_defined_names()
        {
            const string masterText = "αA1+A1+A1β+𐐀A1";
            const string followerText = "αA1+B1+A1β+𐐀A1";
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("B1", 11U, "B1:C1", masterText, "1"),
                    SharedFollower("C1", 11U, "2")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);

            Assert.Equal(followerText, sheet.Snapshot.Cells["C1"].FormulaText);
            Assert.DoesNotContain(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
        }

        [Theory]
        [InlineData("XFD1", "XFC1", "XFC1:XFD1")]
        [InlineData("R1C1+1", "A1", "A1:A2")]
        [InlineData("RC+1", "A1", "A1:A2")]
        [InlineData("R1C+1", "A1", "A1:A2")]
        [InlineData("RC1+1", "A1", "A1:A2")]
        [InlineData("r1c+1", "A1", "A1:A2")]
        [InlineData("SUM(Table1[Amount])", "A1", "A1:A2")]
        [InlineData("\"unterminated", "A1", "A1:A2")]
        public void Shared_formula_follower_fails_closed_when_lexical_translation_is_not_proven(
            string formula,
            string masterReference,
            string range)
        {
            string followerReference = masterReference == "XFC1" ? "XFD1" : "A2";
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster(masterReference, 3U, range, formula, "1"),
                    SharedFollower(followerReference, 3U, "2")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);
            OpenXmlFormulaObservation follower =
                Assert.Single(sheet.FormulaObservations, item => item.CellReference == followerReference);

            Assert.Null(follower.EffectiveFormulaText);
            Assert.False(follower.WasReconstructed);
            Assert.StartsWith(
                "#OPENXML-COVERAGE[" + OpenXmlCoverageCodes.SharedFormulaNotExpanded,
                sheet.Snapshot.Cells[followerReference].FormulaText,
                StringComparison.Ordinal);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                        item.CellReference == followerReference &&
                        item.AffectsExactCellComparison);
        }

        [Theory]
        [InlineData("S1:S3!A1")]
        [InlineData("A:C!A1")]
        public void Shared_translation_rejects_unquoted_sheet_and_3D_qualifiers(string masterText)
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 14U, "A1:B1", masterText, "1"),
                    SharedFollower("B1", 14U, "2")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);
            OpenXmlFormulaObservation follower =
                Assert.Single(sheet.FormulaObservations, item => item.CellReference == "B1");

            Assert.Null(follower.EffectiveFormulaText);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                        item.CellReference == "B1" &&
                        item.Message.Contains("qualifier", StringComparison.OrdinalIgnoreCase));
        }

        [Theory]
        [InlineData("Sheet2!A1", "Sheet2!B1")]
        [InlineData("'[1]Sheet'!A1", "'[1]Sheet'!B1")]
        public void Shared_translation_copies_single_sheet_qualifiers_and_translates_only_the_reference(
            string masterText,
            string followerText)
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 15U, "A1:B1", masterText, "1"),
                    SharedFollower("B1", 15U, "2")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);

            Assert.Equal(followerText, sheet.Snapshot.Cells["B1"].FormulaText);
            Assert.DoesNotContain(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
        }

        [Fact]
        public void Missing_and_duplicate_shared_masters_are_explicit_coverage_not_guessed_formulas()
        {
            byte[] missingBytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(SharedFollower("A2", 1U, "2")),
            });
            byte[] duplicateBytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 2U, "A1:A2", "A1+1", "1"),
                    SharedMaster("B1", 2U, "B1:B2", "B1+1", "1")),
            });

            OpenXmlWorksheetSnapshot missing = Read(missingBytes);
            OpenXmlWorksheetSnapshot duplicate = Read(duplicateBytes);

            Assert.All(missing.FormulaObservations, observation => Assert.Null(observation.EffectiveFormulaText));
            Assert.All(duplicate.FormulaObservations, observation => Assert.Null(observation.EffectiveFormulaText));
            Assert.Contains(
                missing.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
            Assert.Equal(
                2,
                duplicate.CoverageMessages.Count(
                    item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded));
        }

        [Fact]
        public void Repeated_formula_coverage_is_bounded_with_representatives_and_a_counted_summary()
        {
            var cells = new List<Cell>();
            for (uint row = 1; row <= 20; row++)
                cells.Add(SharedFollower("A" + row, row, "1"));

            OpenXmlWorksheetSnapshot sheet = Read(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(cells.ToArray()) }));
            List<OpenXmlCoverageMessage> findings = sheet.CoverageMessages
                .Where(item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded)
                .ToList();

            Assert.Equal(9, findings.Count);
            Assert.Equal(8, findings.Count(item => item.CellReference != null));
            Assert.Contains(
                findings,
                item => item.CellReference == null &&
                        item.Message.Contains("12 additional", StringComparison.Ordinal));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Shared_follower_with_text_or_ref_rejects_the_entire_group(bool storedText)
        {
            var followerFormula = new CellFormula(storedText ? "C2+999" : "")
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 12U,
            };
            if (!storedText) followerFormula.Reference = "A1:A2";
            var follower = new Cell(followerFormula, new CellValue("2"))
            {
                CellReference = "A2",
                DataType = CellValues.Number,
            };
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 12U, "A1:A2", "B1+1", "1"),
                    follower),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);

            Assert.All(sheet.FormulaObservations, item => Assert.Null(item.EffectiveFormulaText));
            Assert.False(sheet.CanAssertExactCellEquality);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                        item.Message.Contains("empty-text/no-ref", StringComparison.Ordinal));
            if (storedText)
                Assert.Equal("C2+999", sheet.Snapshot.Cells["A2"].FormulaText);
        }

        [Fact]
        public void Out_of_range_shared_follower_blocks_the_entire_ambiguous_group()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 4U, "A1:A2", "B1+1", "1"),
                    SharedFollower("A3", 4U, "3")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);
            OpenXmlFormulaObservation master =
                Assert.Single(sheet.FormulaObservations, item => item.CellReference == "A1");
            OpenXmlFormulaObservation follower =
                Assert.Single(sheet.FormulaObservations, item => item.CellReference == "A3");

            Assert.Null(master.EffectiveFormulaText);
            Assert.Null(follower.EffectiveFormulaText);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                        item.CellReference == "A1" &&
                        item.Message.Contains("outside", StringComparison.OrdinalIgnoreCase));
        }

        [Theory]
        [InlineData("hole")]
        [InlineData("constant")]
        [InlineData("normal-formula")]
        [InlineData("foreign-si")]
        public void Shared_range_reconstructs_only_actual_members_and_permits_other_content_inside_its_envelope(
            string middleKind)
        {
            var cells = new List<Cell>
            {
                SharedMaster("A1", 8U, "A1:A3", "B1+1", "1"),
                SharedFollower("A3", 8U, "3"),
            };
            switch (middleKind)
            {
                case "constant":
                    cells.Add(Number("A2", "2"));
                    break;
                case "normal-formula":
                    cells.Add(Formula("A2", "B2+1", "2", CellValues.Number));
                    break;
                case "foreign-si":
                    cells.Add(SharedMaster("A2", 9U, "A2:A2", "B2+1", "2"));
                    break;
            }

            OpenXmlWorksheetSnapshot sheet = Read(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(cells.ToArray()) }));

            Assert.Equal("B1+1", Assert.Single(
                sheet.FormulaObservations,
                item => item.CellReference == "A1").EffectiveFormulaText);
            Assert.Equal("B3+1", Assert.Single(
                sheet.FormulaObservations,
                item => item.CellReference == "A3").EffectiveFormulaText);
            Assert.DoesNotContain(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
            Assert.True(sheet.CanAssertExactCellEquality);
        }

        [Fact]
        public void Impractically_large_shared_range_fails_before_rectangle_enumeration()
        {
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(SharedMaster("A1", 10U, "A1:XFD1048576", "B1+1", "1")),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);
            OpenXmlFormulaObservation observation = Assert.Single(sheet.FormulaObservations);

            Assert.Null(observation.EffectiveFormulaText);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                        item.Message.Contains("safe reconstruction limit", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Long_shared_master_times_dense_group_fails_before_follower_text_allocation()
        {
            string masterText = string.Join("+", Enumerable.Repeat("A1", 2667));
            var cells = new List<Cell>
            {
                SharedMaster("A1", 13U, "A1:A1000", masterText, "1"),
            };
            for (int row = 2; row <= 1000; row++)
                cells.Add(SharedFollower("A" + row, 13U, row.ToString()));

            OpenXmlWorksheetSnapshot sheet = Read(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(cells.ToArray()) }));

            Assert.Equal(1000, sheet.FormulaObservations.Count);
            Assert.All(sheet.FormulaObservations, item => Assert.Null(item.EffectiveFormulaText));
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                        item.Message.Contains("derived-formula evidence limit", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Many_disjoint_single_cell_shared_groups_are_processed_without_pairwise_range_scans()
        {
            const int groupCount = 4096;
            var cells = new List<Cell>(groupCount);
            for (int row = 1; row <= groupCount; row++)
            {
                string reference = "A" + row;
                cells.Add(SharedMaster(reference, (uint)row, reference + ":" + reference, "B" + row, "1"));
            }

            OpenXmlWorksheetSnapshot sheet = Read(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(cells.ToArray()) }));

            Assert.Equal(groupCount, sheet.FormulaObservations.Count);
            Assert.All(sheet.FormulaObservations, item => Assert.NotNull(item.EffectiveFormulaText));
            Assert.DoesNotContain(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
        }

        [Fact]
        public void Shared_master_text_edit_emits_formula_changes_without_formula_structure_noise()
        {
            OpenXmlWorksheetSnapshot before = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 21U, "A1:A2", "B1+1", "1"),
                    SharedFollower("A2", 21U, "2")),
            }));
            OpenXmlWorksheetSnapshot after = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(
                    SharedMaster("A1", 21U, "A1:A2", "B1+2", "1"),
                    SharedFollower("A2", 21U, "2")),
            }));

            IReadOnlyList<AtomicCellChange> changes =
                ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot);

            Assert.Equal(2, changes.Count);
            Assert.All(changes, change => Assert.Equal(CellChangeAspect.Formula, change.Aspect));
        }

        [Fact]
        public void Unresolved_formula_tokens_ignore_physical_worksheet_part_identity()
        {
            OpenXmlWorksheetSnapshot sheet1 = ReadNamed(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    NamedSheet("Data", 1, SharedFollower("A1", 99U, "1")),
                    NamedSheet("Dummy", 2, Number("A1", "0")),
                }),
                "Data");
            OpenXmlWorksheetSnapshot sheet2 = ReadNamed(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    NamedSheet("Dummy", 2, Number("A1", "0")),
                    NamedSheet("Data", 1, SharedFollower("A1", 99U, "1")),
                }),
                "Data");

            Assert.NotEqual(sheet1.PartUri, sheet2.PartUri);
            Assert.Empty(ExactComparisonEngine.Compare(sheet1.Snapshot, sheet2.Snapshot));
            Assert.Equal(
                sheet1.Snapshot.Cells["A1"].FormulaText,
                sheet2.Snapshot.Cells["A1"].FormulaText);
        }

        [Fact]
        public void Unresolved_formula_fingerprint_survives_confirmed_alignment_to_a_new_physical_address()
        {
            OpenXmlWorksheetSnapshot before = ReadNamed(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    NamedSheet("Data", 1, SharedFollower("A1", 99U, "1")),
                    NamedSheet("Dummy", 2, Number("A1", "0")),
                }),
                "Data");
            OpenXmlWorksheetSnapshot after = ReadNamed(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    NamedSheet("Dummy", 2, Number("A1", "0")),
                    NamedSheet("Data", 1, SharedFollower("B2", 99U, "1")),
                }),
                "Data");
            CanonicalCellState beforeCell = before.Snapshot.Cells["A1"];
            CanonicalCellState afterCell = after.Snapshot.Cells["B2"];
            var beforeAligned = new RevisionSnapshot("before", new[]
            {
                CanonicalCellState.Formula(
                    "logical-row/field",
                    beforeCell.FormulaText,
                    beforeCell.CachedResult,
                    exactFormulaMetadata: beforeCell.FormulaMetadata),
            });
            var afterAligned = new RevisionSnapshot("after", new[]
            {
                CanonicalCellState.Formula(
                    "logical-row/field",
                    afterCell.FormulaText,
                    afterCell.CachedResult,
                    exactFormulaMetadata: afterCell.FormulaMetadata),
            });

            Assert.NotEqual(before.PartUri, after.PartUri);
            Assert.Empty(ExactComparisonEngine.Compare(beforeAligned, afterAligned));
            Assert.False(before.CanAssertExactCellEquality);
            Assert.False(after.CanAssertExactCellEquality);
        }

        [Fact]
        public void Unsupported_value_fingerprint_excludes_physical_address_part_and_unrelated_style()
        {
            var beforeData = new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = "Data",
                SheetId = 1,
                Cells = new[] { UnsupportedValue("A1", "payload", 7U) },
            };
            var afterData = new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = "Data",
                SheetId = 1,
                Cells = new[] { UnsupportedValue("B2", "payload", 91U) },
            };
            OpenXmlWorksheetSnapshot before = ReadNamed(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    beforeData,
                    NamedSheet("Dummy", 2, Number("A1", "0")),
                }),
                "Data");
            OpenXmlWorksheetSnapshot after = ReadNamed(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    NamedSheet("Dummy", 2, Number("A1", "0")),
                    afterData,
                }),
                "Data");

            CanonicalCellState beforeAligned = CanonicalCellState.Constant(
                "logical-row/field",
                before.Snapshot.Cells["A1"].ConstantValue);
            CanonicalCellState afterAligned = CanonicalCellState.Constant(
                "logical-row/field",
                after.Snapshot.Cells["B2"].ConstantValue);
            var beforeRevision = new RevisionSnapshot("before", new[] { beforeAligned });
            var afterRevision = new RevisionSnapshot("after", new[] { afterAligned });

            Assert.NotEqual(before.PartUri, after.PartUri);
            Assert.Empty(ExactComparisonEngine.Compare(beforeRevision, afterRevision));
            Assert.False(before.CanAssertExactCellEquality);
            Assert.False(after.CanAssertExactCellEquality);
        }

        [Fact]
        public void External_formula_text_is_exact_but_dependency_resolution_is_explicitly_out_of_scope()
        {
            const string formula = "'[1]External Sheet'!$B$2+A1";
            byte[] bytes = OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(Formula("A1", formula, "7", CellValues.Number)),
            });

            OpenXmlWorksheetSnapshot sheet = Read(bytes);

            Assert.Equal(formula, sheet.Snapshot.Cells["A1"].FormulaText);
            Assert.Equal(formula, Assert.Single(sheet.FormulaObservations).StoredFormulaText);
            OpenXmlCoverageMessage message = Assert.Single(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.ExternalFormulaDependencyNotResolved);
            Assert.False(message.AffectsExactCellComparison);
            Assert.True(sheet.CanAssertExactCellEquality);
        }

        [Fact]
        public void Array_and_dynamic_array_metadata_are_retained_without_spill_expansion()
        {
            var formula = new Cell(
                new CellFormula("_xlfn.UNIQUE(A1:A3)")
                {
                    FormulaType = CellFormulaValues.Array,
                    Reference = "B1:B3",
                },
                new CellValue("1"))
            {
                CellReference = "B1",
                CellMetaIndex = 1U,
                DataType = CellValues.Number,
            };
            byte[] bytes = AddDynamicArrayMetadata(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(formula) }));

            OpenXmlWorksheetSnapshot sheet = Read(bytes);
            OpenXmlFormulaObservation observation = Assert.Single(sheet.FormulaObservations);

            Assert.True(
                observation.Form == OpenXmlFormulaForm.DynamicArray,
                "form=" + observation.Form + "; coverage=" + string.Join(
                    " | ",
                    sheet.CoverageMessages.Select(item => item.Code + ":" + item.Message)));
            Assert.Equal("_xlfn.UNIQUE(A1:A3)", observation.StoredFormulaText);
            Assert.Equal("B1:B3", observation.FormulaRange);
            Assert.Equal(CanonicalValue.Number(1), observation.CachedResult);
            Assert.NotNull(observation.ExactMetadata);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.DynamicArrayNotExpanded &&
                        item.AffectsExactCellComparison);
            Assert.DoesNotContain(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.FormulaMetadataUnresolved);
            Assert.False(sheet.CanProjectExactStoredFormulaValueEvidence);
        }

        [Fact]
        public void Single_cell_dynamic_array_is_exact_stored_review_evidence_but_not_merge_evidence()
        {
            var formula = new Cell(
                new CellFormula("_xlfn.XLOOKUP(A1,B:B,C:C)")
                {
                    FormulaType = CellFormulaValues.Array,
                    Reference = "B1",
                },
                new CellValue("7"))
            {
                CellReference = "B1",
                CellMetaIndex = 1U,
                DataType = CellValues.Number,
            };
            OpenXmlWorksheetSnapshot sheet = Read(AddDynamicArrayMetadata(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(formula) })));

            Assert.False(sheet.CanAssertExactCellEquality);
            Assert.True(sheet.CanProjectExactStoredFormulaValueEvidence);
            Assert.Equal("B1", Assert.Single(sheet.FormulaObservations).FormulaRange);
        }

        [Fact]
        public void Independent_shared_and_dynamic_array_coverage_codes_are_both_retained()
        {
            Cell follower = SharedFollower("A1", 77U, "1");
            follower.CellMetaIndex = 1U;
            OpenXmlWorksheetSnapshot sheet = Read(AddDynamicArrayMetadata(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(follower) })));

            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded);
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.DynamicArrayNotExpanded);
        }

        [Fact]
        public void Metadata_budget_failure_is_whole_scope_and_stable_when_physical_cells_are_reversed()
        {
            Cell first = Formula("A1", "1+1", "2", CellValues.Number);
            first.CellMetaIndex = 1U;
            Cell second = Formula("A2", "1+1", "2", CellValues.Number);
            second.CellMetaIndex = 1U;
            byte[] normal = AddDynamicArrayMetadata(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(first, second) }));
            byte[] reversed = ReversePhysicalRows(normal);
            var limits = new OpenXmlFormulaMetadataLimits(
                1000,
                1024 * 1024,
                1000,
                1);

            OpenXmlWorksheetSnapshot before = Assert.Single(
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("before", normal, limits).Sheets);
            OpenXmlWorksheetSnapshot after = Assert.Single(
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("after", reversed, limits).Sheets);

            Assert.Empty(ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));
            Assert.All(
                before.FormulaObservations,
                observation => Assert.Contains(
                    before.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.FormulaMetadataUnresolved &&
                            item.CellReference == observation.CellReference &&
                            item.Message.Contains("as a whole", StringComparison.OrdinalIgnoreCase)));
            Assert.Equal(
                before.FormulaObservations.Select(item => item.ExactMetadata),
                after.FormulaObservations.Select(item => item.ExactMetadata));
        }

        [Fact]
        public void Metadata_exhaustion_does_not_materialize_unique_cm_evidence_during_fallback()
        {
            const int count = 64;
            var alphaCells = new List<Cell>(count / 2);
            var betaCells = new List<Cell>(count / 2);
            for (int row = 1; row <= count; row++)
            {
                int localRow = row <= count / 2 ? row : row - count / 2;
                Cell formula = Formula("A" + localRow, "1+1", "2", CellValues.Number);
                formula.CellMetaIndex = (uint)row;
                (row <= count / 2 ? alphaCells : betaCells).Add(formula);
            }
            byte[] bytes = AddDynamicArrayMetadata(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    NamedSheet("Alpha", 1, alphaCells.ToArray()),
                    NamedSheet("Beta", 2, betaCells.ToArray()),
                }),
                count);

            using var stream = new MemoryStream(bytes, writable: false);
            using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            var catalog = new OpenXmlFormulaMetadataCatalog(
                workbookPart,
                new OpenXmlFormulaMetadataLimits(1000, 1024 * 1024, 1000, 1));

            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
            {
                List<KeyValuePair<string, Cell>> cells = worksheetPart.Worksheet!.Descendants<Cell>()
                    .Select(cell => new KeyValuePair<string, Cell>(cell.CellReference!.Value!, cell))
                    .ToList();
                catalog.PrepareScope(cells, 1000, 1);

                Assert.Equal(0, catalog.ResolvedEvidenceCacheCountForTesting);
                Assert.All(
                    cells,
                    pair => Assert.Contains(
                        "as a whole",
                        catalog.Resolve(pair.Value, pair.Key).FailureReason,
                        StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Workbook_partitioned_metadata_budget_fails_both_sheets_independent_of_sheet_order()
        {
            byte[] ordered = TwoSheetDynamicMetadataWorkbook(false);
            byte[] swapped = TwoSheetDynamicMetadataWorkbook(true);
            var limits = new OpenXmlFormulaMetadataLimits(1000, 1024 * 1024, 1000, 1);

            OpenXmlWorkbookSnapshot first =
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("first", ordered, limits);
            OpenXmlWorkbookSnapshot second =
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("second", swapped, limits);

            foreach (string name in new[] { "Alpha", "Beta" })
            {
                OpenXmlWorksheetSnapshot left = Assert.Single(first.Sheets, sheet => sheet.Name == name);
                OpenXmlWorksheetSnapshot right = Assert.Single(second.Sheets, sheet => sheet.Name == name);
                Assert.Empty(ExactComparisonEngine.Compare(left.Snapshot, right.Snapshot));
                Assert.Contains(
                    left.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.FormulaMetadataUnresolved &&
                            item.Message.Contains("as a whole", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(
                    right.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.FormulaMetadataUnresolved &&
                            item.Message.Contains("as a whole", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Workbook_partitioned_shared_formula_budget_is_sheet_order_independent()
        {
            byte[] ordered = TwoSheetSharedWorkbook(false);
            byte[] swapped = TwoSheetSharedWorkbook(true);
            var limits = new OpenXmlFormulaMetadataLimits(
                1000,
                1024 * 1024,
                1000,
                1024 * 1024,
                maxSharedFormulaDerivedCharacters: 1);

            OpenXmlWorkbookSnapshot first =
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("first", ordered, limits);
            OpenXmlWorkbookSnapshot second =
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("second", swapped, limits);

            foreach (string name in new[] { "Alpha", "Beta" })
            {
                OpenXmlWorksheetSnapshot left = Assert.Single(first.Sheets, sheet => sheet.Name == name);
                OpenXmlWorksheetSnapshot right = Assert.Single(second.Sheets, sheet => sheet.Name == name);
                Assert.Empty(ExactComparisonEngine.Compare(left.Snapshot, right.Snapshot));
                Assert.All(left.FormulaObservations, item => Assert.Null(item.EffectiveFormulaText));
                Assert.All(right.FormulaObservations, item => Assert.Null(item.EffectiveFormulaText));
                Assert.Contains(
                    left.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                            item.Message.Contains("workbook-partitioned", StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Workbook_budget_counts_duplicate_sheet_declarations_that_share_one_physical_part()
        {
            byte[] source = OpenXmlWorkbookFixture.Create(new[]
            {
                NamedSheet(
                    "Data",
                    1,
                    SharedMaster("A1", 1U, "A1:A2", "B1+1", "1"),
                    SharedFollower("A2", 1U, "2")),
            });
            byte[] duplicated = AddDuplicateSheetDeclaration(source, "Clone", 2U);
            var limits = new OpenXmlFormulaMetadataLimits(
                1000,
                1024 * 1024,
                1000,
                1024 * 1024,
                maxSharedFormulaDerivedCharacters: 3000);

            OpenXmlWorkbookSnapshot workbook =
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("duplicates", duplicated, limits);

            Assert.Equal(2, workbook.Sheets.Count);
            Assert.All(workbook.Sheets, sheet =>
            {
                Assert.All(sheet.FormulaObservations, item => Assert.Null(item.EffectiveFormulaText));
                Assert.Contains(
                    sheet.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.SharedFormulaNotExpanded &&
                            item.Message.Contains("workbook-partitioned", StringComparison.OrdinalIgnoreCase));
            });
        }

        [Fact]
        public void Table_definition_formulas_are_exact_observations_but_are_not_projected_to_cells()
        {
            byte[] bytes = AddTable(
                OpenXmlWorkbookFixture.Create(new[]
                {
                    Sheet(
                        InlineString("A1", "Qty"),
                        InlineString("B1", "Calc"),
                        Number("A2", "2"),
                        Number("B2", "4")),
                }));

            OpenXmlWorksheetSnapshot sheet = Read(bytes);

            Assert.Collection(
                sheet.TableFormulaObservations,
                calculated =>
                {
                    Assert.Equal(OpenXmlTableFormulaKind.CalculatedColumn, calculated.Kind);
                    Assert.Equal("T_Data", calculated.TableName);
                    Assert.Equal("A1:B3", calculated.TableRange);
                    Assert.Equal("Calc", calculated.ColumnName);
                    Assert.Equal("[@Qty]*2", calculated.ExactFormulaText);
                    Assert.NotNull(calculated.ExactMetadata);
                },
                totals =>
                {
                    Assert.Equal(OpenXmlTableFormulaKind.TotalsRow, totals.Kind);
                    Assert.Equal("SUBTOTAL(109,[Calc])", totals.ExactFormulaText);
                });
            Assert.Contains(
                sheet.CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.TableFormulaNotProjected &&
                        item.AffectsExactCellComparison);
            Assert.False(sheet.CanAssertExactCellEquality);
            Assert.Same(
                sheet.TableFormulaObservations[0].ExactTableMetadata,
                sheet.TableFormulaObservations[1].ExactTableMetadata);
        }

        [Fact]
        public void Workbook_partitioned_table_formula_budget_discards_each_sheet_as_a_whole()
        {
            byte[] source = OpenXmlWorkbookFixture.Create(new[]
            {
                NamedSheet("Alpha", 1, Number("A1", "1")),
                NamedSheet("Beta", 2, Number("A1", "1")),
            });
            byte[] withTables = AddTableToEveryWorksheet(source, new string('x', 512));
            var limits = new OpenXmlFormulaMetadataLimits(
                1000,
                1024 * 1024,
                1000,
                1024 * 1024,
                maxSharedFormulaDerivedCharacters: 1024 * 1024,
                maxTableFormulaRecords: 100,
                maxTableFormulaEvidenceCharacters: 128);

            OpenXmlWorkbookSnapshot workbook =
                OpenXmlWorkbookSnapshotAdapter.FromBytesForTesting("tables", withTables, limits);

            Assert.Equal(2, workbook.Sheets.Count);
            Assert.All(workbook.Sheets, sheet =>
            {
                Assert.Empty(sheet.TableFormulaObservations);
                Assert.Contains(
                    sheet.CoverageMessages,
                    item => item.Code == OpenXmlCoverageCodes.TableFormulaPayloadMalformed &&
                            item.Message.Contains("discarded as a whole", StringComparison.OrdinalIgnoreCase));
            });
        }

        [Fact]
        public void Cached_result_changes_remain_secondary_and_are_never_recalculated()
        {
            OpenXmlWorksheetSnapshot before = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(Formula("A1", "1+1", "2", CellValues.Number)),
            }));
            OpenXmlWorksheetSnapshot after = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(Formula("A1", "1+1", "3", CellValues.Number)),
            }));

            AtomicCellChange change = Assert.Single(
                ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));

            Assert.Equal(CellChangeAspect.CachedResult, change.Aspect);
            Assert.True(change.IsSecondary);
            Assert.Equal(CanonicalValue.Number(2), before.FormulaObservations[0].CachedResult);
            Assert.Equal(CanonicalValue.Number(3), after.FormulaObservations[0].CachedResult);
        }

        [Fact]
        public void Constant_vm_changes_are_indeterminate_without_a_false_primary_value_edit()
        {
            Cell beforeCell = Number("A1", "7");
            beforeCell.SetAttribute(new OpenXmlAttribute("vm", "", "1"));
            Cell afterCell = Number("A1", "7");
            afterCell.SetAttribute(new OpenXmlAttribute("vm", "", "2"));
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[] { Sheet(beforeCell) }));
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[] { Sheet(afterCell) }));

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(before, after);

            Assert.Equal(ExactComparisonStatus.Indeterminate, result.Status);
            Assert.Equal(0, result.AtomicChangeCount);
            Assert.False(result.CanAssertExactCellEquality);
            Assert.Equal(CanonicalValue.Number(7), before.Sheets[0].Snapshot.Cells["A1"].ConstantValue);
            Assert.Equal(CanonicalValue.Number(7), after.Sheets[0].Snapshot.Cells["A1"].ConstantValue);
            Assert.All(new[] { before, after }, workbook => Assert.Contains(
                workbook.Sheets[0].CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.ValueMetadataNotInterpreted &&
                        item.AffectsExactCellComparison));
        }

        [Fact]
        public void Constant_unresolved_cm_and_cell_extensions_are_indeterminate_without_a_value_edit()
        {
            Cell beforeCell = NumberWithCellMetadata("A1", "7", "malformed", "{TEST-ONE}");
            Cell afterCell = NumberWithCellMetadata("A1", "7", "999", "{TEST-TWO}");
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[] { Sheet(beforeCell) }));
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[] { Sheet(afterCell) }));

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(before, after);

            Assert.Equal(ExactComparisonStatus.Indeterminate, result.Status);
            Assert.Equal(0, result.AtomicChangeCount);
            Assert.All(new[] { before, after }, workbook => Assert.Contains(
                workbook.Sheets[0].CoverageMessages,
                item => item.Code == OpenXmlCoverageCodes.ValueMetadataNotInterpreted &&
                        item.Message.Contains("extLst=present", StringComparison.Ordinal)));
        }

        [Fact]
        public void Numeric_fallback_fingerprint_normalizes_absent_t_and_explicit_n_aliases()
        {
            var implicitNumber = new Cell(new CellValue("not-a-number")) { CellReference = "A1" };
            var explicitNumber = new Cell(new CellValue("not-a-number"))
            {
                CellReference = "A1",
                DataType = CellValues.Number,
            };
            OpenXmlWorkbookSnapshot before = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "before",
                OpenXmlWorkbookFixture.Create(new[] { Sheet(implicitNumber) }));
            OpenXmlWorkbookSnapshot after = OpenXmlWorkbookSnapshotAdapter.FromBytes(
                "after",
                OpenXmlWorkbookFixture.Create(new[] { Sheet(explicitNumber) }));

            ExactWorkbookComparisonResult result = ExactWorkbookComparisonEngine.Compare(before, after);

            Assert.Equal(ExactComparisonStatus.Indeterminate, result.Status);
            Assert.Equal(0, result.AtomicChangeCount);
            Assert.Equal(
                before.Sheets[0].Snapshot.Cells["A1"].ConstantValue,
                after.Sheets[0].Snapshot.Cells["A1"].ConstantValue);
        }

        [Fact]
        public void Invalid_formula_cache_diagnostics_are_bounded_and_remain_cell_equality_affecting()
        {
            var cells = new List<Cell>();
            for (int row = 1; row <= 20; row++)
                cells.Add(Formula("A" + row, "1+1", "not-a-number", CellValues.Number));

            OpenXmlWorksheetSnapshot sheet = Read(
                OpenXmlWorkbookFixture.Create(new[] { Sheet(cells.ToArray()) }));
            List<OpenXmlCoverageMessage> findings = sheet.CoverageMessages
                .Where(item => item.Code == OpenXmlCoverageCodes.InvalidNumericValue)
                .ToList();

            Assert.Equal(9, findings.Count);
            Assert.Equal(8, findings.Count(item => item.CellReference != null));
            Assert.Contains(
                findings,
                item => item.CellReference == null &&
                        item.Message.Contains("12 additional", StringComparison.Ordinal));
            Assert.All(findings, item => Assert.True(item.AffectsExactCellComparison));
            Assert.False(sheet.CanAssertExactCellEquality);
        }

        [Fact]
        public void Unresolved_formula_placeholder_is_stable_across_cache_only_changes()
        {
            OpenXmlWorksheetSnapshot before = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(SharedFollower("A1", 99U, "1")),
            }));
            OpenXmlWorksheetSnapshot after = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(SharedFollower("A1", 99U, "2")),
            }));

            IReadOnlyList<AtomicCellChange> changes =
                ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot);

            AtomicCellChange change = Assert.Single(changes);
            Assert.Equal(CellChangeAspect.CachedResult, change.Aspect);
            Assert.DoesNotContain(changes, item => item.Aspect == CellChangeAspect.Formula);
            Assert.False(before.CanAssertExactCellEquality);
            Assert.False(after.CanAssertExactCellEquality);
            Assert.Equal(
                before.Snapshot.Cells["A1"].FormulaText,
                after.Snapshot.Cells["A1"].FormulaText);
        }

        [Fact]
        public void Formula_character_diff_reconstructs_exact_OOXML_text_without_normalization()
        {
            const string beforeText = "SUM(A1:A9) + \"Cafe\u0301 A1\"";
            const string afterText = "SUM(A1:A10)+\"Café A1\"";
            OpenXmlWorksheetSnapshot before = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(Formula("A1", beforeText, "1", CellValues.Number)),
            }));
            OpenXmlWorksheetSnapshot after = Read(OpenXmlWorkbookFixture.Create(new[]
            {
                Sheet(Formula("A1", afterText, "1", CellValues.Number)),
            }));

            AtomicCellChange formulaChange = Assert.Single(
                ExactComparisonEngine.Compare(before.Snapshot, after.Snapshot));
            CharacterDiffResult diff = CharacterDiff.Compare(
                formulaChange.BeforeText,
                formulaChange.AfterText);

            Assert.Equal(beforeText, diff.BeforeText);
            Assert.Equal(afterText, diff.AfterText);
            AssertReconstructs(diff);
        }

        private static OpenXmlWorkbookFixture.SheetDefinition Sheet(params Cell[] cells) =>
            NamedSheet("Data", 1, cells);

        private static OpenXmlWorkbookFixture.SheetDefinition NamedSheet(
            string name,
            uint sheetId,
            params Cell[] cells) =>
            new OpenXmlWorkbookFixture.SheetDefinition
            {
                Name = name,
                SheetId = sheetId,
                CodeName = name + "Code",
                Cells = cells,
            };

        private static OpenXmlWorksheetSnapshot Read(byte[] bytes) =>
            Assert.Single(OpenXmlWorkbookSnapshotAdapter.FromBytes("formula-spike", bytes).Sheets);

        private static OpenXmlWorksheetSnapshot ReadNamed(byte[] bytes, string name) =>
            Assert.Single(
                OpenXmlWorkbookSnapshotAdapter.FromBytes("formula-spike", bytes).Sheets,
                sheet => sheet.Name == name);

        private static Cell SharedMaster(
            string reference,
            uint index,
            string range,
            string formula,
            string cached) =>
            new Cell(
                new CellFormula(formula)
                {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = index,
                    Reference = range,
                },
                new CellValue(cached))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell SharedFollower(string reference, uint index, string cached) =>
            new Cell(
                new CellFormula
                {
                    FormulaType = CellFormulaValues.Shared,
                    SharedIndex = index,
                },
                new CellValue(cached))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell Formula(
            string reference,
            string formula,
            string cached,
            CellValues cachedType) =>
            new Cell(new CellFormula(formula), new CellValue(cached))
            {
                CellReference = reference,
                DataType = cachedType,
            };

        private static Cell Number(string reference, string value) =>
            new Cell(new CellValue(value))
            {
                CellReference = reference,
                DataType = CellValues.Number,
            };

        private static Cell NumberWithCellMetadata(
            string reference,
            string value,
            string rawCellMetadataIndex,
            string extensionUri)
        {
            Cell cell = Number(reference, value);
            cell.SetAttribute(new OpenXmlAttribute("cm", "", rawCellMetadataIndex));
            cell.Append(new SpreadsheetExtensionList(
                new SpreadsheetExtension { Uri = extensionUri }));
            return cell;
        }

        private static Cell InlineString(string reference, string value) =>
            new Cell(new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(value)))
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
            };

        private static Cell UnsupportedValue(string reference, string value, uint styleIndex)
        {
            var cell = new Cell(new CellValue(value))
            {
                CellReference = reference,
                StyleIndex = styleIndex,
            };
            cell.SetAttribute(new OpenXmlAttribute("t", "", "future-type"));
            return cell;
        }

        private static byte[] AddDynamicArrayMetadata(byte[] source, int blockCount = 1)
        {
            if (blockCount < 1) throw new ArgumentOutOfRangeException(nameof(blockCount));
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart!;
                CellMetadataPart metadataPart = workbookPart.AddNewPart<CellMetadataPart>();
                var future = new FutureMetadata
                {
                    Name = "XLDAPR",
                    Count = (uint)blockCount,
                };
                var cellMetadata = new CellMetadata { Count = (uint)blockCount };
                for (int index = 0; index < blockCount; index++)
                {
                    var dynamicProperties = new DynamicArrayProperties
                    {
                        FDynamic = true,
                        FCollapsed = false,
                    };
                    var extension = new SpreadsheetExtension(dynamicProperties)
                    {
                        Uri = "{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}",
                    };
                    future.Append(new FutureMetadataBlock(new SpreadsheetExtensionList(extension)));
                    cellMetadata.Append(new MetadataBlock(
                        new MetadataRecord { TypeIndex = 1U, Val = (uint)index }));
                }
                metadataPart.Metadata = new Metadata(
                    new MetadataTypes(
                        new MetadataType
                        {
                            Name = "XLDAPR",
                            MinSupportedVersion = 120000U,
                            Copy = true,
                            PasteAll = true,
                            PasteFormulas = true,
                            PasteValues = true,
                            Merge = true,
                            SplitFirst = true,
                            RowColumnShift = true,
                            ClearAll = true,
                            ClearContents = true,
                            Assign = true,
                            Coerce = true,
                            CellMeta = true,
                        })
                    { Count = 1U },
                    future,
                    cellMetadata);
                metadataPart.Metadata.Save();
            }
            return stream.ToArray();
        }

        private static byte[] ReversePhysicalRows(byte[] source)
        {
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                Worksheet worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;
                List<Row> rows = sheetData.Elements<Row>().ToList();
                sheetData.RemoveAllChildren<Row>();
                foreach (Row row in rows.AsEnumerable().Reverse()) sheetData.Append(row);
                worksheet.Save();
            }
            return stream.ToArray();
        }

        private static byte[] TwoSheetDynamicMetadataWorkbook(bool swapped)
        {
            Cell alphaCell = Formula("A1", "1+1", "2", CellValues.Number);
            alphaCell.CellMetaIndex = 1U;
            Cell betaCell = Formula("A1", "1+1", "2", CellValues.Number);
            betaCell.CellMetaIndex = 1U;
            OpenXmlWorkbookFixture.SheetDefinition alpha = NamedSheet("Alpha", 1, alphaCell);
            OpenXmlWorkbookFixture.SheetDefinition beta = NamedSheet("Beta", 2, betaCell);
            return AddDynamicArrayMetadata(OpenXmlWorkbookFixture.Create(
                swapped ? new[] { beta, alpha } : new[] { alpha, beta }));
        }

        private static byte[] TwoSheetSharedWorkbook(bool swapped)
        {
            OpenXmlWorkbookFixture.SheetDefinition alpha = NamedSheet(
                "Alpha",
                1,
                SharedMaster("A1", 1U, "A1:A2", "B1+1", "1"),
                SharedFollower("A2", 1U, "2"));
            OpenXmlWorkbookFixture.SheetDefinition beta = NamedSheet(
                "Beta",
                2,
                SharedMaster("A1", 1U, "A1:A2", "B1+1", "1"),
                SharedFollower("A2", 1U, "2"));
            return OpenXmlWorkbookFixture.Create(swapped ? new[] { beta, alpha } : new[] { alpha, beta });
        }

        private static byte[] AddDuplicateSheetDeclaration(byte[] source, string name, uint sheetId)
        {
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                Workbook workbook = document.WorkbookPart!.Workbook!;
                Sheets sheets = workbook.Sheets!;
                Sheet original = sheets.Elements<Sheet>().Single();
                sheets.Append(new Sheet
                {
                    Name = name,
                    SheetId = sheetId,
                    Id = original.Id!.Value,
                });
                workbook.Save();
            }
            return stream.ToArray();
        }

        private static byte[] AddTable(byte[] source)
        {
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                WorksheetPart worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
                AttachTable(worksheetPart, 1U, "T_Data", null);
            }
            return stream.ToArray();
        }

        private static byte[] AddTableToEveryWorksheet(byte[] source, string largeAttribute)
        {
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true))
            {
                uint index = 0;
                foreach (WorksheetPart worksheetPart in document.WorkbookPart!.WorksheetParts)
                {
                    index++;
                    AttachTable(worksheetPart, index, "T_" + index, largeAttribute);
                }
            }
            return stream.ToArray();
        }

        private static void AttachTable(
            WorksheetPart worksheetPart,
            uint id,
            string name,
            string? largeAttribute)
        {
            TableDefinitionPart tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
            tablePart.Table = new Table(
                new AutoFilter { Reference = "A1:B3" },
                new TableColumns(
                    new TableColumn(new CalculatedColumnFormula("[@Qty]*2"))
                    {
                        Id = 1U,
                        Name = "Calc",
                    },
                    new TableColumn(new TotalsRowFormula("SUBTOTAL(109,[Calc])"))
                    {
                        Id = 2U,
                        Name = "Total",
                    })
                { Count = 2U })
            {
                Id = id,
                Name = name,
                DisplayName = name,
                Reference = "A1:B3",
                HeaderRowCount = 1U,
                TotalsRowCount = 1U,
            };
            if (largeAttribute != null)
                tablePart.Table!.SetAttribute(new OpenXmlAttribute("large", "urn:test", largeAttribute));
            tablePart.Table!.Save();
            worksheetPart.Worksheet!.Append(
                new TableParts(
                    new TablePart { Id = worksheetPart.GetIdOfPart(tablePart) })
                { Count = 1U });
            worksheetPart.Worksheet.Save();
        }

        private static void AssertReconstructs(CharacterDiffResult result)
        {
            var before = new StringBuilder();
            var after = new StringBuilder();
            foreach (CharacterDiffRun run in result.Runs)
            {
                if (run.Kind != CharacterDiffKind.Added) before.Append(run.Text);
                if (run.Kind != CharacterDiffKind.Removed) after.Append(run.Text);
            }
            Assert.Equal(result.BeforeText, before.ToString());
            Assert.Equal(result.AfterText, after.ToString());
        }
    }
}
