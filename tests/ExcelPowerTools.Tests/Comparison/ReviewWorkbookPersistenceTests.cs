using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewWorkbookPersistenceTests : IDisposable
    {
        private readonly string _directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelPowerTools.ReviewWorkbook.Tests",
            Guid.NewGuid().ToString("N"));

        public ReviewWorkbookPersistenceTests()
        {
            Directory.CreateDirectory(_directory);
        }

        [Fact]
        public void CreateNew_writes_exactly_eight_visible_Office2007_valid_sheets()
        {
            string path = CreateWorkbook();

            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
            WorkbookPart workbookPart = Assert.IsType<WorkbookPart>(document.WorkbookPart);
            List<Sheet> sheets = workbookPart.Workbook!.Descendants<Sheet>().ToList();

            Assert.Equal(ReviewWorkbookFile.SheetNames, sheets.Select(sheet => sheet.Name!.Value).ToList());
            Assert.Equal(8, sheets.Count);
            Assert.All(sheets, sheet => Assert.Equal(SheetStateValues.Visible, sheet.State!.Value));

            var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
            Assert.Empty(validator.Validate(document));
        }

        [Fact]
        public void Details_contains_rich_character_diff_runs_and_Overview_has_only_internal_link()
        {
            var state = new ReviewWorkbookState("project-rich-diff", "{\"state\":1}")
            {
                DetailBefore = "Gross margin 41.0%",
                DetailAfter = "Gross margin 42.5%",
            };
            string path = CreateWorkbook(state);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
            WorkbookPart workbookPart = document.WorkbookPart!;
            WorksheetPart details = GetSheetPart(workbookPart, "Details");
            Cell before = details.Worksheet!.Descendants<Cell>()
                .Single(cell => cell.CellReference!.Value == "B3");
            Cell after = details.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference!.Value == "B4");

            List<Run> beforeRuns = before.InlineString!.Elements<Run>().ToList();
            List<Run> afterRuns = after.InlineString!.Elements<Run>().ToList();
            Assert.True(beforeRuns.Count >= 3);
            Assert.True(afterRuns.Count >= 3);
            Assert.Contains(beforeRuns, run => run.RunProperties?.GetFirstChild<Strike>() != null &&
                                               run.RunProperties.GetFirstChild<Color>()?.Rgb?.Value == "FFC62828");
            Assert.Contains(afterRuns, run => run.RunProperties?.GetFirstChild<Bold>() != null &&
                                              run.RunProperties.GetFirstChild<Color>()?.Rgb?.Value == "FF2E7D32");
            Assert.Equal("Gross margin 41.0%", string.Concat(beforeRuns.Select(run => run.Text!.Text)));
            Assert.Equal("Gross margin 42.5%", string.Concat(afterRuns.Select(run => run.Text!.Text)));

            WorksheetPart overview = GetSheetPart(workbookPart, "Overview");
            List<Hyperlink> navigationLinks = overview.Worksheet!.Descendants<Hyperlink>().ToList();
            Assert.Equal(4, navigationLinks.Count);
            Hyperlink link = navigationLinks.Single(candidate => candidate.Reference!.Value == "A7");
            Assert.Equal("A7", link.Reference!.Value);
            Assert.Equal("'Changes'!A1", link.Location!.Value);
            Assert.Null(link.Id);
            Assert.Empty(overview.HyperlinkRelationships);
        }

        [Fact]
        public void Details_makes_changed_whitespace_visible_without_mutating_embedded_state()
        {
            const string stateJson = "{\"state\":\"exact \\t\\r\\n\"}";
            var state = new ReviewWorkbookState("project-visible-whitespace", stateJson)
            {
                DetailBefore = "A B\tC\r\nD",
                DetailAfter = "A  B C\nD",
            };
            string path = CreateWorkbook(state);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorksheetPart details = GetSheetPart(document.WorkbookPart!, "Details");
                Cell before = details.Worksheet!.Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "B3");
                Cell after = details.Worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "B4");

                string renderedBefore = string.Concat(
                    before.InlineString!.Elements<Run>().Select(run => run.Text!.Text));
                string renderedAfter = string.Concat(
                    after.InlineString!.Elements<Run>().Select(run => run.Text!.Text));
                Assert.Contains("·", renderedBefore + renderedAfter, StringComparison.Ordinal);
                Assert.Contains("⇥", renderedBefore, StringComparison.Ordinal);
                Assert.Contains("␍", renderedBefore, StringComparison.Ordinal);
                Assert.Contains("␊", renderedBefore + renderedAfter, StringComparison.Ordinal);
                Assert.DoesNotContain('\t', renderedBefore + renderedAfter);
                Assert.DoesNotContain('\r', renderedBefore + renderedAfter);
                Assert.DoesNotContain('\n', renderedBefore + renderedAfter);
            }

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(stateJson, read.StateJson);
        }

        [Fact]
        public void Changes_sheet_streams_visible_rows_that_are_bound_to_the_embedded_state()
        {
            var state = new ReviewWorkbookState("project-change-rows", "{\"changes\":2}");
            state.ChangeRows.Add(new ReviewChangeProjectionRow(
                "change-1", "Input", "B2", "Value", "before", "after"));
            state.ChangeRows.Add(new ReviewChangeProjectionRow(
                "change-2", "Input", "C7", "Formula", "=A1", "=A1+1"));
            string path = CreateWorkbook(state);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                List<Row> rows = changes.Worksheet!.Descendants<Row>().ToList();
                Assert.Equal(3, rows.Count);
                Assert.Equal(
                    "Change IDSheetAddressAspectBeforeAfterGroup IDFlagsRevisionsDistinct statesAlignment confidence",
                    rows[0].InnerText);
                Assert.Equal("change-1InputB2Valuebeforeafter22", rows[1].InnerText);
                Assert.Equal("change-2InputC7Formula=A1=A1+122", rows[2].InnerText);
            }

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(state.StateJson, read.StateJson);
        }

        [Fact]
        public void Generated_sheets_write_content_sized_visible_columns()
        {
            const string sheetName = "A meaningfully wider sheet label";
            var state = new ReviewWorkbookState("project-column-widths", "{\"widths\":true}");
            state.ChangeRows.Add(new ReviewChangeProjectionRow(
                "change-width", sheetName, "B27", "Value", "short", "a wider replacement value"));
            state.DiagnosticRows.Add(new ReviewDiagnosticProjectionRow(
                "finding-width",
                "CMP-WIDTH",
                "Warning",
                "revision-1",
                "Width check",
                new string('x', 1000),
                "Inspect the generated review."));
            string path = CreateWorkbook(state);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart!;
                foreach (Sheet sheet in workbookPart.Workbook!.Descendants<Sheet>())
                {
                    WorksheetPart part = Assert.IsType<WorksheetPart>(workbookPart.GetPartById(sheet.Id!.Value!));
                    Columns columns = Assert.Single(part.Worksheet!.Elements<Columns>());
                    Assert.NotEmpty(columns.Elements<Column>());
                    Assert.All(columns.Elements<Column>(), column =>
                    {
                        Assert.True(column.CustomWidth?.Value);
                        Assert.True(column.BestFit?.Value);
                        Assert.InRange(column.Width!.Value, 8.43D, 255D);
                    });
                }

                WorksheetPart changes = GetSheetPart(workbookPart, "Changes");
                Column sheetColumn = changes.Worksheet!.Descendants<Column>()
                    .Single(column => column.Min!.Value == 2U && column.Max!.Value == 2U);
                Assert.True(sheetColumn.Width!.Value >= sheetName.Length + 2D);

                WorksheetPart diagnostics = GetSheetPart(workbookPart, "Diagnostics");
                Column detailColumn = diagnostics.Worksheet!.Descendants<Column>()
                    .Single(column => column.Min!.Value == 6U && column.Max!.Value == 6U);
                Assert.Equal(255D, detailColumn.Width!.Value);
            }

            Assert.Equal(state.StateJson, ReviewWorkbookFile.ReadAndValidate(path).StateJson);
        }

        [Fact]
        public void Revision_and_diagnostic_projections_are_visible_and_integrity_bound()
        {
            var state = new ReviewWorkbookState("project-visible-context", "{\"schema\":\"history-v1\"}");
            state.RevisionRows.Add(new ReviewRevisionProjectionRow(
                "r2",
                "v2",
                1,
                "r1",
                "Master",
                "workbook-v2.xlsx",
                @"C:\mail\review\workbook-v2.xlsx",
                new string('a', 64),
                "2026-07-17T10:00:00Z",
                "explicit v2 token",
                "Strong"));
            state.DiagnosticRows.Add(new ReviewDiagnosticProjectionRow(
                "finding-1",
                "CMP-SHEET-MAPPING-AMBIGUOUS",
                "Correctable",
                "r2",
                "Two sheets have equal evidence.",
                "Candidates Input A and Input B both scored 0.72.",
                "Confirm the sheet mapping."));
            string path = CreateWorkbook(state);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                List<Row> revisionRows = GetSheetPart(document.WorkbookPart!, "Revisions")
                    .Worksheet!.Descendants<Row>().ToList();
                List<Row> diagnosticRows = GetSheetPart(document.WorkbookPart!, "Diagnostics")
                    .Worksheet!.Descendants<Row>().ToList();
                Assert.Equal(2, revisionRows.Count);
                Assert.Contains("workbook-v2.xlsx", revisionRows[1].InnerText, StringComparison.Ordinal);
                Assert.Contains(@"C:\mail\review\workbook-v2.xlsx", revisionRows[1].InnerText, StringComparison.Ordinal);
                Assert.Equal(2, diagnosticRows.Count);
                Assert.Contains("CMP-SHEET-MAPPING-AMBIGUOUS", diagnosticRows[1].InnerText, StringComparison.Ordinal);
            }

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(state.StateJson, read.StateJson);
        }

        [Fact]
        public void Custom_state_roundtrips_across_multiple_chunks_discovered_without_part_order_assumptions()
        {
            string json = "{\"project\":\"αβγ\",\"notes\":\"" + new string('x', 137) + "\"}";
            var state = new ReviewWorkbookState("project-roundtrip", json);
            string path = CreateWorkbook(state, chunkSize: 17);

            List<string> roots;
            List<string> persistentIds;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                List<CustomXmlPart> parts = document.WorkbookPart!.CustomXmlParts.ToList();
                roots = parts.Select(ReadRootLocalName).ToList();
                persistentIds = parts
                    .Select(part => part.CustomXmlPropertiesPart!.DataStoreItem!.ItemId!.Value!)
                    .ToList();
            }

            Assert.Equal("stateChunk", roots.First());
            Assert.Equal("reviewManifest", roots.Last());
            Assert.True(roots.Count(name => name == "stateChunk") > 1);
            Assert.Equal(persistentIds.Count, persistentIds.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.All(persistentIds, itemId => Assert.True(Guid.TryParse(itemId, out _)));

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(state.ProjectId, read.ProjectId);
            Assert.Equal(json, read.StateJson);
            Assert.True(read.ChunkCount > 1);
            Assert.Equal(read.ChunkCount, read.ChunkItemIds.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.DoesNotContain(read.ManifestItemId, read.ChunkItemIds, StringComparer.OrdinalIgnoreCase);
            Assert.Equal(64, read.StateSha256.Length);
            Assert.Equal(64, read.ProjectionSha256.Length);
        }

        [Fact]
        public void Empty_embedded_state_roundtrips_as_one_bounded_zero_length_chunk()
        {
            string path = CreateWorkbook(new ReviewWorkbookState("project-empty-state", ""));

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);

            Assert.Equal("", read.StateJson);
            Assert.Equal(1, read.ChunkCount);
            Assert.False(read.HasDecisionState);
        }

        [Fact]
        public void Empty_decision_generation_roundtrips_with_full_evidence_catalog_and_manifest_binding()
        {
            ReviewDecisionGeneration expected = CreateDecisionRoot();
            ReviewWorkbookState state = StateWithDecision(
                expected,
                "{\"state\":\"decision-root\"}");
            string path = CreateWorkbook(state);

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);

            Assert.True(read.HasDecisionState);
            Assert.NotNull(read.DecisionGeneration);
            Assert.Equal(expected.GenerationId, read.DecisionGeneration!.GenerationId);
            Assert.Equal(expected.CatalogId, read.DecisionGeneration.CatalogId);
            Assert.Equal(expected.BaseGenerationId, read.DecisionGeneration.BaseGenerationId);
            Assert.Equal(expected.MappingGenerationId, read.DecisionGeneration.MappingGenerationId);
            Assert.Empty(read.DecisionGeneration.Events);
            Assert.Equal(2, read.DecisionGeneration.Catalog.Subjects.Count);
            Assert.Equal(4, read.DecisionGeneration.Catalog.Observations.Count);
            Assert.Equal(2, read.DecisionGeneration.Catalog.Groups.Count);
            Assert.Equal(64, read.DecisionSha256.Length);
            Assert.True(Guid.TryParse(read.DecisionItemId, out _));

            Observation formula = read.DecisionGeneration.Catalog.Observations
                .Single(item => item.ObservationId == DecisionObservationId(
                    "subject-formula",
                    "candidate",
                    ObservationExactType.Formula,
                    new string('1', 64)));
            Assert.Equal("=A1+1", formula.Formula);
            Assert.Equal("42", formula.CachedResult!.CanonicalText);
            Assert.Equal("logical-sheet", formula.Provenance.LogicalSheetId);
            Assert.Equal("candidate.xlsx", formula.Provenance.SourceLabel);
            Assert.Equal(7, formula.Provenance.PhysicalRow);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
            CustomXmlPart manifest = document.WorkbookPart!.CustomXmlParts
                .Single(part => ReadRootLocalName(part) == "reviewManifest");
            using Stream stream = manifest.GetStream(FileMode.Open, FileAccess.Read);
            XElement root = XDocument.Load(stream).Root!;
            Assert.Equal(read.DecisionItemId, root.Attribute("decisionItemId")!.Value);
            Assert.Equal(read.DecisionSha256, root.Attribute("decisionSha256")!.Value);
            Assert.Equal(
                expected.Catalog.EvidenceGenerationId,
                root.Attribute("decisionEvidenceGenerationId")!.Value);
            Assert.Equal(
                expected.MappingGenerationId,
                root.Attribute("decisionMappingGenerationId")!.Value);
            Assert.Equal(
                ReviewWorkbookFile.DecisionBoundStreamingProjectionManifestSchemaVersion,
                root.Attribute("schemaVersion")!.Value);
            Assert.Equal(
                ReviewWorkbookFile.StreamingProjectionDigestVersion,
                root.Attribute("projectionDigestVersion")!.Value);

            CustomXmlPart decision = GetDecisionPart(document.WorkbookPart!);
            using Stream decisionStream = decision.GetStream(FileMode.Open, FileAccess.Read);
            XElement decisionRoot = XDocument.Load(decisionStream).Root!;
            Assert.Equal(
                ReviewWorkbookFile.SchemaVersion,
                decisionRoot.Attribute("schemaVersion")!.Value);
            Assert.Equal("0", decisionRoot.Attribute("eventCount")!.Value);
            XElement rootGeneration = decisionRoot.Elements()
                .Single(element => element.Name.LocalName == "generation");
            Assert.Equal("0", rootGeneration.Attribute("eventCount")!.Value);
            Assert.Empty(rootGeneration.Elements());
        }

        [Fact]
        public void Nonroot_decision_generation_roundtrips_complete_ordered_event_audit_by_replay()
        {
            ReviewDecisionGeneration expected = CreateAdvancedDecisionGeneration();
            string path = CreateWorkbook(StateWithDecision(
                expected,
                "{\"state\":\"advanced-decision\"}"));

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);

            Assert.True(read.HasDecisionState);
            Assert.NotNull(read.DecisionGeneration);
            AssertDecisionGenerationEqual(expected, read.DecisionGeneration!);
            Assert.Equal(3, read.DecisionGeneration!.Events.Count);
            Assert.NotEmpty(read.DecisionGeneration.Events[0].RequestId);
            Assert.Equal("Confirmed\u0001 reviewer", read.DecisionGeneration.Events[0]
                .ReviewerSnapshot.ConfirmedDisplayName);
            Assert.Equal("build\u0003candidate", read.DecisionGeneration.Events[1].AddInBuild);
            Assert.Equal("selected\u0004candidate", read.DecisionGeneration.Events[1].Note);
            Assert.Equal(2, read.DecisionGeneration.Events[1].Locations.Count);
            Assert.Equal(
                read.DecisionGeneration.Events[1].EventId,
                read.DecisionGeneration.Events[2].SupersededEventId);

            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, false);
            using Stream stream = GetDecisionPart(document.WorkbookPart!).GetStream(
                FileMode.Open,
                FileAccess.Read);
            XElement decisionRoot = XDocument.Load(stream).Root!;
            Assert.Equal(
                ReviewWorkbookFile.DecisionEventStateSchemaVersion,
                decisionRoot.Attribute("schemaVersion")!.Value);
            Assert.Equal("3", decisionRoot.Attribute("eventCount")!.Value);
            XElement generation = decisionRoot.Elements()
                .Single(element => element.Name.LocalName == "generation");
            List<XElement> records = generation.Elements().ToList();
            Assert.Equal(3, records.Count);
            Assert.Equal(new[] { "0", "1", "2" }, records
                .Select(element => element.Attribute("sequence")!.Value));
            Assert.All(
                records,
                element => Assert.Equal(
                    ReviewDecisionEventContractVersions.V1.ToString(CultureInfo.InvariantCulture),
                    element.Attribute("contractVersion")!.Value));
        }

        [Fact]
        public void Reader_rejects_tampered_or_reordered_v2_decision_event_ledgers()
        {
            string payloadPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(payloadPath, root =>
            {
                XElement note = root.Descendants()
                    .First(element => element.Name.LocalName == "note");
                byte[] bytes = Encoding.UTF8.GetBytes("tampered event note");
                note.SetAttributeValue("byteLength", bytes.Length);
                note.Value = Convert.ToBase64String(bytes);
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(payloadPath));

            string recordOrderPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(recordOrderPath, root =>
            {
                XElement generation = root.Elements()
                    .Single(element => element.Name.LocalName == "generation");
                ReverseFirstTwoChildren(generation);
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(recordOrderPath));

            string grammarOrderPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(grammarOrderPath, root =>
            {
                XElement firstEvent = root.Descendants()
                    .First(element => element.Name.LocalName == "event");
                XElement note = firstEvent.Elements()
                    .Single(element => element.Name.LocalName == "note");
                XElement addInBuild = firstEvent.Elements()
                    .Single(element => element.Name.LocalName == "addInBuild");
                note.Remove();
                addInBuild.AddBeforeSelf(note);
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(grammarOrderPath));

            string requestPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(requestPath, root => root.Descendants()
                .First(element => element.Name.LocalName == "event")
                .SetAttributeValue("requestId", new string('f', 64)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(requestPath));
        }

        [Fact]
        public void Reader_enforces_v1_root_only_and_v2_event_collection_limits()
        {
            string v1Path = CreateDecisionWorkbook();
            MutateDecisionAndRebind(v1Path, root =>
            {
                root.SetAttributeValue("eventCount", "1");
                root.Elements().Single(element => element.Name.LocalName == "generation")
                    .SetAttributeValue("eventCount", "1");
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(v1Path));

            string eventLimitPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(eventLimitPath, root => root.SetAttributeValue(
                "eventCount",
                (ReviewWorkbookFile.MaxDecisionEventCount + 1).ToString(
                    CultureInfo.InvariantCulture)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(eventLimitPath));

            string locationLimitPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(locationLimitPath, root => root.Descendants()
                .First(element => element.Name.LocalName == "event")
                .SetAttributeValue(
                    "locationCount",
                    (ReviewWorkbookFile.MaxDecisionLocationCountPerEvent + 1).ToString(
                        CultureInfo.InvariantCulture)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(locationLimitPath));

            string contractPath = CreateAdvancedDecisionWorkbook();
            MutateDecisionAndRebind(contractPath, root => root.Descendants()
                .First(element => element.Name.LocalName == "replayRecord")
                .SetAttributeValue("contractVersion", "2"));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(contractPath));
        }

        [Fact]
        public void Reader_rejects_missing_or_orphaned_decision_state_and_unpaired_manifest_binding()
        {
            string missingPath = CreateDecisionWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(missingPath, true))
            {
                CustomXmlPart decision = GetDecisionPart(document.WorkbookPart!);
                document.WorkbookPart!.DeletePart(decision);
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(missingPath));

            string orphanPath = CreateDecisionWorkbook();
            MutateManifest(orphanPath, root =>
            {
                root.Attribute("decisionItemId")!.Remove();
                root.Attribute("decisionSha256")!.Remove();
                root.Attribute("decisionEvidenceGenerationId")!.Remove();
                root.Attribute("decisionMappingGenerationId")!.Remove();
                root.SetAttributeValue("schemaVersion", ReviewWorkbookFile.SchemaVersion);
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(orphanPath));

            string unpairedPath = CreateDecisionWorkbook();
            MutateManifest(unpairedPath, root => root.Attribute("decisionSha256")!.Remove());
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(unpairedPath));

            string strippedPath = CreateDecisionWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(strippedPath, true))
                document.WorkbookPart!.DeletePart(GetDecisionPart(document.WorkbookPart));
            MutateManifest(strippedPath, root =>
            {
                root.Attribute("decisionItemId")!.Remove();
                root.Attribute("decisionSha256")!.Remove();
                root.Attribute("decisionEvidenceGenerationId")!.Remove();
                root.Attribute("decisionMappingGenerationId")!.Remove();
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(strippedPath));
        }

        [Fact]
        public void Reader_rejects_duplicate_tampered_unknown_and_over_limit_decision_state()
        {
            string duplicatePath = CreateDecisionWorkbook();
            AddDuplicateDecisionPart(duplicatePath);
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(duplicatePath));

            string digestPath = CreateDecisionWorkbook();
            MutateManifest(digestPath, root => root.SetAttributeValue("decisionSha256", new string('0', 64)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(digestPath));

            string authorityPath = CreateDecisionWorkbook();
            MutateManifest(
                authorityPath,
                root => root.SetAttributeValue(
                    "decisionEvidenceGenerationId",
                    new string('5', 64)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(authorityPath));

            string unknownPath = CreateDecisionWorkbook();
            MutateDecisionAndRebind(
                unknownPath,
                root => root.SetAttributeValue("futureSilentMeaning", "not-allowed-in-v1"));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(unknownPath));

            string unknownRootPath = CreateDecisionWorkbook();
            MutateDecisionAndRebind(
                unknownRootPath,
                root => root.Name = root.Name.Namespace + "futureDecisionState");
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(unknownRootPath));

            string limitPath = CreateDecisionWorkbook();
            MutateDecisionAndRebind(
                limitPath,
                root => root.SetAttributeValue(
                    "subjectCount",
                    (ReviewWorkbookFile.MaxDecisionSubjectCount + 1).ToString(CultureInfo.InvariantCulture)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(limitPath));

            string bindingPath = CreateDecisionWorkbook();
            MutateDecisionAndRebind(
                bindingPath,
                root => root.SetAttributeValue("stateSha256", new string('f', 64)));
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(bindingPath));
        }

        [Fact]
        public void Reader_rejects_noncanonical_decision_collection_order()
        {
            AssertDecisionReorderRejected("evidence", "revision");
            AssertDecisionReorderRejected("subjects", "subject");
            AssertDecisionReorderRejected("observations", "observation");
            AssertDecisionReorderRejected("groups", "group");

            string memberPath = CreateDecisionWorkbook();
            MutateDecisionAndRebind(memberPath, root =>
            {
                XElement group = root.Descendants()
                    .Single(element => element.Name.LocalName == "group" &&
                                       element.Elements().Count() == 2);
                ReverseFirstTwoChildren(group);
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(memberPath));
        }

        [Fact]
        public void Streaming_preflight_rejects_high_node_decision_xml_before_catalog_materialization()
        {
            string path = CreateDecisionWorkbook();
            InjectDecisionElementsAndRebind(
                path,
                ReviewWorkbookFile.MaxDecisionXmlElementCount + 1);

            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(path));

            Assert.Contains("element-count", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Root_discovery_accepts_a_valid_decision_sidecar_larger_than_a_state_chunk()
        {
            string large = new string('x', 2300 * 1024);
            string contentHash = new string('9', 64);
            var evidence = new ReviewEvidenceGeneration(
                "project-large-decision-sidecar",
                "mapping-large-decision-sidecar",
                new[] { new ReviewEvidenceRevisionIdentity("master", 0, new string('8', 64)) });
            var subject = new ReviewDecisionSubject(
                "large-subject",
                "large-logical-subject",
                ComparisonAspect.Formula,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            var observation = new Observation(
                DecisionObservationId(
                    subject.SubjectId,
                    "master",
                    ObservationExactType.Formula,
                    contentHash),
                subject.SubjectId,
                "master",
                ObservationExactType.Formula,
                contentHash,
                formula: "=" + large,
                displayText: large);
            ReviewDecisionGeneration generation = ReviewDecisionEngine.Create(
                new ReviewDecisionCatalog(
                    evidence,
                    "master",
                    new[] { subject },
                    new[] { observation }));
            string path = CreateWorkbook(StateWithDecision(generation, "{}"));

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            using (Stream stream = GetDecisionPart(document.WorkbookPart!).GetStream(
                       FileMode.Open,
                       FileAccess.Read))
                Assert.True(stream.Length > 5L * 1024 * 1024);
            Assert.Equal(
                generation.GenerationId,
                ReviewWorkbookFile.ReadAndValidate(path).DecisionGeneration!.GenerationId);
        }

        [Fact]
        public void Writer_rejects_unverifiable_decision_generations_before_publication()
        {
            ReviewDecisionGeneration root = CreateDecisionRoot();
            ReviewWorkbookState staleEvidence = StateWithDecision(root, "{}");
            staleEvidence.DecisionEvidenceGenerationId = new string('0', 64);
            string staleEvidencePath = Path.Combine(_directory, "stale-decision-evidence.xlsx");
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateNew(
                staleEvidencePath,
                staleEvidence));
            Assert.False(File.Exists(staleEvidencePath));

            ReviewWorkbookState staleMapping = StateWithDecision(root, "{}");
            staleMapping.DecisionMappingGenerationId = "other-mapping-generation";
            string staleMappingPath = Path.Combine(_directory, "stale-decision-mapping.xlsx");
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateNew(
                staleMappingPath,
                staleMapping));
            Assert.False(File.Exists(staleMappingPath));

            ReviewDecisionCatalog source = root.Catalog;
            var opaqueCatalog = new ReviewDecisionCatalog(
                source.ProjectId,
                source.BaseGenerationId,
                source.MappingGenerationId,
                source.MasterRevisionId,
                source.Subjects,
                source.Observations,
                source.Groups);
            string opaquePath = Path.Combine(_directory, "opaque-evidence.xlsx");
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateNew(
                opaquePath,
                StateWithDecision(ReviewDecisionEngine.Create(opaqueCatalog), "{}")));
            Assert.False(File.Exists(opaquePath));

            var overlongCatalog = new ReviewDecisionCatalog(
                source.EvidenceGeneration!,
                source.MasterRevisionId,
                source.Subjects,
                source.Observations,
                source.Groups.Concat(new[]
                {
                    new ReviewDecisionGroup(new string('g', 5000), new[] { "subject-value" }),
                }));
            string overlongPath = Path.Combine(_directory, "overlong-decision-attribute.xlsx");
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateNew(
                overlongPath,
                StateWithDecision(ReviewDecisionEngine.Create(overlongCatalog), "{}")));
            Assert.False(File.Exists(overlongPath));
        }

        [Fact]
        public void Writer_rejects_an_oversized_decision_payload_and_cleans_the_temporary_package()
        {
            string contentHash = new string('7', 64);
            var evidence = new ReviewEvidenceGeneration(
                "project-oversized-decision-payload",
                "mapping-oversized-decision-payload",
                new[] { new ReviewEvidenceRevisionIdentity("master", 0, new string('6', 64)) });
            var subject = new ReviewDecisionSubject(
                "oversized-subject",
                "oversized-logical-subject",
                ComparisonAspect.Formula,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            var observation = new Observation(
                DecisionObservationId(
                    subject.SubjectId,
                    "master",
                    ObservationExactType.Formula,
                    contentHash),
                subject.SubjectId,
                "master",
                ObservationExactType.Formula,
                contentHash,
                formula: "=" + new string('q', ReviewWorkbookFile.MaxDecisionPayloadBytes));
            ReviewDecisionGeneration generation = ReviewDecisionEngine.Create(
                new ReviewDecisionCatalog(
                    evidence,
                    "master",
                    new[] { subject },
                    new[] { observation }));
            string path = Path.Combine(_directory, "oversized-decision-payload.xlsx");

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateNew(
                path,
                StateWithDecision(generation, "{}")));

            Assert.False(File.Exists(path));
            Assert.Empty(Directory.GetFiles(_directory, ".oversized-decision-payload.xlsx.*.tmp"));
        }

        [Fact]
        public void Writer_rejects_oversized_decision_event_audit_text_before_publication()
        {
            ReviewDecisionGeneration root = CreateDecisionRoot();
            ReviewDecisionGeneration advanced = ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    "subject-value",
                    ReviewDecisionStatus.Equivalent,
                    ReviewerSnapshot.Empty,
                    new DateTimeOffset(2026, 7, 17, 13, 0, 0, TimeSpan.Zero),
                    note: new string('n', ReviewWorkbookFile.MaxDecisionPayloadBytes + 1)));
            string path = Path.Combine(_directory, "oversized-decision-audit.xlsx");

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.CreateNew(
                path,
                StateWithDecision(advanced, "{}")));

            Assert.False(File.Exists(path));
            Assert.Empty(Directory.GetFiles(_directory, ".oversized-decision-audit.xlsx.*.tmp"));
        }

        [Fact]
        public void Reader_rejects_a_modified_state_chunk()
        {
            string path = CreateWorkbook(
                new ReviewWorkbookState("project-corrupt", "{\"important\":true}"),
                chunkSize: 5);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                CustomXmlPart chunk = document.WorkbookPart!.CustomXmlParts
                    .First(part => ReadRootLocalName(part) == "stateChunk");
                XDocument xml;
                using (Stream input = chunk.GetStream(FileMode.Open, FileAccess.Read))
                    xml = XDocument.Load(input);
                xml.Root!.Value = Convert.ToBase64String(new byte[] { 1, 2, 3, 4, 5 });
                using Stream output = chunk.GetStream(FileMode.Create, FileAccess.Write);
                xml.Save(output);
            }

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(path));
        }

        [Fact]
        public void Reader_rejects_visible_projection_tampering()
        {
            string path = CreateWorkbook(new ReviewWorkbookState("project-projection", "{\"state\":1}")
            {
                DetailBefore = "Before evidence",
                DetailAfter = "After evidence",
            });

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorksheetPart details = GetSheetPart(document.WorkbookPart!, "Details");
                Cell cell = details.Worksheet!.Descendants<Cell>()
                    .Single(candidate => candidate.CellReference!.Value == "B3");
                DocumentFormat.OpenXml.Spreadsheet.Text text = cell
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.Text>()
                    .First();
                text.Text = "Altered evidence";
                details.Worksheet.Save();
            }

            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(path));
            Assert.Contains("projection", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void New_manifest_binds_streaming_digest_version_and_cannot_be_downgraded_by_removal()
        {
            string path = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                CustomXmlPart manifest = document.WorkbookPart!.CustomXmlParts
                    .Single(part => ReadRootLocalName(part) == "reviewManifest");
                using Stream input = manifest.GetStream(FileMode.Open, FileAccess.Read);
                XElement root = XDocument.Load(input).Root!;
                Assert.Equal(
                    ReviewWorkbookFile.StreamingProjectionManifestSchemaVersion,
                    root.Attribute("schemaVersion")!.Value);
                Assert.Equal(
                    ReviewWorkbookFile.StreamingProjectionDigestVersion,
                    root.Attribute("projectionDigestVersion")!.Value);
            }

            string missingVersion = Path.Combine(_directory, "missing-projection-version.xlsx");
            File.Copy(path, missingVersion);
            MutateManifest(
                missingVersion,
                root => root.Attribute("projectionDigestVersion")!.Remove());
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(missingVersion));

            string downgrade = Path.Combine(_directory, "projection-downgrade.xlsx");
            File.Copy(path, downgrade);
            MutateManifest(downgrade, root =>
            {
                root.SetAttributeValue("schemaVersion", ReviewWorkbookFile.SchemaVersion);
                root.Attribute("projectionDigestVersion")!.Remove();
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(downgrade));
        }

        [Fact]
        public void Legacy_v1_projection_manifest_remains_readable_inside_the_original_limits()
        {
            string path = CreateWorkbook(new ReviewWorkbookState(
                "project-legacy-projection",
                "{\"legacy\":true}"));
            ConvertManifestToLegacyProjection(path);

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);

            Assert.Equal("project-legacy-projection", read.ProjectId);
            Assert.Equal("{\"legacy\":true}", read.StateJson);
        }

        [Fact]
        public void Legacy_projection_rejects_unchanged_digest_validation_formulas_and_concealing_views()
        {
            string validationPath = CreateWorkbook();
            ConvertManifestToLegacyProjection(validationPath);
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(validationPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                var validation = new DataValidation
                {
                    Type = DataValidationValues.Custom,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" },
                    ShowInputMessage = true,
                    Prompt = "Ignore the comparison report",
                };
                validation.Append(new Formula1("WEBSERVICE(\"https://example.invalid/\")"));
                changes.Worksheet!.Append(new DataValidations(validation) { Count = 1U });
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(validationPath));

            string viewportPath = CreateWorkbook();
            ConvertManifestToLegacyProjection(viewportPath);
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(viewportPath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                overview.Worksheet!.InsertBefore(
                    new SheetViews(new SheetView
                    {
                        WorkbookViewId = 0U,
                        TopLeftCell = "XFD1048576",
                    }),
                    overview.Worksheet.GetFirstChild<SheetData>());
                overview.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(viewportPath));

            string panePath = CreateWorkbook();
            ConvertManifestToLegacyProjection(panePath);
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(panePath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                var view = new SheetView { WorkbookViewId = 0U };
                view.Append(new Pane
                {
                    State = PaneStateValues.Frozen,
                    TopLeftCell = "A2",
                    VerticalSplit = 1D,
                });
                overview.Worksheet!.InsertBefore(
                    new SheetViews(view),
                    overview.Worksheet.GetFirstChild<SheetData>());
                overview.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(panePath));
        }

        [Fact]
        public void Streaming_projection_requires_canonical_monotonic_rows_and_cells()
        {
            var state = new ReviewWorkbookState("project-order", "{}");
            state.ChangeRows.Add(new ReviewChangeProjectionRow(
                "change-a", "Sheet1", "A1", "Value", "before-a", "after-a"));
            state.ChangeRows.Add(new ReviewChangeProjectionRow(
                "change-b", "Sheet1", "A2", "Value", "before-b", "after-b"));

            string rowPath = CreateWorkbook(state);
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(rowPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                SheetData data = changes.Worksheet!.GetFirstChild<SheetData>()!;
                List<Row> rows = data.Elements<Row>().ToList();
                rows[2].Remove();
                data.InsertBefore(rows[2], rows[1]);
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(rowPath));

            string cellPath = CreateWorkbook(state);
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(cellPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                Row header = changes.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First();
                List<Cell> cells = header.Elements<Cell>().Take(2).ToList();
                cells[1].Remove();
                header.InsertBefore(cells[1], cells[0]);
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(cellPath));
        }

        [Fact]
        public void Streaming_projection_normalizes_equivalent_inline_and_shared_string_text()
        {
            string path = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart!;
                SharedStringTablePart shared = workbookPart.AddNewPart<SharedStringTablePart>();
                shared.SharedStringTable = new SharedStringTable(
                    new SharedStringItem(
                        new DocumentFormat.OpenXml.Spreadsheet.Text("Sheet")),
                    new SharedStringItem(
                        new DocumentFormat.OpenXml.Spreadsheet.Text("Change ID")));
                shared.SharedStringTable.Save();

                WorksheetPart changes = GetSheetPart(workbookPart, "Changes");
                Cell first = changes.Worksheet!.Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "A1");
                first.RemoveAllChildren();
                first.DataType = CellValues.SharedString;
                first.CellValue = new CellValue("1");
                Cell second = changes.Worksheet.Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "B1");
                second.RemoveAllChildren();
                second.DataType = CellValues.SharedString;
                second.CellValue = new CellValue("0");
                changes.Worksheet.Save();
            }

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(2, read.ProjectionValidationMetrics.SharedStringItemCount);
            Assert.Equal(2, read.ProjectionValidationMetrics.SharedStringLookupCount);
        }

        [Fact]
        public void Streaming_shared_string_hot_cache_bounds_nonlocal_repeated_lookup_work()
        {
            const int rowCount = 5_000;
            var state = new ReviewWorkbookState("project-sst-cache", "{}");
            for (int index = 0; index < rowCount; index++)
            {
                state.ChangeRows.Add(new ReviewChangeProjectionRow(
                    "change-" + index.ToString(CultureInfo.InvariantCulture),
                    "Sheet1",
                    "A1",
                    "Value",
                    "",
                    "",
                    "",
                    "",
                    0,
                    0,
                    ""));
            }

            string path = CreateWorkbook(state);
            int sharedCellCount;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                sharedCellCount = ConvertSheetToSharedStrings(
                    document.WorkbookPart!,
                    "Changes");
            }

            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            ReviewWorkbookProjectionValidationMetrics metrics = read.ProjectionValidationMetrics;
            Assert.NotNull(metrics);
            Assert.Equal(
                ReviewWorkbookFile.StreamingProjectionDigestVersion,
                metrics.DigestVersion);
            Assert.Equal(sharedCellCount, metrics.SharedStringLookupCount);
            Assert.True(metrics.SharedStringCacheHitCount > sharedCellCount / 2);
            Assert.True(metrics.SharedStringDiskReadCount < sharedCellCount / 4);
            Assert.True(metrics.SharedStringRandomSeekCount < metrics.SharedStringDiskReadCount);
            Assert.InRange(
                metrics.PeakSharedStringCacheEntries,
                1,
                ReviewWorkbookFile.MaxStreamingSharedStringCacheEntries);
            Assert.InRange(
                metrics.PeakSharedStringCacheCharacters,
                1,
                ReviewWorkbookFile.MaxStreamingSharedStringCacheCharacters);
            Assert.Equal(16, metrics.WorksheetPassCount);
        }

        [Fact]
        public void Reader_accepts_Excels_redundant_internal_hyperlink_display_but_rejects_a_different_label()
        {
            string normalizedPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(normalizedPath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                Hyperlink link = overview.Worksheet!.Descendants<Hyperlink>()
                    .Single(candidate => candidate.Reference!.Value == "A7");
                link.Display = "Open Changes";
                overview.Worksheet.Save();
            }
            ReviewWorkbookFile.ReadAndValidate(normalizedPath);

            string alteredPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(alteredPath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                Hyperlink link = overview.Worksheet!.Descendants<Hyperlink>()
                    .Single(candidate => candidate.Reference!.Value == "A7");
                link.Display = "Misleading label";
                overview.Worksheet.Save();
            }
            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(alteredPath));
            Assert.Contains("projection", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Reader_rejects_renamed_or_hidden_projection_sheets()
        {
            string renamedPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(renamedPath, true))
            {
                Sheet sheet = document.WorkbookPart!.Workbook!.Descendants<Sheet>().First();
                sheet.Name = "Altered Overview";
                document.WorkbookPart.Workbook.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(renamedPath));

            string hiddenPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(hiddenPath, true))
            {
                Sheet sheet = document.WorkbookPart!.Workbook!.Descendants<Sheet>().First();
                sheet.State = SheetStateValues.Hidden;
                document.WorkbookPart.Workbook.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(hiddenPath));
        }

        [Fact]
        public void Reader_rejects_formulas_and_external_relationships_added_after_creation()
        {
            string formulaPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(formulaPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                Cell cell = changes.Worksheet!.Descendants<Cell>()
                    .Single(candidate => candidate.CellReference!.Value == "A1");
                cell.InlineString = null;
                cell.DataType = null;
                cell.CellFormula = new CellFormula("1+1");
                cell.CellValue = new CellValue("2");
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(formulaPath));

            string externalPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(externalPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                changes.AddHyperlinkRelationship(new Uri("https://example.invalid/"), true);
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(externalPath));
        }

        [Fact]
        public void Streaming_reader_rejects_type_inconsistent_cells_and_unbound_attributes()
        {
            string inconsistentPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(inconsistentPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                Cell cell = changes.Worksheet!.Descendants<Cell>()
                    .Single(candidate => candidate.CellReference!.Value == "A1");
                cell.Append(new CellValue("shadow value"));
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(inconsistentPath));

            string typePath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(typePath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                Cell cell = changes.Worksheet!.Descendants<Cell>()
                    .Single(candidate => candidate.CellReference!.Value == "A1");
                cell.SetAttribute(new OpenXmlAttribute("t", "", "future-type"));
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(typePath));

            string attributePath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(attributePath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                Cell cell = changes.Worksheet!.Descendants<Cell>()
                    .Single(candidate => candidate.CellReference!.Value == "A1");
                cell.SetAttribute(new OpenXmlAttribute("future", "", "1"));
                changes.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(attributePath));
        }

        [Fact]
        public void Streaming_reader_rejects_concealing_views_and_invalid_boolean_lexemes()
        {
            string viewPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(viewPath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                overview.Worksheet!.InsertBefore(
                    new SheetViews(new SheetView
                    {
                        WorkbookViewId = 0U,
                        TopLeftCell = "XFD1048576",
                        ZoomScale = 10U,
                    }),
                    overview.Worksheet.GetFirstChild<SheetData>());
                overview.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(viewPath));

            string hiddenWindowPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(hiddenWindowPath, true))
            {
                Workbook workbook = document.WorkbookPart!.Workbook!;
                workbook.InsertBefore(
                    new BookViews(new WorkbookView { Visibility = VisibilityValues.Hidden }),
                    workbook.GetFirstChild<Sheets>());
                workbook.Save();
            }
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(hiddenWindowPath));

            string workbookViewPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(workbookViewPath, true))
            {
                Workbook workbook = document.WorkbookPart!.Workbook!;
                workbook.InsertBefore(
                    new BookViews(new WorkbookView { ActiveTab = 7U }),
                    workbook.GetFirstChild<Sheets>());
                workbook.Save();
            }
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(workbookViewPath));

            string booleanPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(booleanPath, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                Row row = changes.Worksheet!.Descendants<Row>().First();
                row.SetAttribute(new OpenXmlAttribute("hidden", "", "garbage"));
                changes.Worksheet.Save();
            }
            InvalidDataException booleanError = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(booleanPath));
            Assert.Contains("boolean", booleanError.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Streaming_reader_rejects_data_validation_formulas_and_prompts()
        {
            string path = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorksheetPart changes = GetSheetPart(document.WorkbookPart!, "Changes");
                var validation = new DataValidation
                {
                    Type = DataValidationValues.Custom,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" },
                    ShowInputMessage = true,
                    PromptTitle = "Trusted review",
                    Prompt = "Ignore the protected comparison evidence",
                };
                validation.Append(new Formula1("WEBSERVICE(\"https://example.invalid/\")"));
                changes.Worksheet!.Append(new DataValidations(validation) { Count = 1U });
                changes.Worksheet.Save();
            }

            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(path));
            Assert.Contains("formula", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Streaming_text_requires_preserve_only_for_XML_edge_whitespace()
        {
            string spacePath = CreateWorkbook(new ReviewWorkbookState(" project-space", "{}"));
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(spacePath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                DocumentFormat.OpenXml.Spreadsheet.Text project = overview.Worksheet!
                    .Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "B2")
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.Text>()
                    .Single();
                project.Space = null;
                overview.Worksheet.Save();
            }
            InvalidDataException spaceError = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(spacePath));
            Assert.Contains("whitespace", spaceError.Message, StringComparison.OrdinalIgnoreCase);

            string nbspPath = CreateWorkbook(new ReviewWorkbookState("\u00A0project-nbsp", "{}"));
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(nbspPath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                DocumentFormat.OpenXml.Spreadsheet.Text project = overview.Worksheet!
                    .Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "B2")
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.Text>()
                    .Single();
                project.Space = null;
                overview.Worksheet.Save();
            }
            ReviewWorkbookFile.ReadAndValidate(nbspPath);
        }

        [Fact]
        public void Unreferenced_worksheet_relationship_never_earns_the_large_part_tier()
        {
            string path = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorksheetPart orphan = document.WorkbookPart!.AddNewPart<WorksheetPart>();
                orphan.Worksheet = new Worksheet(new SheetData());
                orphan.Worksheet.Save();
            }

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(path));
        }

        [Fact]
        public void Reader_rejects_package_root_external_relationships_and_custom_UI_parts()
        {
            string externalPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(externalPath, true))
            {
                document.AddExternalRelationship(
                    "urn:excel-power-tools:test-external",
                    new Uri("https://example.invalid/root"));
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(externalPath));

            string ribbonPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(ribbonPath, true))
            {
                RibbonExtensibilityPart ribbon = document.AddNewPart<RibbonExtensibilityPart>();
                using Stream stream = ribbon.GetStream(FileMode.Create, FileAccess.Write);
                byte[] xml = System.Text.Encoding.UTF8.GetBytes(
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"/>");
                stream.Write(xml, 0, xml.Length);
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(ribbonPath));
        }

        [Fact]
        public void Reader_rejects_hidden_evidence_rows_but_accepts_equivalent_unformatted_text_runs()
        {
            string hiddenPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(hiddenPath, true))
            {
                WorksheetPart overview = GetSheetPart(document.WorkbookPart!, "Overview");
                overview.Worksheet!.Descendants<Row>().First().Hidden = true;
                overview.Worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(hiddenPath));

            string equivalentPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(equivalentPath, true))
            {
                WorksheetPart details = GetSheetPart(document.WorkbookPart!, "Details");
                Cell cell = details.Worksheet!.Descendants<Cell>().Single();
                string displayedText = cell.InnerText;
                cell.InlineString = new InlineString(
                    new Run(new DocumentFormat.OpenXml.Spreadsheet.Text(displayedText)));
                details.Worksheet.Save();
            }

            ReviewWorkbookReadResult result = ReviewWorkbookFile.ReadAndValidate(equivalentPath);
            Assert.Equal("project-default", result.ProjectId);
        }

        [Fact]
        public void Reader_rejects_macro_enabled_content_type_even_without_a_VBA_part()
        {
            string path = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
                document.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(path));
        }

        [Fact]
        public void Reader_rejects_unknown_v1_manifest_grammar()
        {
            string path = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                CustomXmlPart manifest = document.WorkbookPart!.CustomXmlParts
                    .Single(part => ReadRootLocalName(part) == "reviewManifest");
                XDocument xml;
                using (Stream input = manifest.GetStream(FileMode.Open, FileAccess.Read))
                    xml = XDocument.Load(input);
                xml.Root!.SetAttributeValue("futureSilentMeaning", "not-allowed-in-v1");
                using Stream output = manifest.GetStream(FileMode.Create, FileAccess.Write);
                xml.Save(output);
            }

            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(path));
        }

        [Fact]
        public void Empty_state_has_no_fabricated_detail_and_invalid_XML_characters_are_visible_replacements()
        {
            string emptyPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(emptyPath, false))
            {
                WorksheetPart details = GetSheetPart(document.WorkbookPart!, "Details");
                Assert.Equal(
                    "No formula or text change is available for character-level detail.",
                    details.Worksheet!.Descendants<Cell>().Single().InnerText);
                Assert.DoesNotContain("Gross margin", details.Worksheet.InnerText, StringComparison.Ordinal);
            }

            string replacementPath = CreateWorkbook(new ReviewWorkbookState("project-control", "{}")
            {
                DetailBefore = "a\u0001b",
                DetailAfter = "ab",
            });
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(replacementPath, false))
            {
                WorksheetPart details = GetSheetPart(document.WorkbookPart!, "Details");
                Cell before = details.Worksheet!.Descendants<Cell>()
                    .Single(cell => cell.CellReference!.Value == "B3");
                Assert.Contains("a\uFFFDb", before.InnerText, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Detail_projection_enforces_the_32767_character_boundary_with_a_visible_marker()
        {
            string boundary = new string('a', ReviewWorkbookFile.MaxExcelCellTextCharacters);
            string boundaryJson = "{\"detail\":\"" + boundary + "\"}";
            string boundaryPath = CreateWorkbook(new ReviewWorkbookState("project-boundary", boundaryJson)
            {
                DetailBefore = boundary,
                DetailAfter = "",
            });

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(boundaryPath, false))
            {
                Cell before = GetSheetPart(document.WorkbookPart!, "Details").Worksheet!
                    .Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B3");
                Assert.Equal(ReviewWorkbookFile.MaxExcelCellTextCharacters, before.InnerText.Length);
                Assert.Equal(boundary, before.InnerText);
                Assert.DoesNotContain("truncated", before.InnerText, StringComparison.OrdinalIgnoreCase);
            }
            Assert.Equal(boundaryJson, ReviewWorkbookFile.ReadAndValidate(boundaryPath).StateJson);

            string overBoundary = new string('b', ReviewWorkbookFile.MaxExcelCellTextCharacters + 1);
            string overBoundaryAfter = new string('c', ReviewWorkbookFile.MaxExcelCellTextCharacters + 1);
            string overBoundaryJson = "{\"before\":\"" + overBoundary +
                                      "\",\"after\":\"" + overBoundaryAfter + "\"}";
            string overBoundaryPath = CreateWorkbook(new ReviewWorkbookState(
                "project-over-boundary",
                overBoundaryJson)
            {
                DetailBefore = overBoundary,
                DetailAfter = overBoundaryAfter,
            });

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(overBoundaryPath, false))
            {
                Cell before = GetSheetPart(document.WorkbookPart!, "Details").Worksheet!
                    .Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B3");
                Assert.Equal(ReviewWorkbookFile.MaxExcelCellTextCharacters, before.InnerText.Length);
                Assert.EndsWith(
                    "[truncated; full detail is preserved in embedded comparison state]",
                    before.InnerText,
                    StringComparison.Ordinal);
                Cell after = GetSheetPart(document.WorkbookPart!, "Details").Worksheet!
                    .Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B4");
                Assert.Equal(ReviewWorkbookFile.MaxExcelCellTextCharacters, after.InnerText.Length);
                Assert.EndsWith(
                    "[truncated; full detail is preserved in embedded comparison state]",
                    after.InnerText,
                    StringComparison.Ordinal);
            }
            Assert.Equal(overBoundaryJson, ReviewWorkbookFile.ReadAndValidate(overBoundaryPath).StateJson);
        }

        [Fact]
        public void Reader_preflight_rejects_duplicate_and_case_colliding_ZIP_entry_names()
        {
            string path = CreateWorkbook();
            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                ZipArchiveEntry duplicate = archive.CreateEntry("XL/WORKBOOK.XML");
                using Stream stream = duplicate.Open();
                stream.WriteByte(0);
            }

            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(path));
            Assert.Contains("duplicate", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Relationship_bound_streaming_worksheet_exceeds_legacy_tier_without_loosening_it()
        {
            const int rowCount = 100_000;
            var state = new ReviewWorkbookState("project-large-tier", "{}");
            var compact = new ReviewChangeProjectionRow(
                "x", "", "", "", "", "", "", "", 0, 0, "");
            for (int index = 0; index < rowCount; index++) state.ChangeRows.Add(compact);

            string path = CreateWorkbook(state);
            string changesEntryName;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
                changesEntryName = GetSheetPart(document.WorkbookPart!, "Changes")
                    .Uri.OriginalString.TrimStart('/');
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry changes = archive.GetEntry(changesEntryName)!;
                Assert.True(changes.Length > 64L * 1024 * 1024);
                Assert.True(changes.Length < ReviewWorkbookFile.MaxStreamingWorksheetCharacters);
            }
            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(0, read.ProjectionValidationMetrics.SharedStringItemCount);
            Assert.Equal(0, read.ProjectionValidationMetrics.SharedStringLookupCount);

            string legacy = Path.Combine(_directory, "large-tier-legacy.xlsx");
            File.Copy(path, legacy);
            MutateManifest(legacy, root =>
            {
                root.SetAttributeValue("schemaVersion", ReviewWorkbookFile.SchemaVersion);
                root.Attribute("projectionDigestVersion")!.Remove();
            });
            InvalidDataException legacyError = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(legacy));
            Assert.Contains("64 MiB", legacyError.Message, StringComparison.OrdinalIgnoreCase);

            string wrongContentType = Path.Combine(_directory, "large-tier-wrong-content-type.xlsx");
            File.Copy(path, wrongContentType);
            RewriteZipXmlEntry(wrongContentType, "[Content_Types].xml", xml =>
            {
                XElement worksheetOverride = xml.Root!.Elements()
                    .Single(element => string.Equals(
                        element.Attribute("PartName")?.Value,
                        "/" + changesEntryName,
                        StringComparison.Ordinal));
                worksheetOverride.SetAttributeValue("ContentType", "application/xml");
            });
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(wrongContentType));
        }

        [Fact]
        public void Relationship_bound_shared_strings_can_cross_64MiB_but_remain_version_and_type_gated()
        {
            const int itemCount = 2_100;
            const int charactersPerItem = ReviewWorkbookFile.MaxExcelCellTextCharacters;
            string path = CreateWorkbook();
            string sharedEntryName;
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                SharedStringTablePart shared = document.WorkbookPart!
                    .AddNewPart<SharedStringTablePart>();
                sharedEntryName = shared.Uri.OriginalString.TrimStart('/');
                WriteLargeSharedStringTable(shared, itemCount, charactersPerItem);
            }

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry shared = archive.GetEntry(sharedEntryName)!;
                Assert.True(shared.Length > 64L * 1024 * 1024);
                Assert.True(shared.Length < ReviewWorkbookFile.MaxStreamingSharedStringCharacters);
            }
            ReviewWorkbookReadResult read = ReviewWorkbookFile.ReadAndValidate(path);
            Assert.Equal(itemCount, read.ProjectionValidationMetrics.SharedStringItemCount);
            Assert.Equal(0, read.ProjectionValidationMetrics.SharedStringLookupCount);

            string legacy = Path.Combine(_directory, "large-sst-legacy.xlsx");
            File.Copy(path, legacy);
            MutateManifest(legacy, root =>
            {
                root.SetAttributeValue("schemaVersion", ReviewWorkbookFile.SchemaVersion);
                root.Attribute("projectionDigestVersion")!.Remove();
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(legacy));

            string wrongContentType = Path.Combine(_directory, "large-sst-wrong-content-type.xlsx");
            File.Copy(path, wrongContentType);
            RewriteZipXmlEntry(wrongContentType, "[Content_Types].xml", xml =>
            {
                XElement sharedOverride = xml.Root!.Elements()
                    .Single(element => string.Equals(
                        element.Attribute("PartName")?.Value,
                        "/" + sharedEntryName,
                        StringComparison.Ordinal));
                sharedOverride.SetAttributeValue("ContentType", "application/xml");
            });
            Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(wrongContentType));
        }

        [Fact]
        public void Reader_preflight_rejects_unsafe_ZIP_envelopes_and_declared_expansion()
        {
            byte[] valid = File.ReadAllBytes(CreateWorkbook());

            byte[] excessiveEntries = (byte[])valid.Clone();
            int eocd = FindEndOfCentralDirectory(excessiveEntries);
            WriteUInt16(excessiveEntries, eocd + 8, 10001);
            WriteUInt16(excessiveEntries, eocd + 10, 10001);
            AssertPreflightRejects(excessiveEntries, "entries.xlsx");

            byte[] excessiveCentralDirectory = (byte[])valid.Clone();
            eocd = FindEndOfCentralDirectory(excessiveCentralDirectory);
            WriteUInt32(excessiveCentralDirectory, eocd + 12, 16U * 1024U * 1024U + 1U);
            AssertPreflightRejects(excessiveCentralDirectory, "central-directory.xlsx");

            List<int> centralEntries = FindCentralDirectoryEntries(valid);
            Assert.True(centralEntries.Count >= 5);

            byte[] excessiveEntry = (byte[])valid.Clone();
            WriteUInt32(excessiveEntry, centralEntries[0] + 24, 64U * 1024U * 1024U + 1U);
            AssertPreflightRejects(excessiveEntry, "entry-expansion.xlsx");

            byte[] excessiveAggregate = (byte[])valid.Clone();
            foreach (int offset in centralEntries.Take(5))
                WriteUInt32(excessiveAggregate, offset + 24, 60U * 1024U * 1024U);
            AssertPreflightRejects(excessiveAggregate, "aggregate-expansion.xlsx");

            byte[] suspiciousRatio = (byte[])valid.Clone();
            WriteUInt32(suspiciousRatio, centralEntries[0] + 20, 1U);
            WriteUInt32(suspiciousRatio, centralEntries[0] + 24, 2U * 1024U * 1024U);
            AssertPreflightRejects(suspiciousRatio, "compression-ratio.xlsx");
        }

        [Fact]
        public void Reader_preflight_rejects_a_source_file_above_the_safe_package_limit()
        {
            string path = Path.Combine(_directory, "oversized-package.xlsx");
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                stream.SetLength(128L * 1024 * 1024 + 1);

            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(path));
            Assert.Contains("source-size", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Reader_rejects_tiny_dimensions_and_conditional_formatting_that_can_conceal_evidence()
        {
            string rowPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(rowPath, true))
            {
                Worksheet worksheet = GetSheetPart(document.WorkbookPart!, "Details").Worksheet!;
                Row row = worksheet.Descendants<Row>().First();
                row.Height = 0.5D;
                row.CustomHeight = true;
                worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(rowPath));

            string columnPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(columnPath, true))
            {
                Worksheet worksheet = GetSheetPart(document.WorkbookPart!, "Details").Worksheet!;
                worksheet.InsertBefore(
                    new Columns(new Column
                    {
                        Min = 2U,
                        Max = 2U,
                        Width = 0.5D,
                        CustomWidth = true,
                    }),
                    worksheet.GetFirstChild<SheetData>());
                worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(columnPath));

            string conditionalPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(conditionalPath, true))
            {
                Worksheet worksheet = GetSheetPart(document.WorkbookPart!, "Details").Worksheet!;
                worksheet.Append(new ConditionalFormatting
                {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B3" },
                });
                worksheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(conditionalPath));
        }

        [Fact]
        public void Reader_rejects_unsupported_rich_run_font_properties()
        {
            string path = CreateWorkbook(new ReviewWorkbookState("project-font", "{}")
            {
                DetailBefore = "before",
                DetailAfter = "after",
            });
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, true))
            {
                WorksheetPart details = GetSheetPart(document.WorkbookPart!, "Details");
                RunProperties properties = details.Worksheet!.Descendants<RunProperties>().First();
                properties.Append(new FontSize { Val = 1D });
                details.Worksheet.Save();
            }

            InvalidDataException error = Assert.Throws<InvalidDataException>(
                () => ReviewWorkbookFile.ReadAndValidate(path));
            Assert.Contains("presentation", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Reader_rejects_global_and_shared_style_parts_not_emitted_by_the_writer()
        {
            string stylesPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stylesPath, true))
            {
                WorkbookStylesPart styles = document.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = new Stylesheet(new Fonts(new Font(new Color { Rgb = "FFFFFFFF" })));
                styles.Stylesheet.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(stylesPath));

            string sharedStringsPath = CreateWorkbook();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(sharedStringsPath, true))
            {
                SharedStringTablePart shared = document.WorkbookPart!.AddNewPart<SharedStringTablePart>();
                shared.SharedStringTable = new SharedStringTable(
                    new SharedStringItem(
                        new Run(
                            new RunProperties(new FontSize { Val = 1D }),
                            new DocumentFormat.OpenXml.Spreadsheet.Text("concealed"))));
                shared.SharedStringTable.Save();
            }
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(sharedStringsPath));
        }

        [Fact]
        public void Writer_rejects_a_chunk_size_above_the_safe_limit()
        {
            string path = Path.Combine(_directory, "oversized-chunk.xlsx");
            var options = new ReviewWorkbookWriteOptions
            {
                StateChunkSizeBytes = ReviewWorkbookFile.MaxStateChunkSizeBytes + 1,
            };

            Assert.Throws<ArgumentOutOfRangeException>(() => ReviewWorkbookFile.CreateNew(
                path,
                new ReviewWorkbookState("project-limit", "{}"),
                options));
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void Workbook_is_macro_free_and_contains_no_external_relationships()
        {
            string path = CreateWorkbook();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
            {
                Assert.Equal(SpreadsheetDocumentType.Workbook, document.DocumentType);
                Assert.Null(document.WorkbookPart!.VbaProjectPart);
                Assert.Empty(document.WorkbookPart.ExternalWorkbookParts);
            }

            using ZipArchive archive = ZipFile.OpenRead(path);
            string contentTypes;
            using (var reader = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open()))
                contentTypes = reader.ReadToEnd();
            Assert.DoesNotContain("vbaProject", contentTypes, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("macroEnabled", contentTypes, StringComparison.OrdinalIgnoreCase);

            foreach (ZipArchiveEntry relationships in archive.Entries.Where(entry =>
                         entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
            {
                using Stream stream = relationships.Open();
                XDocument xml = XDocument.Load(stream);
                Assert.DoesNotContain(
                    xml.Descendants(),
                    element => string.Equals(
                        element.Attribute("TargetMode")?.Value,
                        "External",
                        StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Existing_destination_is_never_replaced()
        {
            string path = Path.Combine(_directory, "existing.xlsx");
            byte[] sentinel = { 10, 20, 30, 40 };
            File.WriteAllBytes(path, sentinel);

            Assert.Throws<IOException>(() => ReviewWorkbookFile.CreateNew(
                path,
                new ReviewWorkbookState("project-existing", "{}")));

            Assert.Equal(sentinel, File.ReadAllBytes(path));
        }

        [Fact]
        public void Generation_failure_leaves_final_absent_and_cleans_same_directory_temporary_file()
        {
            string path = Path.Combine(_directory, "failure.xlsx");
            var invalid = new ReviewWorkbookState("project-\u0001-invalid", "{}");

            Assert.ThrowsAny<Exception>(() => ReviewWorkbookFile.CreateNew(path, invalid));

            Assert.False(File.Exists(path));
            Assert.Empty(Directory.GetFiles(_directory, ".failure.xlsx.*.tmp"));
        }

        [Fact]
        public void Public_persistence_API_does_not_expose_OpenXml_types()
        {
            Type[] apiTypes = typeof(ReviewWorkbookFile).Assembly.GetExportedTypes()
                .Where(type => type.Namespace == typeof(ReviewWorkbookFile).Namespace)
                .ToArray();
            var signatures = new List<Type>();
            foreach (Type type in apiTypes)
            {
                foreach (MethodInfo method in type.GetMethods(
                             BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static |
                             BindingFlags.DeclaredOnly))
                {
                    signatures.Add(method.ReturnType);
                    signatures.AddRange(method.GetParameters().Select(parameter => parameter.ParameterType));
                }

                foreach (ConstructorInfo constructor in type.GetConstructors())
                    signatures.AddRange(constructor.GetParameters().Select(parameter => parameter.ParameterType));
            }

            Assert.DoesNotContain(signatures, ContainsOpenXmlType);
        }

        public void Dispose()
        {
            if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
        }

        private static ReviewDecisionGeneration CreateDecisionRoot()
        {
            var evidence = new ReviewEvidenceGeneration(
                "project-decision-sidecar",
                "mapping-decision-sidecar",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("candidate", 0, new string('a', 64)),
                    new ReviewEvidenceRevisionIdentity("master", 1, new string('b', 64)),
                });
            var subjects = new[]
            {
                new ReviewDecisionSubject(
                    "subject-formula",
                    "logical-formula",
                    ComparisonAspect.Formula,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable),
                new ReviewDecisionSubject(
                    "subject-value",
                    "logical-value",
                    ComparisonAspect.Value,
                    ReviewDecisionSubjectPolicy.SelectableAndMergeable),
            };
            var observations = new[]
            {
                new Observation(
                    DecisionObservationId(
                        "subject-formula", "candidate", ObservationExactType.Formula, new string('1', 64)),
                    "subject-formula",
                    "candidate",
                    ObservationExactType.Formula,
                    new string('1', 64),
                    formula: "=A1+1",
                    cachedResult: new CanonicalScalar(ObservationExactType.Number, "42"),
                    cachedResultTrust: CachedResultTrust.Trusted,
                    displayText: "=A1+1",
                    provenance: new ObservationProvenance(
                        "candidate.xlsx", "Input", "C7", 7, 3, "logical-sheet")),
                new Observation(
                    DecisionObservationId(
                        "subject-formula", "master", ObservationExactType.Formula, new string('2', 64)),
                    "subject-formula",
                    "master",
                    ObservationExactType.Formula,
                    new string('2', 64),
                    formula: "=A1",
                    provenance: new ObservationProvenance(
                        "master.xlsx", "Input", "C6", 6, 3, "logical-sheet")),
                new Observation(
                    DecisionObservationId(
                        "subject-value", "candidate", ObservationExactType.String, new string('3', 64)),
                    "subject-value",
                    "candidate",
                    ObservationExactType.String,
                    new string('3', 64),
                    value: new CanonicalScalar(ObservationExactType.String, "new\u0001value"),
                    displayText: "new value",
                    provenance: new ObservationProvenance(
                        "candidate.xlsx", "Input", "D7", 7, 4, "logical-sheet")),
                new Observation(
                    DecisionObservationId(
                        "subject-value", "master", ObservationExactType.String, new string('4', 64)),
                    "subject-value",
                    "master",
                    ObservationExactType.String,
                    new string('4', 64),
                    value: new CanonicalScalar(ObservationExactType.String, "old"),
                    provenance: new ObservationProvenance(
                        "master.xlsx", "Input", "D6", 6, 4, "logical-sheet")),
            };
            var catalog = new ReviewDecisionCatalog(
                evidence,
                "master",
                subjects,
                observations,
                new[]
                {
                    new ReviewDecisionGroup("group-formula", new[] { "subject-formula" }),
                    new ReviewDecisionGroup(
                        "group-formula-value",
                        subjects.Select(item => item.SubjectId)),
                });
            return ReviewDecisionEngine.Create(catalog);
        }

        private static ReviewDecisionGeneration CreateAdvancedDecisionGeneration()
        {
            ReviewDecisionGeneration root = CreateDecisionRoot();
            ReviewDecisionGroup group = root.Catalog.Groups
                .Single(item => item.GroupId == "group-formula-value");
            var groupReviewer = new ReviewerSnapshot(
                "Confirmed\u0001 reviewer",
                "Office reviewer",
                new[] { "office\u0002identity", "manual confirmation" });
            DateTimeOffset created = new DateTimeOffset(
                2026,
                7,
                17,
                12,
                0,
                0,
                TimeSpan.Zero);
            ReviewDecisionGeneration grouped = ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForGroup(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    group.GroupId,
                    group.SubjectIds,
                    ReviewDecisionStatus.NeedsInvestigation,
                    groupReviewer,
                    created,
                    runId: "run-group",
                    addInBuild: "build-group",
                    note: "group\u0004note"));

            string candidateObservationId = DecisionObservationId(
                "subject-formula",
                "candidate",
                ObservationExactType.Formula,
                new string('1', 64));
            ReviewDecisionGeneration selected = ReviewDecisionEngine.Apply(
                grouped,
                ReviewDecisionRequest.ForSubject(
                    grouped.ProjectId,
                    grouped.BaseGenerationId,
                    grouped.MappingGenerationId,
                    grouped.GenerationId,
                    "subject-formula",
                    ReviewDecisionStatus.UseCandidate,
                    new ReviewerSnapshot("Candidate reviewer", "", new[] { "manual" }),
                    created.AddMinutes(1),
                    selectedObservationId: candidateObservationId,
                    runId: "run-candidate",
                    addInBuild: "build\u0003candidate",
                    note: "selected\u0004candidate"));

            return ReviewDecisionEngine.Apply(
                selected,
                ReviewDecisionRequest.ClearSubject(
                    selected.ProjectId,
                    selected.BaseGenerationId,
                    selected.MappingGenerationId,
                    selected.GenerationId,
                    "subject-formula",
                    new ReviewerSnapshot("Final reviewer"),
                    created.AddMinutes(2),
                    runId: "run-clear",
                    addInBuild: "build-clear",
                    note: "cleared after review"));
        }

        private static void AssertDecisionGenerationEqual(
            ReviewDecisionGeneration expected,
            ReviewDecisionGeneration actual)
        {
            Assert.Equal(expected.CatalogId, actual.CatalogId);
            Assert.Equal(expected.GenerationId, actual.GenerationId);
            Assert.Equal(expected.PriorGenerationId, actual.PriorGenerationId);
            Assert.Equal(expected.Events.Count, actual.Events.Count);
            for (int index = 0; index < expected.Events.Count; index++)
                AssertDecisionEventEqual(expected.Events[index], actual.Events[index]);
            Assert.Equal(
                expected.SubjectStates.Keys.OrderBy(item => item, StringComparer.Ordinal),
                actual.SubjectStates.Keys.OrderBy(item => item, StringComparer.Ordinal));
            foreach (string subjectId in expected.SubjectStates.Keys)
            {
                ReviewDecisionState left = expected.SubjectStates[subjectId];
                ReviewDecisionState right = actual.SubjectStates[subjectId];
                Assert.Equal(left.Status, right.Status);
                Assert.Equal(left.SelectedObservationId, right.SelectedObservationId);
                Assert.Equal(left.EffectiveEventId, right.EffectiveEventId);
            }
        }

        private static void AssertDecisionEventEqual(DecisionEvent expected, DecisionEvent actual)
        {
            Assert.Equal(expected.EventId, actual.EventId);
            Assert.Equal(expected.RequestId, actual.RequestId);
            Assert.Equal(expected.ProjectId, actual.ProjectId);
            Assert.Equal(expected.EvidenceGenerationId, actual.EvidenceGenerationId);
            Assert.Equal(expected.MappingGenerationId, actual.MappingGenerationId);
            Assert.Equal(expected.PriorDecision, actual.PriorDecision);
            Assert.Equal(expected.NewDecision, actual.NewDecision);
            Assert.Equal(expected.CreatedUtc, actual.CreatedUtc);
            Assert.Equal(expected.SelectionScope, actual.SelectionScope);
            Assert.Equal(expected.RunId, actual.RunId);
            Assert.Equal(expected.SubjectId, actual.SubjectId);
            Assert.Equal(expected.GroupId, actual.GroupId);
            Assert.Equal(expected.AffectedSubjectIds, actual.AffectedSubjectIds);
            Assert.Equal(expected.SelectedObservationId, actual.SelectedObservationId);
            Assert.Equal(expected.SupportingRevisionIds, actual.SupportingRevisionIds);
            Assert.Equal(expected.SupportingContentHashes, actual.SupportingContentHashes);
            Assert.Equal(
                expected.ReviewerSnapshot.ConfirmedDisplayName,
                actual.ReviewerSnapshot.ConfirmedDisplayName);
            Assert.Equal(
                expected.ReviewerSnapshot.OfficeDisplayName,
                actual.ReviewerSnapshot.OfficeDisplayName);
            Assert.Equal(
                expected.ReviewerSnapshot.IdentitySourceLabels,
                actual.ReviewerSnapshot.IdentitySourceLabels);
            Assert.Equal(expected.AddInBuild, actual.AddInBuild);
            Assert.Equal(expected.Note, actual.Note);
            Assert.Equal(expected.SupersededEventId, actual.SupersededEventId);
            Assert.Equal(expected.Locations.Count, actual.Locations.Count);
            for (int index = 0; index < expected.Locations.Count; index++)
            {
                DecisionLocationContext left = expected.Locations[index];
                DecisionLocationContext right = actual.Locations[index];
                Assert.Equal(left.Role, right.Role);
                Assert.Equal(left.RevisionId, right.RevisionId);
                Assert.Equal(left.LogicalSheetId, right.LogicalSheetId);
                Assert.Equal(left.SheetName, right.SheetName);
                Assert.Equal(left.Address, right.Address);
                Assert.Equal(left.SourceLabel, right.SourceLabel);
                Assert.Equal(left.PhysicalRow, right.PhysicalRow);
                Assert.Equal(left.PhysicalColumn, right.PhysicalColumn);
            }
        }

        private static string DecisionObservationId(
            string subjectId,
            string revisionId,
            ObservationExactType exactType,
            string contentHash) =>
            ComparisonIds.Observation(subjectId, revisionId, exactType, contentHash);

        private static ReviewWorkbookState StateWithDecision(
            ReviewDecisionGeneration generation,
            string stateJson) =>
            new ReviewWorkbookState(generation.ProjectId, stateJson)
            {
                DecisionGeneration = generation,
                DecisionEvidenceGenerationId = generation.Catalog.EvidenceGenerationId,
                DecisionMappingGenerationId = generation.MappingGenerationId,
            };

        private string CreateDecisionWorkbook()
        {
            ReviewDecisionGeneration generation = CreateDecisionRoot();
            return CreateWorkbook(StateWithDecision(generation, "{\"decision\":true}"));
        }

        private string CreateAdvancedDecisionWorkbook()
        {
            ReviewDecisionGeneration generation = CreateAdvancedDecisionGeneration();
            return CreateWorkbook(StateWithDecision(generation, "{\"decision\":\"advanced\"}"));
        }

        private static CustomXmlPart GetDecisionPart(WorkbookPart workbookPart) =>
            workbookPart.CustomXmlParts.Single(part => ReadRootLocalName(part) == "decisionState");

        private static void ConvertManifestToLegacyProjection(string path)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = document.WorkbookPart!;
            MethodInfo method = typeof(ReviewWorkbookFile).GetMethod(
                "ComputeLegacyProjectionSha256",
                BindingFlags.NonPublic | BindingFlags.Static)!;
            string legacyDigest = Assert.IsType<string>(method.Invoke(null, new object[] { workbookPart }));
            CustomXmlPart manifest = workbookPart.CustomXmlParts
                .Single(part => ReadRootLocalName(part) == "reviewManifest");
            XDocument xml;
            using (Stream input = manifest.GetStream(FileMode.Open, FileAccess.Read))
                xml = XDocument.Load(input);
            xml.Root!.SetAttributeValue("schemaVersion", ReviewWorkbookFile.SchemaVersion);
            xml.Root.SetAttributeValue("projectionSha256", legacyDigest);
            xml.Root.Attribute("projectionDigestVersion")!.Remove();
            using Stream output = manifest.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(output, SaveOptions.DisableFormatting);
        }

        private static void MutateManifest(string path, Action<XElement> mutation)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            CustomXmlPart manifest = document.WorkbookPart!.CustomXmlParts
                .Single(part => ReadRootLocalName(part) == "reviewManifest");
            XDocument xml;
            using (Stream input = manifest.GetStream(FileMode.Open, FileAccess.Read))
                xml = XDocument.Load(input);
            mutation(xml.Root!);
            using Stream output = manifest.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(output, SaveOptions.DisableFormatting);
        }

        private static void MutateDecisionAndRebind(string path, Action<XElement> mutation)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = document.WorkbookPart!;
            CustomXmlPart decision = GetDecisionPart(workbookPart);
            XDocument decisionXml;
            using (Stream input = decision.GetStream(FileMode.Open, FileAccess.Read))
                decisionXml = XDocument.Load(input);
            mutation(decisionXml.Root!);
            using (Stream output = decision.GetStream(FileMode.Create, FileAccess.Write))
                decisionXml.Save(output, SaveOptions.DisableFormatting);

            RebindDecisionDigest(workbookPart, decision);
        }

        private static void InjectDecisionElementsAndRebind(string path, int count)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = document.WorkbookPart!;
            CustomXmlPart decision = GetDecisionPart(workbookPart);
            string xml;
            using (Stream input = decision.GetStream(FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(input, new UTF8Encoding(false, true)))
                xml = reader.ReadToEnd();
            const string closing = "</eptd:observations>";
            int insertionPoint = xml.IndexOf(closing, StringComparison.Ordinal);
            Assert.True(insertionPoint >= 0);
            var rewritten = new StringBuilder(xml.Length + checked(count * 15));
            rewritten.Append(xml, 0, insertionPoint);
            for (int index = 0; index < count; index++)
                rewritten.Append("<eptd:future/>");
            rewritten.Append(xml, insertionPoint, xml.Length - insertionPoint);
            byte[] bytes = new UTF8Encoding(false, true).GetBytes(rewritten.ToString());
            using (Stream output = decision.GetStream(FileMode.Create, FileAccess.Write))
                output.Write(bytes, 0, bytes.Length);

            RebindDecisionDigest(workbookPart, decision);
        }

        private static void RebindDecisionDigest(
            WorkbookPart workbookPart,
            CustomXmlPart decision)
        {

            string digest;
            using (Stream input = decision.GetStream(FileMode.Open, FileAccess.Read))
            using (SHA256 sha256 = SHA256.Create())
                digest = string.Concat(sha256.ComputeHash(input)
                    .Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));

            CustomXmlPart manifest = workbookPart.CustomXmlParts
                .Single(part => ReadRootLocalName(part) == "reviewManifest");
            XDocument manifestXml;
            using (Stream input = manifest.GetStream(FileMode.Open, FileAccess.Read))
                manifestXml = XDocument.Load(input);
            manifestXml.Root!.SetAttributeValue("decisionSha256", digest);
            using Stream manifestOutput = manifest.GetStream(FileMode.Create, FileAccess.Write);
            manifestXml.Save(manifestOutput, SaveOptions.DisableFormatting);
        }

        private static void AddDuplicateDecisionPart(string path)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Open(path, true);
            WorkbookPart workbookPart = document.WorkbookPart!;
            XDocument xml;
            using (Stream input = GetDecisionPart(workbookPart).GetStream(FileMode.Open, FileAccess.Read))
                xml = XDocument.Load(input);

            string itemId = Guid.NewGuid().ToString("B").ToUpperInvariant();
            xml.Root!.SetAttributeValue("itemId", itemId);
            CustomXmlPart duplicate = workbookPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            CustomXmlPropertiesPart properties = duplicate.AddNewPart<CustomXmlPropertiesPart>();
            var references = new SchemaReferences();
            references.Append(new SchemaReference { Uri = ReviewWorkbookFile.DecisionCustomXmlNamespace });
            var dataStoreItem = new DataStoreItem { ItemId = itemId };
            dataStoreItem.Append(references);
            properties.DataStoreItem = dataStoreItem;
            using Stream output = duplicate.GetStream(FileMode.Create, FileAccess.Write);
            xml.Save(output, SaveOptions.DisableFormatting);
        }

        private void AssertDecisionReorderRejected(string containerName, string childName)
        {
            string path = CreateDecisionWorkbook();
            MutateDecisionAndRebind(path, root =>
            {
                XElement container = root.Elements()
                    .Single(element => element.Name.LocalName == containerName);
                Assert.All(
                    container.Elements(),
                    element => Assert.Equal(childName, element.Name.LocalName));
                ReverseFirstTwoChildren(container);
            });
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(path));
        }

        private static void ReverseFirstTwoChildren(XElement container)
        {
            List<XElement> children = container.Elements().Take(2).ToList();
            Assert.Equal(2, children.Count);
            children[0].Remove();
            children[1].AddAfterSelf(children[0]);
        }

        private string CreateWorkbook(ReviewWorkbookState? state = null, int? chunkSize = null)
        {
            string path = Path.Combine(_directory, Guid.NewGuid().ToString("N") + ".xlsx");
            state ??= new ReviewWorkbookState("project-default", "{\"stage\":\"inventory\"}");
            var options = new ReviewWorkbookWriteOptions();
            if (chunkSize.HasValue) options.StateChunkSizeBytes = chunkSize.Value;
            ReviewWorkbookFile.CreateNew(path, state, options);
            return path;
        }

        private static WorksheetPart GetSheetPart(WorkbookPart workbookPart, string name)
        {
            Sheet sheet = workbookPart.Workbook!.Descendants<Sheet>()
                .Single(candidate => candidate.Name!.Value == name);
            return Assert.IsType<WorksheetPart>(workbookPart.GetPartById(sheet.Id!.Value!));
        }

        private static string ReadRootLocalName(CustomXmlPart part)
        {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream).Root!.Name.LocalName;
        }

        private static void RewriteZipXmlEntry(
            string path,
            string entryName,
            Action<XDocument> mutation)
        {
            using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry(entryName)!;
            XDocument xml;
            using (Stream input = entry.Open()) xml = XDocument.Load(input);
            mutation(xml);
            using Stream output = entry.Open();
            output.SetLength(0);
            xml.Save(output, SaveOptions.DisableFormatting);
        }

        private static int ConvertSheetToSharedStrings(WorkbookPart workbookPart, string sheetName)
        {
            WorksheetPart worksheet = GetSheetPart(workbookPart, sheetName);
            SharedStringTablePart shared = workbookPart.SharedStringTablePart ??
                                           workbookPart.AddNewPart<SharedStringTablePart>();
            var table = new SharedStringTable();
            var indices = new Dictionary<string, int>(StringComparer.Ordinal);
            List<Cell> cells = worksheet.Worksheet!.Descendants<Cell>().ToList();
            foreach (Cell cell in cells)
            {
                string value = cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? "";
                if (!indices.TryGetValue(value, out int index))
                {
                    index = indices.Count;
                    indices.Add(value, index);
                    table.Append(new SharedStringItem(
                        new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
                }
                cell.RemoveAllChildren();
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
            }
            shared.SharedStringTable = table;
            shared.SharedStringTable.Save();
            worksheet.Worksheet.Save();
            return cells.Count;
        }

        private static void WriteLargeSharedStringTable(
            SharedStringTablePart part,
            int itemCount,
            int charactersPerItem)
        {
            const string spreadsheetNamespace =
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string alphabet =
                "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
            var buffer = new char[charactersPerItem];
            uint random = 0x9E3779B9U;
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                CloseOutput = false,
                CheckCharacters = true,
            });
            writer.WriteStartDocument(true);
            writer.WriteStartElement("sst", spreadsheetNamespace);
            writer.WriteAttributeString("count", itemCount.ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("uniqueCount", itemCount.ToString(CultureInfo.InvariantCulture));
            for (int item = 0; item < itemCount; item++)
            {
                for (int index = 0; index < buffer.Length; index++)
                {
                    random ^= random << 13;
                    random ^= random >> 17;
                    random ^= random << 5;
                    buffer[index] = alphabet[(int)(random % (uint)alphabet.Length)];
                }
                writer.WriteStartElement("si", spreadsheetNamespace);
                writer.WriteStartElement("t", spreadsheetNamespace);
                writer.WriteChars(buffer, 0, buffer.Length);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.WriteEndDocument();
        }

        private void AssertPreflightRejects(byte[] bytes, string fileName)
        {
            string path = Path.Combine(_directory, fileName);
            File.WriteAllBytes(path, bytes);
            Assert.Throws<InvalidDataException>(() => ReviewWorkbookFile.ReadAndValidate(path));
        }

        private static int FindEndOfCentralDirectory(byte[] bytes)
        {
            for (int offset = bytes.Length - 22; offset >= Math.Max(0, bytes.Length - 65557); offset--)
            {
                if (ReadUInt32(bytes, offset) == 0x06054B50U &&
                    offset + 22 + ReadUInt16(bytes, offset + 20) == bytes.Length)
                    return offset;
            }
            throw new InvalidOperationException("Fixture ZIP has no end-of-central-directory record.");
        }

        private static List<int> FindCentralDirectoryEntries(byte[] bytes)
        {
            int eocd = FindEndOfCentralDirectory(bytes);
            int count = ReadUInt16(bytes, eocd + 10);
            int offset = checked((int)ReadUInt32(bytes, eocd + 16));
            var entries = new List<int>(count);
            for (int index = 0; index < count; index++)
            {
                if (ReadUInt32(bytes, offset) != 0x02014B50U)
                    throw new InvalidOperationException("Fixture ZIP central directory is malformed.");
                entries.Add(offset);
                offset += 46 + ReadUInt16(bytes, offset + 28) +
                          ReadUInt16(bytes, offset + 30) + ReadUInt16(bytes, offset + 32);
            }
            return entries;
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) =>
            (ushort)(bytes[offset] | bytes[offset + 1] << 8);

        private static uint ReadUInt32(byte[] bytes, int offset) =>
            (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 |
                   bytes[offset + 3] << 24);

        private static void WriteUInt16(byte[] bytes, int offset, ushort value)
        {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteUInt32(byte[] bytes, int offset, uint value)
        {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static bool ContainsOpenXmlType(Type type)
        {
            if (type.Namespace?.StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal) == true)
                return true;
            if (type.HasElementType && type.GetElementType() != null)
                return ContainsOpenXmlType(type.GetElementType()!);
            return type.IsGenericType && type.GetGenericArguments().Any(ContainsOpenXmlType);
        }
    }
}
