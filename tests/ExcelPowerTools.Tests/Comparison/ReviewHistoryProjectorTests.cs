using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Persistence;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewHistoryProjectorTests
    {
        [Fact]
        public void Projects_one_collapsed_row_per_subject_with_endpoint_text_flags_and_counts()
        {
            LogicalHistoryResult history = LogicalHistoryEngine.Build(
                "project-history-review",
                "mapping-history-review",
                new[]
                {
                    Revision("r1", "", 0, true, false,
                        Text("A1", "ABC"), Text("A2", "old"), Formula("B1", "=1+1", 2)),
                    Revision("r2", "r1", 1, false, false,
                        Text("A1", "abc"), Text("A2", "middle"), Formula("B1", "=1+2", 3)),
                    Revision("r3", "r2", 2, false, true,
                        Text("A1", "ABC"), Text("A2", "master"), Formula("B1", "=1+3", 4)),
                    Revision("branch", "r1", 3, false, false,
                        Text("A1", "other"), Text("A2", "branch"), Formula("B1", "=1+4", 5)),
                });

            ReviewWorkbookState state = ReviewHistoryProjector.Project(
                history,
                "{\"schema\":\"logical-history-v1\"}",
                new[]
                {
                    RevisionDisplay("branch", 3, "r1", "branch.xlsx"),
                    RevisionDisplay("r3", 2, "r2", @"\\server\share\master.xlsx"),
                    RevisionDisplay("r1", 0, "", @"C:\private\base.xlsx"),
                    RevisionDisplay("r2", 1, "r1", "/private/mail/r2.xlsx"),
                },
                new[]
                {
                    Display("B1", "Model", "B1", AlignmentConfidence.Medium),
                    Display("A2", "Model", "A2", AlignmentConfidence.Confirmed),
                    Display("A1", "Model", "A1", AlignmentConfidence.Strong),
                });

            Assert.Equal("project-history-review", state.ProjectId);
            Assert.Equal("{\"schema\":\"logical-history-v1\"}", state.StateJson);
            Assert.Equal(3, state.ChangeRows.Count);

            ReviewChangeProjectionRow reverted = Assert.Single(state.ChangeRows, row => row.Address == "A1");
            Assert.Equal("Value", reverted.Aspect);
            Assert.Equal("ABC", reverted.BeforeText);
            Assert.Equal("ABC", reverted.AfterText);
            Assert.Equal("Equivalent; Reverted; Conflict; Branch-only", reverted.Flags);
            Assert.Equal(4, reverted.RevisionCount);
            Assert.Equal(3, reverted.DistinctStateCount);
            Assert.Equal("Strong", reverted.AlignmentConfidence);

            ReviewChangeProjectionRow changed = Assert.Single(state.ChangeRows, row => row.Address == "A2");
            Assert.Equal("old", changed.BeforeText);
            Assert.Equal("master", changed.AfterText);
            Assert.Equal("Conflict; Branch-only", changed.Flags);
            Assert.Equal("Confirmed", changed.AlignmentConfidence);

            ReviewChangeProjectionRow formula = Assert.Single(state.ChangeRows, row => row.Address == "B1");
            Assert.Equal("Formula", formula.Aspect);
            Assert.Equal("=1+1", formula.BeforeText);
            Assert.Equal("=1+3", formula.AfterText);

            // A changed primary value is a better representative character detail than a reverted endpoint or
            // a formula/cached-result secondary candidate.
            Assert.Equal("old", state.DetailBefore);
            Assert.Equal("master", state.DetailAfter);
        }

        [Fact]
        public void Populates_revision_graph_roles_in_order_and_retains_original_source_paths()
        {
            LogicalHistoryResult history = LogicalHistoryEngine.Build(
                "project-portable",
                "mapping-portable",
                new[]
                {
                    Revision("root", "", 0, true, false, Text("A1", "old")),
                    Revision("master", "root", 1, false, true, Text("A1", "master")),
                    Revision("branch", "root", 2, false, false, Text("A1", "branch")),
                });
            DateTimeOffset modified = new DateTimeOffset(2026, 7, 17, 10, 11, 12, TimeSpan.FromHours(2));

            ReviewWorkbookState state = ReviewHistoryProjector.Project(
                history,
                "{}",
                new[]
                {
                    new ReviewRevisionDisplayMetadata(
                        "branch", "Branch copy", 2, "root", "/srv/private/branch.xlsx"),
                    new ReviewRevisionDisplayMetadata(
                        "master", "Reviewed", 1, "root", @"\\server\hidden\master.xlsm"),
                    new ReviewRevisionDisplayMetadata(
                        "root", "Initial", 0, "", @"C:\Users\secret\root.xlsx",
                        "ABCDEF", modified, "Explicit order", "Confirmed"),
                },
                new[] { Display("A1", "Input", "A1", AlignmentConfidence.Confirmed) });

            Assert.Equal(new[] { "root", "master", "branch" },
                state.RevisionRows.Select(row => row.RevisionId));
            Assert.Equal(new[] { "Trusted baseline", "Master", "Branch" },
                state.RevisionRows.Select(row => row.Role));
            Assert.Equal(new[] { "root.xlsx", "master.xlsm", "branch.xlsx" },
                state.RevisionRows.Select(row => row.SourceBaseName));
            Assert.Equal(
                new[] { @"C:\Users\secret\root.xlsx", @"\\server\hidden\master.xlsm", "/srv/private/branch.xlsx" },
                state.RevisionRows.Select(row => row.SourcePath));
            Assert.Equal("2026-07-17T08:11:12.0000000+00:00", state.RevisionRows[0].ModifiedUtc);
            Assert.Equal("Explicit order", state.RevisionRows[0].Inference);
            Assert.Equal("Confirmed", state.RevisionRows[0].Confidence);

            string visibleRevisionProjection = string.Join("|", state.RevisionRows.Select(row => string.Join("|", new[]
            {
                row.RevisionId,
                row.Label,
                row.ParentRevisionId,
                row.Role,
                row.SourceBaseName,
                row.SourcePath,
                row.ByteContentId,
                row.ModifiedUtc,
                row.Inference,
                row.Confidence,
            })));
            Assert.Contains(@"C:\Users\secret", visibleRevisionProjection, StringComparison.Ordinal);
            Assert.Contains(@"\\server\hidden", visibleRevisionProjection, StringComparison.Ordinal);
            Assert.Contains("/srv/private", visibleRevisionProjection, StringComparison.Ordinal);
        }

        [Fact]
        public void Preserves_a_single_history_group_id_on_each_collapsed_member()
        {
            LogicalHistoryResult history = LogicalHistoryEngine.Build(
                "project-group",
                "mapping-group",
                new[]
                {
                    Revision("r1", "", 0, true, false, Text("A1", "old"), Text("A2", "old")),
                    Revision("r2", "r1", 1, false, true, Text("A1", "new"), Text("A2", "new")),
                });
            HistoryChangeGroup group = Assert.Single(history.Groups);

            ReviewWorkbookState state = ReviewHistoryProjector.Project(
                history,
                "{}",
                new[]
                {
                    RevisionDisplay("r2", 1, "r1", "r2.xlsx"),
                    RevisionDisplay("r1", 0, "", "r1.xlsx"),
                },
                new[]
                {
                    Display("A2", "Input", "A2", AlignmentConfidence.Confirmed),
                    Display("A1", "Input", "A1", AlignmentConfidence.Confirmed),
                });

            Assert.Equal(2, state.ChangeRows.Count);
            Assert.All(state.ChangeRows, row => Assert.Equal(group.GroupId, row.GroupId));
            Assert.Equal(
                history.Subjects.Select(subject => subject.SubjectId).OrderBy(id => id, StringComparer.Ordinal),
                state.ChangeRows.Select(row => row.ChangeId).OrderBy(id => id, StringComparer.Ordinal));
        }

        [Fact]
        public void Uses_the_first_revision_when_no_trusted_baseline_exists()
        {
            LogicalHistoryResult history = LogicalHistoryEngine.Build(
                "project-first",
                "mapping-first",
                new[]
                {
                    Revision("r1", "", 0, false, false, Text("A1", "first")),
                    Revision("r2", "r1", 1, false, false, Text("A1", "middle")),
                    Revision("r3", "r2", 2, false, true, Text("A1", "master")),
                });

            ReviewWorkbookState state = ReviewHistoryProjector.Project(
                history,
                "{}",
                new[]
                {
                    RevisionDisplay("r1", 0, "", "r1.xlsx"),
                    RevisionDisplay("r2", 1, "r1", "r2.xlsx"),
                    RevisionDisplay("r3", 2, "r2", "r3.xlsx"),
                },
                new[] { Display("A1", "Input", "A1", AlignmentConfidence.Weak) });

            ReviewChangeProjectionRow row = Assert.Single(state.ChangeRows);
            Assert.Equal("first", row.BeforeText);
            Assert.Equal("master", row.AfterText);
        }

        [Fact]
        public void Rejects_stale_or_incomplete_display_metadata_instead_of_fabricating_projection_evidence()
        {
            LogicalHistoryResult history = LogicalHistoryEngine.Build(
                "project-validation",
                "mapping-validation",
                new[]
                {
                    Revision("r1", "", 0, true, false, Text("A1", "old")),
                    Revision("r2", "r1", 1, false, true, Text("A1", "new")),
                });
            ReviewRevisionDisplayMetadata[] revisions =
            {
                RevisionDisplay("r1", 0, "", "r1.xlsx"),
                RevisionDisplay("r2", 1, "r1", "r2.xlsx"),
            };

            Assert.Throws<ArgumentException>(() => ReviewHistoryProjector.Project(
                history,
                "{}",
                revisions,
                Array.Empty<ReviewSubjectDisplayMetadata>()));

            Assert.Throws<ArgumentException>(() => ReviewHistoryProjector.Project(
                history,
                "{}",
                new[]
                {
                    RevisionDisplay("r1", 0, "", "r1.xlsx"),
                    RevisionDisplay("r2", 1, "", "r2.xlsx"),
                },
                new[] { Display("A1", "Input", "A1", AlignmentConfidence.Strong) }));

            Assert.Throws<ArgumentException>(() => new ReviewRevisionDisplayMetadata(
                "r1", "v1", 0, "", @"C:\private\directory\"));
        }

        private static ReviewRevisionDisplayMetadata RevisionDisplay(
            string revisionId,
            int order,
            string parentRevisionId,
            string sourceFileName) =>
            new ReviewRevisionDisplayMetadata(
                revisionId,
                revisionId.ToUpperInvariant(),
                order,
                parentRevisionId,
                sourceFileName,
                new string((char)('A' + order), 64),
                new DateTimeOffset(2026, 1, order + 1, 0, 0, 0, TimeSpan.Zero),
                "Confirmed order",
                "Confirmed");

        private static ReviewSubjectDisplayMetadata Display(
            string key,
            string sheet,
            string address,
            AlignmentConfidence confidence) =>
            new ReviewSubjectDisplayMetadata(key, sheet, address, confidence);

        private static HistoryRevision Revision(
            string revisionId,
            string parentRevisionId,
            int order,
            bool baseline,
            bool master,
            params CanonicalCellState[] cells) =>
            new HistoryRevision(
                revisionId,
                parentRevisionId,
                order,
                new RevisionSnapshot(revisionId, cells),
                master,
                baseline);

        private static CanonicalCellState Text(string key, string value) =>
            CanonicalCellState.Constant(key, CanonicalValue.Text(value));

        private static CanonicalCellState Formula(string key, string formula, double cached) =>
            CanonicalCellState.Formula(key, formula, CanonicalValue.Number(cached));
    }
}
