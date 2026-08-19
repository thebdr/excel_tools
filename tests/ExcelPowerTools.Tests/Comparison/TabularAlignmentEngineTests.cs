using System;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class TabularAlignmentEngineTests
    {
        [Fact]
        public void Strong_key_alignment_absorbs_an_inserted_row_without_downstream_cascade()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Name" },
                new[] { "A", "Alpha" },
                new[] { "B", "Beta" },
                new[] { "C", "Gamma" });
            AlignmentTable after = Table(
                new[] { "ID", "Name" },
                new[] { "A", "Alpha" },
                new[] { "X", "Inserted" },
                new[] { "B", "Beta" },
                new[] { "C", "Gamma" });
            HeaderDetectionResult beforeHeader = ConfirmedHeader(before);
            HeaderDetectionResult afterHeader = ConfirmedHeader(after);
            KeyInferenceResult key = KeyColumnInference.Infer(
                before, after, beforeHeader, afterHeader);

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(
                before, after, beforeHeader, afterHeader, key);

            Assert.Equal(AxisAlignmentStrategy.InferredKey, result.Strategy);
            AssertMapped(result, 1, 1);
            AssertMapped(result, 2, 3);
            AssertMapped(result, 3, 4);
            AxisAlignmentLink inserted = Assert.Single(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter);
            Assert.Equal(2, inserted.AfterIndex);
            Assert.Empty(result.Ambiguities);
        }

        [Fact]
        public void Exact_key_matching_does_not_trim_or_case_fold_source_values()
        {
            AlignmentTable before = Table(
                new[] { "ID" },
                new[] { "A" },
                new[] { "B" },
                new[] { "C" },
                new[] { "D" },
                new[] { "E" });
            AlignmentTable after = Table(
                new[] { "ID" },
                new[] { "A" },
                new[] { " B " },
                new[] { "C" },
                new[] { "D" },
                new[] { "E" });
            HeaderDetectionResult beforeHeader = ConfirmedHeader(before);
            HeaderDetectionResult afterHeader = ConfirmedHeader(after);
            KeyInferenceResult key = KeyColumnInference.Infer(before, after, beforeHeader, afterHeader);

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(
                before, after, beforeHeader, afterHeader, key);

            Assert.Equal(KeyInferenceDecision.StrongKey, key.Decision);
            Assert.DoesNotContain(result.Links,
                link => link.Kind == AxisAlignmentKind.Matched && link.BeforeIndex == 2);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyBefore && link.BeforeIndex == 2);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter && link.AfterIndex == 2);
        }

        [Fact]
        public void Confirmed_duplicate_keys_create_an_ambiguity_block_not_arbitrary_matches()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Value" },
                new[] { "A", "one" },
                new[] { "A", "two" },
                new[] { "B", "three" });
            AlignmentTable after = Table(
                new[] { "ID", "Value" },
                new[] { "A", "changed one" },
                new[] { "A", "changed two" },
                new[] { "B", "three" });
            HeaderDetectionResult beforeHeader = ConfirmedHeader(before);
            HeaderDetectionResult afterHeader = ConfirmedHeader(after);
            var keyOptions = new KeyInferenceOptions
            {
                ConfirmedBeforeColumns = new[] { 0 },
                ConfirmedAfterColumns = new[] { 0 },
            };
            KeyInferenceResult key = KeyColumnInference.Infer(
                before, after, beforeHeader, afterHeader, keyOptions);

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(
                before, after, beforeHeader, afterHeader, key);

            AxisAlignmentAmbiguity ambiguity = Assert.Single(result.Ambiguities);
            Assert.Equal(new[] { 1, 2 }, ambiguity.BeforeIndices);
            Assert.Equal(new[] { 1, 2 }, ambiguity.AfterIndices);
            Assert.DoesNotContain(result.Links, link => link.BeforeIndex == 1 || link.BeforeIndex == 2);
            AssertMapped(result, 3, 3);
            Assert.Contains(ambiguity.Reasons, reason => reason.Code == "row.duplicate-key-ambiguity");
        }

        [Fact]
        public void Unique_row_content_anchors_handle_insertion_without_a_key()
        {
            AlignmentTable before = Table(
                new[] { "A", "1" },
                new[] { "B", "2" },
                new[] { "C", "3" });
            AlignmentTable after = Table(
                new[] { "A", "1" },
                new[] { "X", "new" },
                new[] { "B", "2" },
                new[] { "C", "3" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(before, after);

            AssertMapped(result, 0, 0);
            AssertMapped(result, 1, 2);
            AssertMapped(result, 2, 3);
            Assert.Single(result.Links, link => link.Kind == AxisAlignmentKind.OnlyAfter);
        }

        [Fact]
        public void A_single_changed_row_between_two_exact_anchors_is_neighbor_aligned()
        {
            AlignmentTable before = Table(
                new[] { "anchor-a" },
                new[] { "old content" },
                new[] { "anchor-c" });
            AlignmentTable after = Table(
                new[] { "anchor-a" },
                new[] { "entirely new content" },
                new[] { "anchor-c" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(before, after);

            AxisAlignmentLink changed = Assert.Single(result.Links,
                link => link.BeforeIndex == 1 && link.AfterIndex == 1);
            Assert.Equal(AlignmentConfidence.Medium, changed.Confidence);
            Assert.Equal(AxisAlignmentStrategy.NeighborBoundedSequence, changed.Strategy);
            Assert.Contains(changed.Reasons, reason => reason.Code == "row.two-sided-neighbor-bounds");
        }

        [Fact]
        public void Duplicate_row_occurrences_remain_one_explicit_many_to_many_ambiguity()
        {
            AlignmentTable before = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "X" },
                new[] { "B" });
            AlignmentTable after = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "X" },
                new[] { "X" },
                new[] { "B" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(before, after);

            AxisAlignmentAmbiguity ambiguity = Assert.Single(result.Ambiguities);
            Assert.Equal(new[] { 1, 2 }, ambiguity.BeforeIndices);
            Assert.Equal(new[] { 1, 2, 3 }, ambiguity.AfterIndices);
            Assert.DoesNotContain(result.Links,
                link => link.BeforeIndex == 1 || link.BeforeIndex == 2 ||
                        link.AfterIndex == 1 || link.AfterIndex == 2 || link.AfterIndex == 3);
        }

        [Fact]
        public void Unique_exact_rows_can_be_mapped_across_a_reorder_and_are_labelled_reordered()
        {
            AlignmentTable before = Table(new[] { "A" }, new[] { "B" }, new[] { "C" });
            AlignmentTable after = Table(new[] { "B" }, new[] { "A" }, new[] { "C" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(before, after);

            AssertMapped(result, 0, 1);
            AssertMapped(result, 1, 0);
            Assert.True(result.Links.Single(link => link.BeforeIndex == 0).Reordered);
            Assert.True(result.Links.Single(link => link.BeforeIndex == 1).Reordered);
            Assert.False(result.Links.Single(link => link.BeforeIndex == 2).Reordered);
        }

        [Fact]
        public void Strong_headers_align_an_inserted_column_without_downstream_cascade()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Name", "Amount" },
                new[] { "A", "Alpha", "10" });
            AlignmentTable after = Table(
                new[] { "ID", "Notes", "Name", "Amount" },
                new[] { "A", "new", "Alpha", "10" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            Assert.Equal(AxisAlignmentStrategy.HeaderLabels, result.Strategy);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 1, 2);
            AssertMapped(result, 2, 3);
            Assert.Equal(1, Assert.Single(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter).AfterIndex);
        }

        [Fact]
        public void Duplicate_header_labels_are_not_paired_by_occurrence()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Value", "Value" },
                new[] { "A", "1", "2" });
            AlignmentTable after = Table(
                new[] { "ID", "Value", "Value" },
                new[] { "A", "2", "1" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            AxisAlignmentAmbiguity ambiguity = Assert.Single(result.Ambiguities);
            Assert.Equal(new[] { 1, 2 }, ambiguity.BeforeIndices);
            Assert.Equal(new[] { 1, 2 }, ambiguity.AfterIndices);
            AssertMapped(result, 0, 0);
        }

        [Fact]
        public void Blank_header_cells_are_left_unmapped_even_when_their_physical_positions_match()
        {
            AlignmentTable before = Table(
                new[] { "ID", "" },
                new[] { "A", "before" });
            AlignmentTable after = Table(
                new[] { "ID", "" },
                new[] { "A", "after" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            AssertMapped(result, 0, 0);
            Assert.DoesNotContain(result.Links,
                link => link.Kind == AxisAlignmentKind.Matched && link.BeforeIndex == 1);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyBefore && link.BeforeIndex == 1);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter && link.AfterIndex == 1);
        }

        [Fact]
        public void Exact_column_content_anchors_align_columns_when_no_header_is_accepted()
        {
            AlignmentTable before = Table(
                new[] { "A1", "B1" },
                new[] { "A2", "B2" });
            AlignmentTable after = Table(
                new[] { "A1", "X1", "B1" },
                new[] { "A2", "X2", "B2" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(before, after);

            Assert.Equal(AxisAlignmentStrategy.UniqueContentAnchors, result.Strategy);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 1, 2);
            Assert.Equal(1, Assert.Single(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter).AfterIndex);
        }

        private static void AssertMapped(AxisAlignmentResult result, int before, int after)
        {
            Assert.True(result.TryMapBefore(before, out int actual));
            Assert.Equal(after, actual);
        }

        private static HeaderDetectionResult ConfirmedHeader(AlignmentTable table) =>
            HeaderDetector.Detect(
                table,
                options: new HeaderDetectionOptions { ConfirmedHeaderRowIndex = 0 });

        private static AlignmentTable Table(params string[][] rows) => new AlignmentTable(rows);
    }
}
