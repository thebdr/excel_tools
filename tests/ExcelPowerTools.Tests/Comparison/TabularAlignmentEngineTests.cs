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
        public void Repeated_identical_rows_align_monotonically_with_only_surplus_after()
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

            Assert.Empty(result.Ambiguities);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 1, 1);
            AssertMapped(result, 2, 2);
            AssertMapped(result, 3, 4);
            AxisAlignmentLink surplus = Assert.Single(
                result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter);
            Assert.Equal(3, surplus.AfterIndex);
            Assert.DoesNotContain(result.Links, link => link.Kind == AxisAlignmentKind.OnlyBefore);
        }

        [Fact]
        public void Repeated_identical_rows_align_monotonically_with_only_surplus_before()
        {
            AlignmentTable before = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "X" },
                new[] { "X" },
                new[] { "B" });
            AlignmentTable after = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "X" },
                new[] { "B" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(before, after);

            Assert.Empty(result.Ambiguities);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 1, 1);
            AssertMapped(result, 2, 2);
            AssertMapped(result, 4, 3);
            AxisAlignmentLink surplus = Assert.Single(
                result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyBefore);
            Assert.Equal(3, surplus.BeforeIndex);
            Assert.DoesNotContain(result.Links, link => link.Kind == AxisAlignmentKind.OnlyAfter);
        }

        [Fact]
        public void Reordered_repeated_row_group_crossing_a_unique_anchor_remains_ambiguous()
        {
            AlignmentTable before = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "X" },
                new[] { "Y" },
                new[] { "B" });
            AlignmentTable after = Table(
                new[] { "A" },
                new[] { "Y" },
                new[] { "X" },
                new[] { "X" },
                new[] { "B" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignRows(before, after);

            AxisAlignmentAmbiguity ambiguity = Assert.Single(result.Ambiguities);
            Assert.Equal(new[] { 1, 2 }, ambiguity.BeforeIndices);
            Assert.Equal(new[] { 2, 3 }, ambiguity.AfterIndices);
            Assert.DoesNotContain(result.Links,
                link => link.BeforeIndex == 1 || link.BeforeIndex == 2 ||
                        link.AfterIndex == 2 || link.AfterIndex == 3);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 4, 4);
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
        public void Blank_header_column_with_unique_neighbors_and_ninety_percent_similarity_aligns()
        {
            AlignmentTable before = Table(
                new[] { "ID", "", "Amount" },
                new[] { "A", "one", "1" },
                new[] { "B", "two", "2" },
                new[] { "C", "three", "3" },
                new[] { "D", "four", "4" },
                new[] { "E", "five", "5" },
                new[] { "F", "six", "6" },
                new[] { "G", "seven", "7" },
                new[] { "H", "eight", "8" },
                new[] { "I", "nine", "9" },
                new[] { "J", "ten", "10" });
            AlignmentTable after = Table(
                new[] { "ID", "", "Amount" },
                new[] { "A", "one", "1" },
                new[] { "B", "two", "2" },
                new[] { "C", "three", "3" },
                new[] { "D", "four", "4" },
                new[] { "E", "five", "5" },
                new[] { "F", "six", "6" },
                new[] { "G", "seven", "7" },
                new[] { "H", "eight", "8" },
                new[] { "I", "nine", "9" },
                new[] { "J", "changed", "10" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            Assert.Empty(result.Ambiguities);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 1, 1);
            AssertMapped(result, 2, 2);
        }

        [Fact]
        public void Competing_blank_header_targets_fail_closed_as_unmatched()
        {
            AlignmentTable before = Table(
                new[] { "ID", "", "Amount" },
                new[] { "A", "one", "1" },
                new[] { "B", "two", "2" },
                new[] { "C", "three", "3" },
                new[] { "D", "four", "4" },
                new[] { "E", "five", "5" },
                new[] { "F", "six", "6" },
                new[] { "G", "seven", "7" },
                new[] { "H", "eight", "8" },
                new[] { "I", "nine", "9" },
                new[] { "J", "ten", "10" });
            AlignmentTable after = Table(
                new[] { "ID", "", "", "Amount" },
                new[] { "A", "one", "one", "1" },
                new[] { "B", "two", "two", "2" },
                new[] { "C", "three", "three", "3" },
                new[] { "D", "four", "four", "4" },
                new[] { "E", "five", "five", "5" },
                new[] { "F", "six", "six", "6" },
                new[] { "G", "seven", "seven", "7" },
                new[] { "H", "eight", "eight", "8" },
                new[] { "I", "changed-left", "nine", "9" },
                new[] { "J", "ten", "changed-right", "10" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            Assert.Empty(result.Ambiguities);
            AssertMapped(result, 0, 0);
            AssertMapped(result, 2, 3);
            Assert.DoesNotContain(result.Links,
                link => link.Kind == AxisAlignmentKind.Matched && link.BeforeIndex == 1);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyBefore && link.BeforeIndex == 1);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter && link.AfterIndex == 1);
            Assert.Contains(result.Links,
                link => link.Kind == AxisAlignmentKind.OnlyAfter && link.AfterIndex == 2);
        }

        [Fact]
        public void Blank_header_column_with_low_similarity_remains_unmapped()
        {
            AlignmentTable before = Table(
                new[] { "ID", "", "Amount" },
                new[] { "A", "before", "1" });
            AlignmentTable after = Table(
                new[] { "ID", "", "Amount" },
                new[] { "A", "after", "1" });

            AxisAlignmentResult result = TabularAlignmentEngine.AlignColumns(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            AssertMapped(result, 0, 0);
            AssertMapped(result, 2, 2);
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
