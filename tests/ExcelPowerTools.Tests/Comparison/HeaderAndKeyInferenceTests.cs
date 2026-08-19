using System;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class HeaderAndKeyInferenceTests
    {
        [Fact]
        public void Strong_header_requires_discriminating_evidence_and_excludes_only_the_accepted_row()
        {
            AlignmentTable table = Table(
                new[] { "ID", "Name", "Amount" },
                new[] { "A-001", "Alpha", "10" },
                new[] { "A-002", "Beta", "20" });

            HeaderDetectionResult result = HeaderDetector.Detect(table);

            Assert.Equal(HeaderDecision.StrongHeader, result.Decision);
            Assert.Equal(0, result.CandidateRowIndex);
            Assert.Equal(0, result.AcceptedHeaderRowIndex);
            Assert.Equal(1, result.DataStartRowIndex);
            Assert.True(result.Confidence >= 0.70D);
            Assert.Contains(result.Evidence, reason => reason.Code == "header.strong");
        }

        [Fact]
        public void Early_wordy_data_is_not_promoted_to_a_header()
        {
            AlignmentTable table = Table(
                new[] { "Paris", "France" },
                new[] { "Rome", "Italy" },
                new[] { "Madrid", "Spain" });

            HeaderDetectionResult result = HeaderDetector.Detect(table);

            Assert.False(result.HasAcceptedHeader);
            Assert.Equal(0, result.DataStartRowIndex);
            Assert.Equal(0, result.CandidateRowIndex);
            Assert.True(result.Decision == HeaderDecision.Data || result.Decision == HeaderDecision.Ambiguous);
            Assert.Contains(result.Evidence, reason => reason.Code.EndsWith("treated-as-data", StringComparison.Ordinal));
        }

        [Fact]
        public void Competing_header_candidates_remain_ambiguous_and_both_rows_remain_data()
        {
            AlignmentTable table = Table(
                new[] { "ID", "Name" },
                new[] { "Code", "Description" },
                new[] { "A-001", "Alpha" },
                new[] { "A-002", "Beta" });

            HeaderDetectionResult result = HeaderDetector.Detect(table);

            Assert.Equal(HeaderDecision.Ambiguous, result.Decision);
            Assert.Null(result.AcceptedHeaderRowIndex);
            Assert.Equal(0, result.DataStartRowIndex);
            Assert.True(result.RunnerUpMargin < 0.12D);
        }

        [Fact]
        public void Explicit_confirmation_accepts_a_single_column_header_without_changing_exact_evidence()
        {
            string[] sourceHeader = { "  Part ID  " };
            AlignmentTable table = Table(sourceHeader, new[] { "001" }, new[] { "002" });
            var options = new HeaderDetectionOptions { ConfirmedHeaderRowIndex = 0 };

            HeaderDetectionResult result = HeaderDetector.Detect(table, options: options);
            sourceHeader[0] = "mutated outside";

            Assert.Equal(HeaderDecision.ConfirmedHeader, result.Decision);
            Assert.Equal(1, result.DataStartRowIndex);
            Assert.Equal("  Part ID  ", table.Rows[0].Values[0]);
            Assert.Equal(AlignmentConfidence.Confirmed,
                result.Decision == HeaderDecision.ConfirmedHeader
                    ? AlignmentConfidence.Confirmed
                    : AlignmentConfidence.None);
        }

        [Fact]
        public void Commented_header_can_be_strong_but_the_comment_prefix_is_only_normalized_for_inference()
        {
            AlignmentTable table = Table(
                new[] { "#ID", "Name" },
                new[] { "A-001", "Alpha" },
                new[] { "A-002", "Beta" });

            HeaderDetectionResult result = HeaderDetector.Detect(
                table,
                options: new HeaderDetectionOptions { CommentPrefix = "#" });

            Assert.Equal(HeaderDecision.StrongHeader, result.Decision);
            Assert.Equal("#ID", table.Rows[0].Values[0]);
            Assert.Equal(1, result.DataStartRowIndex);
        }

        [Fact]
        public void Inferred_key_maps_reordered_columns_and_prefers_the_id_alias()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Name" },
                new[] { "A-001", "Alpha" },
                new[] { "A-002", "Beta" },
                new[] { "A-003", "Gamma" });
            AlignmentTable after = Table(
                new[] { "Name", "ID" },
                new[] { "Alpha", "A-001" },
                new[] { "Beta", "A-002" },
                new[] { "Delta", "A-004" });

            KeyInferenceResult result = KeyColumnInference.Infer(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            Assert.Equal(KeyInferenceDecision.StrongKey, result.Decision);
            KeyCandidate selected = Assert.IsType<KeyCandidate>(result.SelectedCandidate);
            Assert.Equal(new[] { 0 }, selected.BeforeColumns);
            Assert.Equal(new[] { 1 }, selected.AfterColumns);
            Assert.False(selected.HasDuplicates);
            Assert.True(selected.CrossRevisionOverlap >= 0.60D);
        }

        [Fact]
        public void Duplicate_candidate_values_are_reported_and_not_selected_as_an_inferred_key()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Status" },
                new[] { "A", "Open" },
                new[] { "A", "Open" },
                new[] { "B", "Closed" });
            AlignmentTable after = Table(
                new[] { "ID", "Status" },
                new[] { "A", "Open" },
                new[] { "A", "Open" },
                new[] { "B", "Closed" });

            KeyInferenceResult result = KeyColumnInference.Infer(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            Assert.Equal(KeyInferenceDecision.Ambiguous, result.Decision);
            Assert.Null(result.SelectedCandidate);
            Assert.Contains(result.Candidates, candidate =>
                candidate.BeforeColumns.SequenceEqual(new[] { 0 }) && candidate.HasDuplicates);
            Assert.Contains(result.Reasons, reason => reason.Code == "key.duplicates");
        }

        [Fact]
        public void Composite_key_is_inferred_only_when_no_single_column_is_unique()
        {
            AlignmentTable before = Table(
                new[] { "Country", "Order" },
                new[] { "US", "1" },
                new[] { "US", "2" },
                new[] { "IT", "1" });
            AlignmentTable after = Table(
                new[] { "Order", "Country" },
                new[] { "1", "US" },
                new[] { "1", "IT" },
                new[] { "2", "US" });

            KeyInferenceResult result = KeyColumnInference.Infer(
                before,
                after,
                ConfirmedHeader(before),
                ConfirmedHeader(after));

            Assert.Equal(KeyInferenceDecision.StrongKey, result.Decision);
            Assert.Equal(new[] { 0, 1 }, result.SelectedCandidate.BeforeColumns);
            Assert.Equal(new[] { 1, 0 }, result.SelectedCandidate.AfterColumns);
            Assert.False(result.SelectedCandidate.HasDuplicates);
        }

        [Fact]
        public void Key_inference_does_not_run_on_a_weak_header_candidate()
        {
            AlignmentTable before = Table(new[] { "Paris", "France" }, new[] { "Rome", "Italy" });
            AlignmentTable after = Table(new[] { "Paris", "France" }, new[] { "Rome", "Italy" });
            HeaderDetectionResult beforeHeader = HeaderDetector.Detect(before);
            HeaderDetectionResult afterHeader = HeaderDetector.Detect(after);

            KeyInferenceResult result = KeyColumnInference.Infer(before, after, beforeHeader, afterHeader);

            Assert.Equal(KeyInferenceDecision.HeaderRequired, result.Decision);
            Assert.Empty(result.Candidates);
            Assert.Null(result.SelectedCandidate);
        }

        private static HeaderDetectionResult ConfirmedHeader(AlignmentTable table) =>
            HeaderDetector.Detect(
                table,
                options: new HeaderDetectionOptions { ConfirmedHeaderRowIndex = 0 });

        private static AlignmentTable Table(params string[][] rows) => new AlignmentTable(rows);
    }
}
