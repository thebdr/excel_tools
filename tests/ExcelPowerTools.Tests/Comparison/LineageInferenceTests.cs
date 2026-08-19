using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Lineage;
using ExcelPowerTools.Core.Comparison.Revision;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class LineageInferenceTests
    {
        [Fact]
        public void Strong_tokens_produce_deterministic_order_parent_proposals_and_explicit_selections()
        {
            LineageRevisionEvidence r1 = Evidence("r1", "Plant_r1.xlsx", 'A', Day(1));
            LineageRevisionEvidence r21 = Evidence("r21", "Plant v2.1.xlsx", 'B', Day(2));
            LineageRevisionEvidence r3 = Evidence("r3", "Plant_rev3.xlsx", 'C', Day(3));
            var selections = new LineageExplicitSelections("r3", "r1");

            LineageInferenceResult first = MultiRevisionLineageInference.Infer(
                new[] { r3, r1, r21 },
                selections);
            LineageInferenceResult second = MultiRevisionLineageInference.Infer(
                new[] { r21, r3, r1 },
                selections);

            Assert.Equal(new[] { "r1", "r21", "r3" }, first.Proposals.Select(value => value.RevisionId));
            Assert.Equal("r1", first.GetProposal("r21").ProposedParentRevisionId);
            Assert.Equal("r21", first.GetProposal("r3").ProposedParentRevisionId);
            Assert.Equal(LineageConfidenceLevel.Strong, first.GetProposal("r3").OrderConfidence.Level);
            Assert.Equal(LineageConfidenceLevel.Moderate, first.GetProposal("r3").ParentConfidence.Level);
            Assert.True(first.GetProposal("r3").ParentConfidence.RequiresConfirmation);
            Assert.Equal("r3", first.ExplicitMasterRevisionId);
            Assert.Equal("r1", first.ExplicitTrustedBaselineRevisionId);
            Assert.Contains(first.Reasons, reason => reason.Code == LineageReasonCode.ExplicitMasterSelection);
            Assert.Contains(first.Reasons, reason => reason.Code == LineageReasonCode.ExplicitBaselineSelection);
            Assert.Equal(first.ProposalGenerationId, second.ProposalGenerationId);
            Assert.Equal(
                first.Proposals.Select(value => value.RevisionId),
                second.Proposals.Select(value => value.RevisionId));
        }

        [Fact]
        public void Terminal_number_ambiguity_is_preserved_and_never_promoted_to_strong_order_evidence()
        {
            LineageInferenceResult result = MultiRevisionLineageInference.Infer(new[]
            {
                Evidence("draft", "Plant draft.xlsx", 'A', Day(1)),
                Evidence("numbered", "Plant 31.xlsx", 'B', Day(2)),
            });

            LineageRevisionProposal proposal = result.GetProposal("numbered");
            Assert.False(proposal.LabelInference.HasStrongCandidate);
            Assert.Contains(proposal.LabelInference.Candidates, candidate =>
                candidate.Kind == RevisionTokenKind.TrailingWholeNumber && candidate.Number.ToString() == "31");
            Assert.Contains(proposal.LabelInference.Candidates, candidate =>
                candidate.Kind == RevisionTokenKind.TrailingSplitDecimal && candidate.Number.ToString() == "3.1");
            Assert.Contains(proposal.OrderReasons, reason => reason.Code == LineageReasonCode.TerminalNumberAmbiguity);
            Assert.Equal(LineageConfidenceLevel.Weak, proposal.OrderConfidence.Level);
            Assert.Equal("draft", proposal.ProposedParentRevisionId);
            Assert.Equal(LineageConfidenceLevel.Weak, proposal.ParentConfidence.Level);
        }

        [Fact]
        public void Exact_hash_duplicates_remain_distinct_and_do_not_fabricate_parentage()
        {
            LineageInferenceResult result = MultiRevisionLineageInference.Infer(new[]
            {
                Evidence("r1-a", "Plant_r1.xlsx", 'A', Day(1)),
                Evidence("r1-b", "Plant_r1 copy.xlsx", 'A', Day(2)),
                Evidence("r2", "Plant_r2.xlsx", 'B', Day(3)),
            });

            LineageDuplicateGroup duplicate = Assert.Single(result.DuplicateGroups);
            Assert.Equal(Hash('A'), duplicate.ByteContentId);
            Assert.Equal(new[] { "r1-a", "r1-b" }, duplicate.RevisionIds);
            Assert.Equal(3, result.Proposals.Count);
            Assert.Null(result.GetProposal("r1-a").ProposedParentRevisionId);
            Assert.Null(result.GetProposal("r1-b").ProposedParentRevisionId);
            Assert.Null(result.GetProposal("r2").ProposedParentRevisionId);
            Assert.Contains(result.GetProposal("r2").ParentReasons, reason =>
                reason.Code == LineageReasonCode.AmbiguousParent);
            Assert.All(new[] { result.GetProposal("r1-a"), result.GetProposal("r1-b") }, proposal =>
            {
                Assert.Single(proposal.ExactDuplicateRevisionIds);
                Assert.Contains(proposal.OrderReasons, reason =>
                    reason.Code == LineageReasonCode.ExactDuplicateContent &&
                    reason.Confidence.Level == LineageConfidenceLevel.Certain);
                Assert.True(proposal.OrderConfidence.RequiresConfirmation);
            });
        }

        [Fact]
        public void Timestamps_only_produce_weak_order_and_parent_evidence()
        {
            LineageInferenceResult result = MultiRevisionLineageInference.Infer(new[]
            {
                Evidence("later", "Later draft.xlsx", 'B', Day(8)),
                Evidence("earlier", "Earlier draft.xlsx", 'A', Day(5)),
            });

            Assert.Equal(new[] { "earlier", "later" }, result.Proposals.Select(value => value.RevisionId));
            LineageRevisionProposal later = result.GetProposal("later");
            Assert.Equal(LineageConfidenceLevel.Weak, later.OrderConfidence.Level);
            Assert.Equal("earlier", later.ProposedParentRevisionId);
            Assert.Equal(0.20, later.ParentConfidence.Score, 6);
            Assert.Contains(later.ParentReasons, reason =>
                reason.Code == LineageReasonCode.WeakTimestampParentProposal);
            Assert.All(result.Proposals, proposal => Assert.True(proposal.OrderConfidence.RequiresConfirmation));
        }

        [Fact]
        public void Explicit_token_order_wins_over_conflicting_weak_timestamps_and_reports_the_conflict()
        {
            LineageInferenceResult result = MultiRevisionLineageInference.Infer(new[]
            {
                Evidence("r1", "Plant_r1.xlsx", 'A', Day(9)),
                Evidence("r2", "Plant_r2.xlsx", 'B', Day(1)),
            });

            Assert.Equal(new[] { "r1", "r2" }, result.Proposals.Select(value => value.RevisionId));
            Assert.All(result.Proposals, proposal =>
            {
                Assert.Contains(proposal.OrderReasons, reason => reason.Code == LineageReasonCode.ConflictingTimestamp);
                Assert.Equal(LineageConfidenceLevel.Moderate, proposal.OrderConfidence.Level);
                Assert.True(proposal.OrderConfidence.RequiresConfirmation);
            });
        }

        [Fact]
        public void Master_and_baseline_are_never_inferred_and_explicit_ids_must_exist()
        {
            LineageRevisionEvidence only = Evidence("r1", "Plant_r1.xlsx", 'A', Day(1));

            LineageInferenceResult result = MultiRevisionLineageInference.Infer(new[] { only });

            Assert.Null(result.ExplicitMasterRevisionId);
            Assert.Null(result.ExplicitTrustedBaselineRevisionId);
            Assert.Equal(2, result.Reasons.Count(reason => reason.Code == LineageReasonCode.SelectionRequired));
            Assert.Throws<ArgumentException>(() => MultiRevisionLineageInference.Infer(
                new[] { only },
                new LineageExplicitSelections("missing", "r1")));

            LineageInferenceResult sameSelection = MultiRevisionLineageInference.Infer(
                new[] { only },
                new LineageExplicitSelections("r1", "r1"));
            Assert.Equal("r1", sameSelection.ExplicitMasterRevisionId);
            Assert.Equal("r1", sameSelection.ExplicitTrustedBaselineRevisionId);
        }

        [Fact]
        public void Proposal_collections_are_immutable_and_input_paths_are_reduced_to_basenames()
        {
            var source = new LineageRevisionEvidence(
                "r1",
                @"C:\private\folder\Plant_r1.xlsx",
                Hash('A'),
                modifiedUtc: Day(1));
            LineageInferenceResult result = MultiRevisionLineageInference.Infer(new[] { source });

            Assert.Equal("Plant_r1.xlsx", source.FileName);
            Assert.Throws<NotSupportedException>(() =>
                ((IList<LineageRevisionProposal>)result.Proposals).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<LineageReason>)result.GetProposal("r1").OrderReasons).Clear());
        }

        private static LineageRevisionEvidence Evidence(
            string revisionId,
            string fileName,
            char hashCharacter,
            DateTimeOffset modifiedUtc) =>
            new LineageRevisionEvidence(
                revisionId,
                fileName,
                Hash(hashCharacter),
                createdUtc: modifiedUtc.AddHours(-1),
                modifiedUtc: modifiedUtc);

        private static string Hash(char value) => new string(value, 64);

        private static DateTimeOffset Day(int day) =>
            new DateTimeOffset(2026, 1, day, 12, 0, 0, TimeSpan.Zero);
    }
}
