using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.History;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class LogicalHistoryEngineTests
    {
        [Fact]
        public void Builds_adjacent_histories_and_marks_a_master_lineage_revert()
        {
            LogicalHistoryResult result = LogicalHistoryEngine.Build(
                "project",
                "mapping-1",
                new[]
                {
                    Revision("r1", "", 0, "A1", "one", baseline: true),
                    Revision("r2", "r1", 1, "A1", "two"),
                    Revision("r3", "r2", 2, "A1", "one", master: true),
                });

            LogicalSubjectHistory value = Assert.Single(result.Subjects);
            Assert.Equal(CellChangeAspect.Value, value.Aspect);
            Assert.True(value.IsReverted);
            Assert.Equal(3, value.Observations.Count);
            Assert.Equal(2, value.CandidateStates.Count);
            Assert.Equal(2, value.Transitions.Count);
            Assert.True(value.Transitions[1].IsRevert);
            Assert.All(value.Transitions, transition => Assert.True(transition.IsOnMasterLineage));
        }

        [Fact]
        public void Marks_divergent_sibling_states_as_conflicting_exact_evidence()
        {
            LogicalHistoryResult result = LogicalHistoryEngine.Build(
                "project",
                "mapping-1",
                new[]
                {
                    Revision("r1", "", 0, "A1", "base", baseline: true),
                    Revision("mine", "r1", 1, "A1", "mine", master: true),
                    Revision("theirs", "r1", 2, "A1", "theirs"),
                });

            LogicalSubjectHistory value = Assert.Single(result.Subjects);
            Assert.True(value.HasConflict);
            Assert.True(value.HasBranchOnlyChange);
            Assert.Equal(2, value.Transitions.Count);
            Assert.All(value.Transitions, transition => Assert.True(transition.IsConflict));
            Assert.False(value.Transitions.Single(t => t.AfterRevisionId == "mine").IsBranchOnly);
            Assert.True(value.Transitions.Single(t => t.AfterRevisionId == "theirs").IsBranchOnly);
        }

        [Fact]
        public void Groups_repeated_exact_transitions_without_losing_children()
        {
            LogicalHistoryResult result = LogicalHistoryEngine.Build(
                "project",
                "mapping-1",
                new[]
                {
                    Revision("r1", "", 0, new Dictionary<string, string>
                    {
                        ["A1"] = "old",
                        ["A2"] = "old",
                    }),
                    Revision("r2", "r1", 1, new Dictionary<string, string>
                    {
                        ["A1"] = "new",
                        ["A2"] = "new",
                    }, master: true),
                });

            HistoryChangeGroup group = Assert.Single(result.Groups);
            Assert.Equal(2, group.SubjectIds.Count);
            Assert.Equal(2, result.Subjects.Count);
            Assert.All(result.Subjects, subject =>
            {
                HistoryTransition transition = Assert.Single(subject.Transitions);
                Assert.Equal(group.GroupId, transition.GroupId);
            });
        }

        [Fact]
        public void Equivalent_labels_do_not_remove_the_exact_transition()
        {
            LogicalHistoryResult result = LogicalHistoryEngine.Build(
                "project",
                "mapping-1",
                new[]
                {
                    Revision("r1", "", 0, "A1", "ABC"),
                    Revision("r2", "r1", 1, "A1", "abc", master: true),
                });

            HistoryTransition transition = Assert.Single(Assert.Single(result.Subjects).Transitions);
            Assert.True((transition.EquivalenceTags & EquivalenceTag.CaseOnly) != 0);
            Assert.Equal(CellChangeKind.Modified, transition.ExactKind);
        }

        [Fact]
        public void Rejects_an_unconfirmed_or_invalid_revision_graph()
        {
            Assert.Throws<ArgumentException>(() => LogicalHistoryEngine.Build(
                "project",
                "mapping-1",
                new[]
                {
                    Revision("r1", "r2", 0, "A1", "one", master: true),
                    Revision("r2", "r1", 1, "A1", "two"),
                }));

            Assert.Throws<ArgumentException>(() => LogicalHistoryEngine.Build(
                "project",
                "mapping-1",
                new[]
                {
                    Revision("r1", "", 0, "A1", "one"),
                    Revision("r2", "r1", 1, "A1", "two"),
                }));
        }

        private static HistoryRevision Revision(
            string revisionId,
            string parentRevisionId,
            int order,
            string key,
            string value,
            bool master = false,
            bool baseline = false) => Revision(
                revisionId,
                parentRevisionId,
                order,
                new Dictionary<string, string> { [key] = value },
                master,
                baseline);

        private static HistoryRevision Revision(
            string revisionId,
            string parentRevisionId,
            int order,
            IDictionary<string, string> values,
            bool master = false,
            bool baseline = false)
        {
            var cells = values
                .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .Select(pair => CanonicalCellState.Constant(pair.Key, CanonicalValue.Text(pair.Value)))
                .ToList();
            return new HistoryRevision(
                revisionId,
                parentRevisionId,
                order,
                new RevisionSnapshot(revisionId, cells),
                master,
                baseline);
        }
    }
}
