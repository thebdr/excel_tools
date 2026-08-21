using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    public class SwpQueueTests
    {
        private static SwpProposedEdit Edit(string proposed = "I-NEW", string current = "OLD",
            SwpField field = SwpField.AddressG, string coordinate = "DB_X.Alarm1>0>0")
        {
            return new SwpProposedEdit
            {
                Sheet = "CABINET",
                CoordinateKey = coordinate,
                Field = field,
                Mnemonic = "CABINETAL00100",
                Current = current,
                Proposed = proposed,
                RowNumber = 50,
            };
        }

        [Fact]
        public void Fresh_edits_become_pending_actions()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "2026-08-20T10:00:00Z");

            SwpAction action = Assert.Single(queue.Actions);
            Assert.Equal(SwpActionStatus.Pending, action.Status);
            Assert.Equal("OLD", action.OldValue);
            Assert.Equal("I-NEW", action.NewValue);
            Assert.True(action.Include);
            Assert.Equal("2026-08-20T10:00:00Z", action.QueuedUtc);
        }

        [Fact]
        public void Recomputed_diff_updates_the_proposal_and_keeps_identity_base_and_include()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");
            SwpAction original = queue.Actions[0];
            original.Include = false;

            queue.ReconcileFromDiff(new[] { Edit(proposed: "I-NEWER") }, a => "OLD", "t1");

            SwpAction action = Assert.Single(queue.Actions);
            Assert.Same(original, action);
            Assert.Equal("OLD", action.OldValue);       // first-seen base survives
            Assert.Equal("I-NEWER", action.NewValue);
            Assert.False(action.Include);               // the exclusion choice survives
            Assert.Equal(SwpActionStatus.Pending, action.Status);
        }

        [Fact]
        public void A_file_change_under_a_standing_diff_rebases_with_evidence()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");

            queue.ReconcileFromDiff(new[] { Edit(current: "MOVED") }, a => "MOVED", "t1");

            SwpAction action = Assert.Single(queue.Actions);
            Assert.Equal(SwpActionStatus.Pending, action.Status);
            Assert.Equal("MOVED", action.OldValue);
            Assert.Contains("OLD", action.StatusNote);   // the previous base is kept as evidence
        }

        [Fact]
        public void A_satisfied_intent_reads_applied_externally()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");

            queue.ReconcileFromDiff(new SwpProposedEdit[0], a => "I-NEW", "t1");

            SwpAction action = Assert.Single(queue.Actions);
            Assert.Equal(SwpActionStatus.AppliedExternally, action.Status);
        }

        [Fact]
        public void A_vanished_intent_is_dropped_as_moot()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");

            // The IO list changed back: the file's value is once again what the diff wants.
            queue.ReconcileFromDiff(new SwpProposedEdit[0], a => "OLD", "t1");
            Assert.Empty(queue.Actions);
        }

        [Fact]
        public void A_changed_intent_that_the_file_already_meets_is_also_moot()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");

            queue.ReconcileFromDiff(new SwpProposedEdit[0], a => "THIRD", "t1");
            Assert.Empty(queue.Actions);
        }

        [Fact]
        public void A_vanished_target_row_goes_stale()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");

            queue.ReconcileFromDiff(new SwpProposedEdit[0], a => null, "t1");

            SwpAction action = Assert.Single(queue.Actions);
            Assert.Equal(SwpActionStatus.ConflictStale, action.Status);
            Assert.Contains("not found", action.StatusNote);
        }

        [Fact]
        public void Applied_history_is_never_revisited_and_new_work_can_join_it()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit() }, a => "OLD", "t0");
            queue.Actions[0].Status = SwpActionStatus.Applied;

            queue.ReconcileFromDiff(new[] { Edit(proposed: "AGAIN", current: "I-NEW") }, a => "I-NEW", "t1");

            Assert.Equal(2, queue.Actions.Count);
            Assert.Equal(SwpActionStatus.Applied, queue.Actions[0].Status);
            Assert.Equal(SwpActionStatus.Pending, queue.Actions[1].Status);
            Assert.Equal("I-NEW", queue.Actions[1].OldValue);
        }

        [Fact]
        public void Included_pending_setinclude_and_purge_work()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[]
            {
                Edit(),
                Edit(field: SwpField.Desc1T, proposed: "D", current: "C"),
            }, a => a.Field == SwpField.AddressG ? "OLD" : "C", "t0");

            Assert.Equal(2, queue.IncludedPending().Count);
            Assert.True(queue.SetInclude(queue.Actions[0].ActionId, false));
            Assert.Single(queue.IncludedPending());
            Assert.False(queue.SetInclude("nope", false));

            queue.Actions[0].Status = SwpActionStatus.Applied;
            queue.Actions[1].Status = SwpActionStatus.AppliedExternally;
            Assert.Equal(2, queue.PurgeHistory());
            Assert.Empty(queue.Actions);
        }

        [Fact]
        public void Duplicate_edit_keys_in_one_diff_collapse_to_the_first()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[] { Edit(proposed: "A"), Edit(proposed: "B") }, a => "OLD", "t0");
            SwpAction action = Assert.Single(queue.Actions);
            Assert.Equal("A", action.NewValue);
        }

        [Fact]
        public void Distinct_fields_on_one_row_stay_distinct_actions()
        {
            var queue = new SwpQueue();
            queue.ReconcileFromDiff(new[]
            {
                Edit(),
                Edit(field: SwpField.ExportB, proposed: "x", current: "0"),
            }, a => a.Field == SwpField.AddressG ? "OLD" : "0", "t0");

            Assert.Equal(2, queue.Actions.Count);
            Assert.Equal(2, queue.Actions.Select(a => a.Key).Distinct().Count());
        }
    }
}
