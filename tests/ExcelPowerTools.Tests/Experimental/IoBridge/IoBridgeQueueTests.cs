using System.Linq;
using ExcelPowerTools.Core.Experimental.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental.IoBridge
{
    public class IoBridgeQueueTests
    {
        private static SignalRef Signal(string address = "I.80.3.0", string mnemonic = "LP100301") =>
            new SignalRef { CardKey = "R1.S1.C1", Address = address, Mnemonic = mnemonic };

        private static BridgeAction Set(string field, string oldValue, string newValue,
            BridgeTarget target = BridgeTarget.InputsCsv, SignalRef? signal = null) =>
            new BridgeAction
            {
                Kind = BridgeActionKind.SetField,
                Target = target,
                Signal = signal ?? Signal(),
                FieldHeader = field,
                OldValue = oldValue,
                NewValue = newValue,
            };

        // ---- identity ------------------------------------------------------------------------------

        [Fact]
        public void Card_key_normalizes_numerically()
        {
            Assert.Equal("R1.S2.C1", SignalRef.MakeCardKey("01", "2", "01"));
        }

        [Fact]
        public void Identity_prefers_the_address_and_falls_back_to_mnemonics()
        {
            var addressed = Signal();
            Assert.Equal("R1.S1.C1|I.80.3.0", addressed.IdentityKey());

            var unaddressed = Signal(address: "", mnemonic: "BUS_CABINETAL00161/ISPC_GROUP_1");
            Assert.Equal(unaddressed.MnemonicIdentityKey(), unaddressed.IdentityKey());

            // BUS_/case/order-insensitive, so both spellings share the fallback identity.
            var other = Signal(address: "", mnemonic: "ispc_group_1/CABINETAL00161");
            Assert.Equal(unaddressed.IdentityKey(), other.IdentityKey());
        }

        [Fact]
        public void Module_rows_are_identified_by_their_node_address()
        {
            var module = Signal(address: "#236", mnemonic: "CABINETAL00161");
            Assert.Equal("R1.S1.C1|#236", module.IdentityKey());
        }

        // ---- supersede -----------------------------------------------------------------------------

        [Fact]
        public void Repeated_edits_supersede_and_keep_the_first_old_value()
        {
            var queue = new IoBridgeQueue();
            queue.Record(Set("signalid", oldValue: "LP100301", newValue: "LP100302"));
            queue.Record(Set("signalid", oldValue: "LP100302", newValue: "LP100303"));

            BridgeAction only = Assert.Single(queue.Actions);
            Assert.Equal("LP100301", only.OldValue);   // three-way base = value at FIRST queue time
            Assert.Equal("LP100303", only.NewValue);
        }

        [Fact]
        public void Supersede_preserves_the_include_choice()
        {
            var queue = new IoBridgeQueue();
            BridgeAction first = queue.Record(Set("signalid", "A", "B"));
            queue.SetInclude(first.ActionId, false);

            queue.Record(Set("signalid", "B", "C"));
            Assert.False(Assert.Single(queue.Actions).Include);
        }

        [Fact]
        public void Different_fields_targets_and_signals_do_not_supersede()
        {
            var queue = new IoBridgeQueue();
            queue.Record(Set("signalid", "A", "B"));
            queue.Record(Set("description", "X", "Y"));
            queue.Record(Set("signalid", "A", "B", target: BridgeTarget.IoList));
            queue.Record(Set("signalid", "A", "B", signal: Signal(address: "I.80.3.1")));

            Assert.Equal(4, queue.Actions.Count);
        }

        [Fact]
        public void Disable_and_enable_are_one_toggle()
        {
            var queue = new IoBridgeQueue();
            queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.DisableCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = Signal(),
                OldValue = "0",
                NewValue = "1",
            });
            queue.Record(new BridgeAction
            {
                Kind = BridgeActionKind.EnableCsvRow,
                Target = BridgeTarget.InputsCsv,
                Signal = Signal(),
                OldValue = "1",
                NewValue = "0",
            });

            BridgeAction only = Assert.Single(queue.Actions);
            Assert.Equal(BridgeActionKind.EnableCsvRow, only.Kind);
            Assert.Equal("0", only.OldValue);   // carried from the superseded disable
        }

        [Fact]
        public void Applied_history_is_never_superseded()
        {
            var queue = new IoBridgeQueue();
            BridgeAction applied = queue.Record(Set("signalid", "A", "B"));
            applied.Status = BridgeActionStatus.Applied;

            queue.Record(Set("signalid", "B", "C"));

            Assert.Equal(2, queue.Actions.Count);
            Assert.Equal("B", queue.Actions[1].OldValue);   // fresh base, not the applied action's
        }

        // ---- unqueue / include / purge -------------------------------------------------------------

        [Fact]
        public void Typing_the_old_value_back_unqueues()
        {
            var queue = new IoBridgeQueue();
            BridgeAction a = queue.Record(Set("signalid", "A", "B"));

            Assert.True(queue.Unqueue(BridgeTarget.InputsCsv, a.Signal.IdentityKey(), "signalid"));
            Assert.Empty(queue.Actions);
            Assert.False(queue.Unqueue(BridgeTarget.InputsCsv, a.Signal.IdentityKey(), "signalid"));
        }

        [Fact]
        public void Included_pending_excludes_unchecked_and_finished_actions()
        {
            var queue = new IoBridgeQueue();
            BridgeAction keep = queue.Record(Set("signalid", "A", "B"));
            BridgeAction excluded = queue.Record(Set("description", "X", "Y"));
            BridgeAction done = queue.Record(Set("logic", "0", "1"));

            queue.SetInclude(excluded.ActionId, false);
            done.Status = BridgeActionStatus.Applied;

            Assert.Equal(new[] { keep.ActionId }, queue.IncludedPending().Select(a => a.ActionId));
        }

        [Fact]
        public void Purge_drops_only_finished_history()
        {
            var queue = new IoBridgeQueue();
            queue.Record(Set("a", "1", "2")).Status = BridgeActionStatus.Applied;
            queue.Record(Set("b", "1", "2")).Status = BridgeActionStatus.AppliedExternally;
            queue.Record(Set("c", "1", "2")).Status = BridgeActionStatus.Conflict;
            queue.Record(Set("d", "1", "2"));

            Assert.Equal(2, queue.PurgeHistory());
            Assert.Equal(2, queue.Actions.Count);
        }

        // ---- re-anchor -----------------------------------------------------------------------------

        [Fact]
        public void Reanchor_applies_the_three_rules_and_not_found()
        {
            var queue = new IoBridgeQueue();
            BridgeAction untouched = queue.Record(Set("a", "old", "new"));
            BridgeAction external = queue.Record(Set("b", "old", "new"));
            BridgeAction drifted = queue.Record(Set("c", "old", "new"));
            BridgeAction vanished = queue.Record(Set("d", "old", "new"));

            queue.Reanchor(a =>
                a.FieldHeader == "a" ? ReanchorProbe.Of("old") :
                a.FieldHeader == "b" ? ReanchorProbe.Of("new") :
                a.FieldHeader == "c" ? ReanchorProbe.Of("other") :
                ReanchorProbe.NotFound);

            Assert.Equal(BridgeActionStatus.Pending, untouched.Status);
            Assert.Equal(BridgeActionStatus.AppliedExternally, external.Status);
            Assert.Equal("observed on refresh", external.StatusNote);
            Assert.Equal(BridgeActionStatus.ConflictStale, drifted.Status);
            Assert.Contains("other", drifted.StatusNote);
            Assert.Equal(BridgeActionStatus.ConflictStale, vanished.Status);
            Assert.Contains("not found", vanished.StatusNote);
        }

        [Fact]
        public void Reanchor_heals_a_stale_conflict_whose_file_reverted()
        {
            var queue = new IoBridgeQueue();
            BridgeAction a = queue.Record(Set("a", "old", "new"));
            a.Status = BridgeActionStatus.ConflictStale;
            a.StatusNote = "file now holds \"other\"";

            queue.Reanchor(_ => ReanchorProbe.Of("old"));

            Assert.Equal(BridgeActionStatus.Pending, a.Status);
            Assert.Equal("", a.StatusNote);
        }

        [Fact]
        public void Reanchor_never_revisits_finished_history()
        {
            var queue = new IoBridgeQueue();
            BridgeAction applied = queue.Record(Set("a", "old", "new"));
            applied.Status = BridgeActionStatus.Applied;

            queue.Reanchor(_ => ReanchorProbe.Of("something else entirely"));

            Assert.Equal(BridgeActionStatus.Applied, applied.Status);
        }
    }
}
