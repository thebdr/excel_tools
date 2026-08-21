using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    public enum SwpActionStatus
    {
        Pending,            // waiting for apply
        Applied,            // written by this add-in
        Conflict,           // apply found the file changed since the diff — skipped
        AppliedExternally,  // the file already holds the intended value
        ConflictStale,      // the target row vanished on refresh
    }

    /// <summary>
    /// One pending SWP cell write, produced by the diff (never typed): the coordinate names the
    /// target row, <c>OldValue</c> is the file's value when the change was FIRST proposed — the
    /// base of the per-action three-way check at apply time.
    /// </summary>
    public sealed class SwpAction
    {
        public string ActionId = Guid.NewGuid().ToString("N");
        public string Sheet = "";
        public string CoordinateKey = "";   // "DB_CABINETS_DIAG.Alarm1>3>17"
        public SwpField Field;
        public string Mnemonic = "";        // provenance, links the action back to its spine row
        public string OldValue = "";
        public string NewValue = "";
        public bool Include = true;         // unchecked on the queue sheet → excluded from apply
        public SwpActionStatus Status = SwpActionStatus.Pending;
        public string StatusNote = "";
        public string QueuedUtc = "";       // ISO-8601 "O", invariant
        public string AppliedUtc = "";

        public bool IsOpen =>
            Status == SwpActionStatus.Pending ||
            Status == SwpActionStatus.Conflict ||
            Status == SwpActionStatus.ConflictStale;

        /// <summary>One logical change per target cell.</summary>
        public string Key => Sheet + "|" + CoordinateKey + "|" + Field;
    }

    /// <summary>One cell difference the diff wants applied: current file value → intended value.</summary>
    public sealed class SwpProposedEdit
    {
        public string Sheet = "";
        public string CoordinateKey = "";
        public SwpField Field;
        public string Mnemonic = "";
        public string Current = "";
        public string Proposed = "";
        public int RowNumber;              // SWP worksheet row, display/preview only

        public string Key => Sheet + "|" + CoordinateKey + "|" + Field;
    }

    /// <summary>
    /// The SWP action queue. Unlike the io-bridge queue (fed by typed edits), this one is
    /// DIFF-DRIVEN: every build recomputes the intended edits and <see cref="ReconcileFromDiff"/>
    /// folds them into the standing actions — include choices and the first-seen old value
    /// survive, satisfied intents retire, and vanished targets go stale. Applied rows remain as
    /// history until purged.
    /// </summary>
    public sealed class SwpQueue
    {
        public List<SwpAction> Actions = new List<SwpAction>();

        public List<SwpAction> IncludedPending()
        {
            var result = new List<SwpAction>();
            foreach (SwpAction a in Actions)
                if (a.Include && a.Status == SwpActionStatus.Pending) result.Add(a);
            return result;
        }

        public bool SetInclude(string actionId, bool include)
        {
            foreach (SwpAction a in Actions)
            {
                if (a.ActionId != actionId) continue;
                a.Include = include;
                return true;
            }
            return false;
        }

        /// <summary>Drops finished history (Applied / AppliedExternally). Returns how many went.</summary>
        public int PurgeHistory()
        {
            return Actions.RemoveAll(a =>
                a.Status == SwpActionStatus.Applied ||
                a.Status == SwpActionStatus.AppliedExternally);
        }

        /// <summary>
        /// Folds a freshly computed diff into the standing queue. Per OPEN action:
        /// target row gone → ConflictStale; still diffing with the same base → Pending with the
        /// fresh proposal (first-seen OldValue and the Include choice survive); still diffing on a
        /// CHANGED base → Pending re-based, the note recording the old base; no longer diffing →
        /// AppliedExternally when the file holds the action's intent, otherwise the intent is moot
        /// and the action is dropped. Diffs with no standing action become new Pending actions.
        /// <paramref name="resolveCurrent"/> returns the file's current value for an action's
        /// target cell, or null when the row no longer exists.
        /// </summary>
        public void ReconcileFromDiff(IReadOnlyList<SwpProposedEdit> edits,
            Func<SwpAction, string> resolveCurrent, string nowUtc)
        {
            if (edits == null) throw new ArgumentNullException(nameof(edits));
            if (resolveCurrent == null) throw new ArgumentNullException(nameof(resolveCurrent));

            var editsByKey = new Dictionary<string, SwpProposedEdit>();
            foreach (SwpProposedEdit edit in edits)
                if (!editsByKey.ContainsKey(edit.Key)) editsByKey[edit.Key] = edit;

            var covered = new HashSet<string>();
            var moot = new List<SwpAction>();

            foreach (SwpAction action in Actions)
            {
                if (!action.IsOpen) continue;

                string current = resolveCurrent(action);
                if (current == null)
                {
                    action.Status = SwpActionStatus.ConflictStale;
                    action.StatusNote = "target row not found on refresh";
                    covered.Add(action.Key);   // suppress a twin action on an unresolvable target
                    continue;
                }

                if (editsByKey.TryGetValue(action.Key, out SwpProposedEdit edit))
                {
                    covered.Add(action.Key);
                    if (current == action.OldValue)
                    {
                        action.Status = SwpActionStatus.Pending;
                        action.StatusNote = "";
                        action.NewValue = edit.Proposed;
                    }
                    else
                    {
                        // The file moved underneath the intent; the diff already shows the new
                        // reality, so the action re-bases on it — keeping the evidence in the note.
                        action.Status = SwpActionStatus.Pending;
                        action.StatusNote = "re-based; the file held \"" + action.OldValue + "\" at first diff";
                        action.OldValue = current;
                        action.NewValue = edit.Proposed;
                    }
                }
                else if (current == action.NewValue)
                {
                    action.Status = SwpActionStatus.AppliedExternally;
                    action.StatusNote = "observed on refresh";
                }
                else
                {
                    moot.Add(action);   // the intent went away (IO list caught up or changed)
                }
            }

            foreach (SwpAction gone in moot) Actions.Remove(gone);

            foreach (SwpProposedEdit edit in edits)
            {
                if (covered.Contains(edit.Key)) continue;
                covered.Add(edit.Key);
                Actions.Add(new SwpAction
                {
                    Sheet = edit.Sheet,
                    CoordinateKey = edit.CoordinateKey,
                    Field = edit.Field,
                    Mnemonic = edit.Mnemonic,
                    OldValue = edit.Current,
                    NewValue = edit.Proposed,
                    QueuedUtc = nowUtc ?? "",
                });
            }
        }
    }
}
