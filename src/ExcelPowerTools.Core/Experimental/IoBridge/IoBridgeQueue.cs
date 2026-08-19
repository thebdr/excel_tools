using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.Experimental.IoBridge
{
    /// <summary>
    /// The result of resolving one action's signal against freshly rebuilt documents. When the signal
    /// was found, <c>CurrentValue</c> is the target's present value under the SAME encoding the action
    /// used for Old/New (field text for SetField; "+"/"" for existence; "1"/"0" for disabled state).
    /// </summary>
    public sealed class ReanchorProbe
    {
        public bool Resolved;
        public string CurrentValue = "";

        public static readonly ReanchorProbe NotFound = new ReanchorProbe { Resolved = false };
        public static ReanchorProbe Of(string currentValue) =>
            new ReanchorProbe { Resolved = true, CurrentValue = currentValue ?? "" };
    }

    /// <summary>
    /// The pending-change queue. Pure semantics only — rendering, persistence and application live
    /// elsewhere. Order is queue order; applied/conflicted rows remain as history.
    /// </summary>
    public sealed class IoBridgeQueue
    {
        public List<BridgeAction> Actions = new List<BridgeAction>();

        /// <summary>
        /// Adds an action, superseding an OPEN action with the same <see cref="BridgeAction.SupersedeKey"/>.
        /// The superseded action's ORIGINAL OldValue is carried forward: the three-way check at apply
        /// time must compare against the file value seen when the change was first queued, never against
        /// an intermediate typed value. Returns the stored action.
        /// </summary>
        public BridgeAction Record(BridgeAction action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));

            string key = action.SupersedeKey();
            for (int i = 0; i < Actions.Count; i++)
            {
                BridgeAction existing = Actions[i];
                if (!existing.IsOpen || existing.SupersedeKey() != key) continue;

                action.OldValue = existing.OldValue;
                action.Include = existing.Include;
                Actions[i] = action;
                return action;
            }

            Actions.Add(action);
            return action;
        }

        /// <summary>
        /// Removes the open action matching the key — the user typed the original value back, so there
        /// is nothing left to apply. True when something was removed.
        /// </summary>
        public bool Unqueue(BridgeTarget target, string identityKey, string fieldHeader)
        {
            for (int i = 0; i < Actions.Count; i++)
            {
                BridgeAction a = Actions[i];
                if (!a.IsOpen || a.Kind != BridgeActionKind.SetField) continue;
                if (a.Target != target || a.FieldHeader != fieldHeader) continue;
                if (a.Signal.IdentityKey() != identityKey) continue;

                Actions.RemoveAt(i);
                return true;
            }
            return false;
        }

        public List<BridgeAction> IncludedPending()
        {
            var result = new List<BridgeAction>();
            foreach (BridgeAction a in Actions)
                if (a.Include && a.Status == BridgeActionStatus.Pending) result.Add(a);
            return result;
        }

        public bool SetInclude(string actionId, bool include)
        {
            foreach (BridgeAction a in Actions)
            {
                if (a.ActionId != actionId) continue;
                a.Include = include;
                return true;
            }
            return false;
        }

        /// <summary>Drops finished history (Applied / AppliedExternally). Returns how many were removed.</summary>
        public int PurgeHistory()
        {
            int removed = Actions.RemoveAll(a =>
                a.Status == BridgeActionStatus.Applied ||
                a.Status == BridgeActionStatus.AppliedExternally);
            return removed;
        }

        /// <summary>
        /// Re-anchors every OPEN action against freshly loaded documents (the workspace was rebuilt).
        /// Confirmed rules: current==Old → Pending (a stale conflict whose file reverted heals);
        /// current==New → AppliedExternally; anything else → ConflictStale carrying what the file now
        /// holds; an unresolvable signal → ConflictStale ("signal not found"). Final statuses
        /// (Applied / AppliedExternally) are never revisited.
        /// </summary>
        public void Reanchor(Func<BridgeAction, ReanchorProbe> resolve)
        {
            if (resolve == null) throw new ArgumentNullException(nameof(resolve));

            foreach (BridgeAction a in Actions)
            {
                if (!a.IsOpen) continue;

                ReanchorProbe probe = resolve(a) ?? ReanchorProbe.NotFound;
                if (!probe.Resolved)
                {
                    a.Status = BridgeActionStatus.ConflictStale;
                    a.StatusNote = "signal not found on refresh";
                }
                else if (probe.CurrentValue == a.OldValue)
                {
                    a.Status = BridgeActionStatus.Pending;
                    a.StatusNote = "";
                }
                else if (probe.CurrentValue == a.NewValue)
                {
                    a.Status = BridgeActionStatus.AppliedExternally;
                    a.StatusNote = "observed on refresh";
                }
                else
                {
                    a.Status = BridgeActionStatus.ConflictStale;
                    a.StatusNote = "file now holds \"" + probe.CurrentValue + "\"";
                }
            }
        }
    }
}
