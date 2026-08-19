using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.Experimental.IoBridge
{
    /// <summary>Which real document a queued action writes to.</summary>
    public enum BridgeTarget { InputsCsv, OutputsCsv, IoList }

    public enum BridgeActionKind
    {
        SetField,       // one field of an existing row/record
        CreateCsvRow,   // integrate an IO-only signal into the CSV
        DisableCsvRow,  // '#'-prefix the record
        EnableCsvRow,   // strip the '#' prefix
        DeleteCsvRow,   // remove the record
    }

    public enum BridgeActionStatus
    {
        Pending,            // waiting for apply
        Applied,            // written by this add-in
        Conflict,           // apply found the file changed since queue time — skipped
        AppliedExternally,  // the file already holds the new value (someone else did it)
        ConflictStale,      // refresh found the file changed, or the signal vanished
    }

    /// <summary>
    /// The four field pairs linked across the two sides: CSV Node/Offset/Bit ↔ IO-list "Bit" address,
    /// CSV Signal ID / Output Name ↔ IO-list Mnemonic, CSV Description ↔ IO-list description (part 1),
    /// CSV "Logic" / "NC-NO Logic" ↔ IO-list "Normal condition".
    /// </summary>
    public enum LinkedField { None, Address, Mnemonic, Description, NormalCondition }

    /// <summary>
    /// Identifies one signal independently of any sheet row: the card sheet it belongs to plus its
    /// address ("I.80.3.0", module rows "#236"), with the mnemonic as the fallback identity for rows
    /// whose address is absent or was itself corrected.
    /// </summary>
    public sealed class SignalRef
    {
        public string CardKey = "";    // "R1.S2.C1", numerically normalized
        public string Address = "";    // "I.80.3.0", "#236", or "" when unaddressed
        public string Mnemonic = "";   // raw mnemonic list at queue time ("A/B")

        public static string MakeCardKey(string rack, string slot, string card) =>
            IoListLogic.ControllerKey(rack, slot) + ".C" + IoListLogic.Norm(card);

        /// <summary>Primary identity: the address when there is one, else the mnemonic identity.</summary>
        public string IdentityKey() =>
            Address.Length > 0 ? CardKey + "|" + Address : MnemonicIdentityKey();

        /// <summary>Fallback identity from the normalized mnemonic set (BUS_/case/order-insensitive).</summary>
        public string MnemonicIdentityKey() =>
            CardKey + "~" + string.Join("/", SortedKeys(IoListLogic.MnemonicKeys(Mnemonic)));

        private static List<string> SortedKeys(List<string> keys)
        {
            keys.Sort(StringComparer.Ordinal);
            return keys;
        }
    }

    /// <summary>
    /// One queued change. <c>OldValue</c> is the target document's value as seen when the action was
    /// FIRST queued — the base of the per-action three-way check at apply time. Row-level kinds encode
    /// their state in Old/New the same way ("" → "+", "0" → "1"), decided by the single builder that
    /// creates them, so the queue and re-anchoring stay kind-agnostic.
    /// </summary>
    public sealed class BridgeAction
    {
        public string ActionId = Guid.NewGuid().ToString("N");
        public BridgeActionKind Kind = BridgeActionKind.SetField;
        public BridgeTarget Target = BridgeTarget.IoList;
        public SignalRef Signal = new SignalRef();
        public LinkedField Linked = LinkedField.None;
        public string FieldHeader = "";     // normalized header of the edited column (SetField), "" otherwise
        public string OldValue = "";
        public string NewValue = "";
        public string[] RowFields;          // CreateCsvRow payload; null otherwise
        public bool Include = true;         // unchecked on the queue sheet → excluded from apply
        public BridgeActionStatus Status = BridgeActionStatus.Pending;
        public string StatusNote = "";
        public string QueuedUtc = "";       // ISO-8601 "O", invariant
        public string AppliedUtc = "";

        /// <summary>True for statuses an apply or refresh may still act on.</summary>
        public bool IsOpen =>
            Status == BridgeActionStatus.Pending ||
            Status == BridgeActionStatus.Conflict ||
            Status == BridgeActionStatus.ConflictStale;

        /// <summary>
        /// Actions with the same key describe the same logical change and supersede each other while
        /// open. Disable/Enable share a key (they are the two directions of one toggle).
        /// </summary>
        public string SupersedeKey()
        {
            string kindGroup;
            switch (Kind)
            {
                case BridgeActionKind.SetField: kindGroup = "field:" + FieldHeader; break;
                case BridgeActionKind.DisableCsvRow:
                case BridgeActionKind.EnableCsvRow: kindGroup = "rowstate"; break;
                case BridgeActionKind.DeleteCsvRow: kindGroup = "delete"; break;
                case BridgeActionKind.CreateCsvRow: kindGroup = "create"; break;
                default: kindGroup = Kind.ToString(); break;
            }
            return Target + "|" + Signal.IdentityKey() + "|" + kindGroup;
        }
    }
}
