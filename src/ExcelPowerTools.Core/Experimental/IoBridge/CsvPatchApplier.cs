using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Core.Experimental.IoBridge
{
    /// <summary>Resolved field positions of one plant CSV, tolerant of both export schemas.</summary>
    public sealed class CsvColumnsMap
    {
        public int HeaderIndex = -1;
        public int Name = -1, Node = -1, Offset = -1, Bit = -1, Board = -1, Type = -1;
        public IReadOnlyList<string> Headers;

        private Dictionary<string, int> _byNormalizedHeader;

        public static readonly string[] NameHeaders = { "signal id", "input name", "output name", "name" };
        public static readonly string[] BoardHeaders = { "card", "board number", "board" };
        public static readonly string[] TypeHeaders = { "type", "input type", "output type" };

        /// <summary>Locates the header record and key columns; null when this is not a plant CSV.</summary>
        public static CsvColumnsMap Locate(CsvDocument doc)
        {
            if (doc == null) return null;
            var rows = new List<string[]>(doc.Records.Count);
            foreach (CsvRecord rec in doc.Records) rows.Add(rec.Fields);

            int headerIdx = CsvDetect.DetectHeaderRow(rows, "#");
            if (headerIdx < 0) return null;

            string[] headers = rows[headerIdx];
            var map = new CsvColumnsMap
            {
                HeaderIndex = headerIdx,
                Headers = headers,
                Name = IoListLogic.FindColumn(headers, NameHeaders),
                Node = IoListLogic.FindColumn(headers, "node"),
                Offset = IoListLogic.FindColumn(headers, "offset"),
                Bit = IoListLogic.FindColumn(headers, "bit"),
                Board = IoListLogic.FindColumn(headers, BoardHeaders),
                Type = IoListLogic.FindColumn(headers, TypeHeaders),
            };
            if (map.Name < 0 || map.Node < 0 || map.Offset < 0 || map.Bit < 0 || map.Board < 0) return null;

            map._byNormalizedHeader = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < headers.Length; i++)
            {
                string key = IoListLogic.NormalizeHeader(headers[i]);
                if (key.Length > 0 && !map._byNormalizedHeader.ContainsKey(key))
                    map._byNormalizedHeader[key] = i;
            }
            return map;
        }

        /// <summary>Field index for a queued action's normalized header; -1 when this file lacks it.</summary>
        public int FieldFor(string normalizedHeader) =>
            normalizedHeader != null && _byNormalizedHeader != null &&
            _byNormalizedHeader.TryGetValue(normalizedHeader, out int index) ? index : -1;
    }

    public sealed class CsvActionOutcome
    {
        public BridgeAction Action;
        public BridgeActionStatus Outcome;   // Applied here means "will be applied once the file is written"
        public string Note = "";
    }

    public sealed class CsvPatchResult
    {
        public bool AnyMutation;
        public List<CsvActionOutcome> Outcomes = new List<CsvActionOutcome>();
    }

    /// <summary>
    /// Applies queued actions to a loaded <see cref="CsvDocument"/> with the per-action three-way
    /// check. This mutates ONLY the in-memory document and reports intended outcomes — the caller
    /// verifies the round trip, writes the file, and only then flips action statuses. Field-0 markers:
    /// "@" active, "#" disabled, "!" include directive (never touched).
    /// </summary>
    public static class CsvPatchApplier
    {
        private const string ActiveMarker = "@";
        private const string DisabledMarker = "#";

        public static CsvPatchResult Apply(CsvDocument doc, IReadOnlyList<BridgeAction> actions, CsvColumnsMap map)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (map == null) throw new ArgumentNullException(nameof(map));

            var result = new CsvPatchResult();
            if (actions == null || actions.Count == 0) return result;

            var index = new CsvRecordResolver(doc, map);
            foreach (BridgeAction action in actions)
                result.Outcomes.Add(ApplyOne(doc, map, index, action, result));
            return result;
        }

        private static CsvActionOutcome ApplyOne(
            CsvDocument doc, CsvColumnsMap map, CsvRecordResolver index, BridgeAction action, CsvPatchResult result)
        {
            var outcome = new CsvActionOutcome { Action = action };

            int recordIdx = index.Resolve(action.Signal, out string resolveNote);
            if (recordIdx == CsvRecordResolver.Ambiguous)
            {
                outcome.Outcome = BridgeActionStatus.Conflict;
                outcome.Note = "ambiguous: " + resolveNote;
                return outcome;
            }

            switch (action.Kind)
            {
                case BridgeActionKind.SetField:
                {
                    if (recordIdx < 0) return NotFound(outcome);

                    int field = map.FieldFor(action.FieldHeader);
                    if (field < 0)
                    {
                        outcome.Outcome = BridgeActionStatus.Conflict;
                        outcome.Note = "column \"" + action.FieldHeader + "\" not found in the file";
                        return outcome;
                    }

                    CsvRecord record = doc.Records[recordIdx];
                    string current = field < record.Fields.Length ? record.Fields[field] : "";

                    if (current == action.NewValue) return External(outcome);
                    if (current != action.OldValue) return Drifted(outcome, current);

                    SetField(doc, recordIdx, field, action.NewValue);
                    result.AnyMutation = true;
                    outcome.Outcome = BridgeActionStatus.Applied;
                    return outcome;
                }

                case BridgeActionKind.CreateCsvRow:
                {
                    if (recordIdx >= 0)
                    {
                        outcome.Outcome = BridgeActionStatus.AppliedExternally;
                        outcome.Note = "row already present";
                        return outcome;
                    }
                    if (action.RowFields == null || action.RowFields.Length == 0)
                    {
                        outcome.Outcome = BridgeActionStatus.Conflict;
                        outcome.Note = "the action carries no row payload";
                        return outcome;
                    }
                    // A record without a signal name is garbage in the plant CSV — refuse it here so
                    // NO caller can ever append one, whatever went wrong upstream.
                    if (map.Name >= action.RowFields.Length ||
                        action.RowFields[map.Name] == null || action.RowFields[map.Name].Trim().Length == 0)
                    {
                        outcome.Outcome = BridgeActionStatus.Conflict;
                        outcome.Note = "the payload carries no signal name — refused";
                        return outcome;
                    }

                    doc.Records.Add(CsvRecord.Data((string[])action.RowFields.Clone()));
                    result.AnyMutation = true;
                    outcome.Outcome = BridgeActionStatus.Applied;
                    return outcome;
                }

                case BridgeActionKind.DisableCsvRow:
                case BridgeActionKind.EnableCsvRow:
                {
                    if (recordIdx < 0) return NotFound(outcome);

                    CsvRecord record = doc.Records[recordIdx];
                    bool disabled = Marker(record).StartsWith(DisabledMarker, StringComparison.Ordinal);
                    bool wantDisabled = action.Kind == BridgeActionKind.DisableCsvRow;

                    if (disabled == wantDisabled)
                    {
                        outcome.Outcome = BridgeActionStatus.AppliedExternally;
                        outcome.Note = wantDisabled ? "already disabled" : "already enabled";
                        return outcome;
                    }

                    SetField(doc, recordIdx, 0, wantDisabled ? DisabledMarker : ActiveMarker);
                    result.AnyMutation = true;
                    outcome.Outcome = BridgeActionStatus.Applied;
                    return outcome;
                }

                case BridgeActionKind.DeleteCsvRow:
                {
                    if (recordIdx < 0)
                    {
                        outcome.Outcome = BridgeActionStatus.AppliedExternally;
                        outcome.Note = "already gone";
                        return outcome;
                    }

                    doc.Records.RemoveAt(recordIdx);
                    index.OnRecordRemoved(recordIdx);
                    result.AnyMutation = true;
                    outcome.Outcome = BridgeActionStatus.Applied;
                    return outcome;
                }

                default:
                    outcome.Outcome = BridgeActionStatus.Conflict;
                    outcome.Note = "unsupported action kind";
                    return outcome;
            }
        }

        private static CsvActionOutcome NotFound(CsvActionOutcome outcome)
        {
            outcome.Outcome = BridgeActionStatus.Conflict;
            outcome.Note = "signal not found in the file";
            return outcome;
        }

        private static CsvActionOutcome External(CsvActionOutcome outcome)
        {
            outcome.Outcome = BridgeActionStatus.AppliedExternally;
            outcome.Note = "the file already holds the new value";
            return outcome;
        }

        private static CsvActionOutcome Drifted(CsvActionOutcome outcome, string current)
        {
            outcome.Outcome = BridgeActionStatus.Conflict;
            outcome.Note = "file now holds \"" + current + "\"";
            return outcome;
        }

        private static string Marker(CsvRecord record) =>
            record.Fields.Length > 0 && record.Fields[0] != null ? record.Fields[0] : "";

        /// <summary>Sets one field, widening the record when the index lies beyond its arity — the
        /// widened record keeps its exact original terminator.</summary>
        private static void SetField(CsvDocument doc, int recordIdx, int field, string value)
        {
            CsvRecord record = doc.Records[recordIdx];
            if (field < record.Fields.Length)
            {
                record.Fields[field] = value;
                return;
            }

            var wider = new string[field + 1];
            for (int i = 0; i < wider.Length; i++)
                wider[i] = i < record.Fields.Length ? record.Fields[i] : "";
            wider[field] = value;

            CsvRecord replacement = CsvRecord.Data(wider);
            replacement.Terminator = record.Terminator;
            doc.Records[recordIdx] = replacement;
        }

    }

    /// <summary>
    /// Signal → record resolution. Address identity first (node.offset.board.bit, or board#node for
    /// module/"#" signals), then the mnemonic identity; more than one candidate is an error, not a
    /// guess. Disabled rows resolve like active ones (Enable must find them); "!" include rows and
    /// the header never do. Public so refresh re-anchoring resolves against fresh documents with the
    /// exact semantics the apply engine uses.
    /// </summary>
    public sealed class CsvRecordResolver
    {
            public const int NotFound = -1;
            public const int Ambiguous = -2;

            private readonly Dictionary<string, List<int>> _byAddress = new Dictionary<string, List<int>>();
            private readonly Dictionary<string, List<int>> _byNode = new Dictionary<string, List<int>>();
            private readonly Dictionary<string, List<int>> _byMnemonic = new Dictionary<string, List<int>>();

            public CsvRecordResolver(CsvDocument doc, CsvColumnsMap map)
            {
                for (int i = map.HeaderIndex + 1; i < doc.Records.Count; i++)
                {
                    string[] f = doc.Records[i].Fields;
                    if (f.Length == 0 || f[0] == null || f[0].StartsWith("!", StringComparison.Ordinal)) continue;

                    string node = At(f, map.Node), board = At(f, map.Board);
                    if (node.Trim().Length == 0 || board.Trim().Length == 0) continue;

                    string type = At(f, map.Type);
                    if (IoListLogic.IsBusDiagnosticType(type))
                    {
                        Add(_byNode, IoListLogic.NodeKey(board, node), i);
                    }
                    else
                    {
                        string off = At(f, map.Offset), bit = At(f, map.Bit);
                        if (off.Trim().Length > 0 && bit.Trim().Length > 0)
                            Add(_byAddress, AddressKey(node, off, board, bit), i);
                    }

                    foreach (string key in IoListLogic.MnemonicKeys(At(f, map.Name)))
                        Add(_byMnemonic, key, i);
                }
            }

            public int Resolve(SignalRef signal, out string note)
            {
                note = "";
                if (signal == null) return NotFound;

                string card = CardOf(signal.CardKey);

                if (signal.Address.StartsWith("#", StringComparison.Ordinal))
                    return Single(_byNode, IoListLogic.NodeKey(card, signal.Address.Substring(1)), out note);

                if (IoListLogic.TryParseAddress(signal.Address, out _, out string node, out string off, out string bit))
                {
                    int found = Single(_byAddress, AddressKey(node, off, card, bit), out note);
                    if (found != NotFound) return found;
                }

                // Fallback: the mnemonic — the signal may have been re-addressed in the file.
                var hits = new List<int>();
                foreach (string key in IoListLogic.MnemonicKeys(signal.Mnemonic))
                    if (_byMnemonic.TryGetValue(key, out List<int> list))
                        foreach (int idx in list)
                            if (!hits.Contains(idx)) hits.Add(idx);

                if (hits.Count == 1) return hits[0];
                if (hits.Count > 1)
                {
                    note = "the mnemonic matches " + hits.Count + " rows";
                    return Ambiguous;
                }
                return NotFound;
            }

            /// <summary>Deleting a record shifts everything after it; kept indices must follow.</summary>
            public void OnRecordRemoved(int removedIdx)
            {
                Shift(_byAddress, removedIdx);
                Shift(_byNode, removedIdx);
                Shift(_byMnemonic, removedIdx);
            }

            private static void Shift(Dictionary<string, List<int>> map, int removedIdx)
            {
                foreach (List<int> list in map.Values)
                    for (int i = list.Count - 1; i >= 0; i--)
                    {
                        if (list[i] == removedIdx) list.RemoveAt(i);
                        else if (list[i] > removedIdx) list[i]--;
                    }
            }

            private static int Single(Dictionary<string, List<int>> map, string key, out string note)
            {
                note = "";
                if (!map.TryGetValue(key, out List<int> list) || list.Count == 0) return NotFound;
                if (list.Count > 1)
                {
                    note = "the address matches " + list.Count + " rows";
                    return Ambiguous;
                }
                return list[0];
            }

            private static string AddressKey(string node, string off, string board, string bit) =>
                IoListLogic.Norm(node) + "|" + IoListLogic.Norm(off) + "|" +
                IoListLogic.Norm(board) + "|" + IoListLogic.Norm(bit);

            private static string CardOf(string cardKey)
            {
                int c = cardKey != null ? cardKey.LastIndexOf(".C", StringComparison.Ordinal) : -1;
                return c >= 0 ? cardKey.Substring(c + 2) : "";
            }

            private static string At(string[] fields, int index) =>
                index >= 0 && index < fields.Length && fields[index] != null ? fields[index] : "";

            private static void Add(Dictionary<string, List<int>> map, string key, int index)
            {
                if (!map.TryGetValue(key, out List<int> list)) map[key] = list = new List<int>();
                list.Add(index);
            }
    }
}
