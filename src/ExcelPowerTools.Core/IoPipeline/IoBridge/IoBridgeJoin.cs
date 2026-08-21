using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Core.IoPipeline.IoBridge
{
    /// <summary>One card sheet as read from the IO-list workbook: raw rows, header not yet located.</summary>
    public sealed class CardSheetGrid
    {
        public string SheetName = "";
        public List<string[]> AllRows = new List<string[]>();
        public int FirstRowNumber = 1;   // worksheet row of AllRows[0]

        /// <summary>Source column widths (1-based column − 1), captured at read time; may be empty.</summary>
        public List<double> ColumnWidths = new List<double>();
    }

    public enum BridgeRowKind { Signal, Module, CsvOnly }

    /// <summary>
    /// Where a bridge row's CSV side lives. One IO point may be represented by SEVERAL CSV records at
    /// the same address (paired signals sharing a physical point) — they are all claimed here.
    /// <see cref="RecordIndex"/>/<see cref="Fields"/> describe the FIRST record for single-record
    /// consumers; <see cref="RecordIndexes"/> carries them all.
    /// </summary>
    public sealed class CsvBinding
    {
        public BridgeTarget File;
        public int RecordIndex;
        public string[] Fields;
        public bool Disabled;   // '#' marker (first record)

        /// <summary>The record is a VIRTUAL twin of the IO point (type 5/6/9/10 carrying the same
        /// mnemonic): it looks like the signal is there, but it acquires nothing from the field.
        /// Bound so the row shows it — and offers "change type" — instead of integrating a double.</summary>
        public bool Virtual;

        public List<int> RecordIndexes = new List<int>();
        public List<string[]> AllFields = new List<string[]>();

        public bool IsMulti => RecordIndexes.Count > 1;
    }

    public sealed class BridgeRow
    {
        public BridgeRowKind Kind;
        public SignalRef Signal = new SignalRef();
        public int IoSourceRow;          // 1-based worksheet row in the ORIGINAL card sheet; 0 for CSV-only
        public string RawBitText = "";   // the Bit cell exactly as read — the later WorksheetFunction.Match key
        public string[] IoValues = Empty;
        public string[] CsvValues = Empty;   // canonical CSV block (CsvCanonical order)
        public CsvBinding Csv;
        public string MatchVerdict = "";     // "" agree · "n/m" ratio · "missing" spine row without a CSV row
        public string RowLabel = "";         // "Signal" / "Module" / "CSV-only" (+ " · disabled" / " · virtual")

        /// <summary>The parts of the IO-list mnemonic cell the exclude regex filtered OUT of the
        /// comparison ("/"-joined, original spelling) — generator keys riding along a real mnemonic
        /// ("BUS_CABINETAL00159/SAFETY_NUMDOORS_16"). Kept so an IO-side write puts them back.</summary>
        public string ExcludedMnemonics = "";

        private static readonly string[] Empty = new string[0];
    }

    public sealed class BridgeSheetModel
    {
        public string SheetName = "";
        public string CardKey = "";
        public string Card = "";
        public string[] IoHeaders = new string[0];
        public int IoHeaderSourceRow;    // 1-based worksheet row of the header on the original sheet
        public int BitColumn = -1;       // into IoHeaders — where addresses live on the original sheet
        public List<double> SourceColumnWidths = new List<double>();   // per IO column, from the source sheet
        public List<BridgeRow> SpineRows = new List<BridgeRow>();
        public List<BridgeRow> CsvOnlyRows = new List<BridgeRow>();
    }

    public sealed class BridgeWorkspaceModel
    {
        public List<BridgeSheetModel> Sheets = new List<BridgeSheetModel>();
        public CsvColumnsMap InputsMap, OutputsMap;

        /// <summary>Per-file REAL normalized headers behind each canonical CSV column (queued actions
        /// carry the file's own header, not the canonical label). Empty string when the file lacks it.</summary>
        public string[] InputsFieldHeaders = new string[0], OutputsFieldHeaders = new string[0];

        /// <summary>Identity → row, across every sheet — the re-anchor and edit-capture lookup.</summary>
        public Dictionary<string, BridgeRow> SignalIndex =
            new Dictionary<string, BridgeRow>(StringComparer.Ordinal);

        public BridgeRow FindByIdentity(SignalRef signal)
        {
            if (signal == null) return null;
            if (SignalIndex.TryGetValue(signal.IdentityKey(), out BridgeRow row)) return row;
            return SignalIndex.TryGetValue(signal.MnemonicIdentityKey(), out row) ? row : null;
        }
    }

    /// <summary>
    /// Composes the plant-CSV description the way the export tool writes it:
    /// "DESC1A DESC1B +LOCATION-DEVICE [DRW]" — e.g.
    /// "CIRCUIT BREAKER TRIPPED POWER SUPPLY COOLING FAN +MS1.CC1-F22001 [Q0001]".
    /// Location and Device concatenate directly; the bracket holds the drawing name with its project
    /// prefix stripped ("8FVZ_Q0001" → "Q0001"). Empty pieces drop out cleanly.
    /// </summary>
    public static class CsvDescription
    {
        public static string Compose(string desc1a, string desc1b, string location, string device, string drawingName)
        {
            var pieces = new List<string>();

            string text = JoinWords(desc1a, desc1b);
            if (text.Length > 0) pieces.Add(text);

            string locDev = (location ?? "").Trim() + (device ?? "").Trim();
            if (locDev.Length > 0) pieces.Add(locDev);

            string drawing = ShortDrawing(drawingName);
            if (drawing.Length > 0) pieces.Add("[" + drawing + "]");

            return string.Join(" ", pieces);
        }

        /// <summary>The drawing code the CSV brackets: everything after the LAST underscore.</summary>
        public static string ShortDrawing(string drawingName)
        {
            string name = (drawingName ?? "").Trim();
            if (name.Length == 0) return "";
            int underscore = name.LastIndexOf('_');
            return underscore >= 0 && underscore < name.Length - 1 ? name.Substring(underscore + 1) : name;
        }

        private static string JoinWords(string a, string b)
        {
            a = (a ?? "").Trim();
            b = (b ?? "").Trim();
            return a.Length == 0 ? b : b.Length == 0 ? a : a + " " + b;
        }
    }

    /// <summary>The unified CSV display block: one column set covering both export schemas.</summary>
    public static class CsvCanonical
    {
        public static readonly string[] Labels =
            { "Name", "Type", "Logic", "Node", "Offset", "Bit", "Card", "Description" };

        public const int Name = 0, Type = 1, Logic = 2, Node = 3, Offset = 4, Bit = 5, Card = 6, Description = 7;

        /// <summary>Layout keys for the CSV block ("csv.name", …) — distinct from any real header.</summary>
        public static string LayoutKey(int canonicalIndex) =>
            "csv." + Labels[canonicalIndex].ToLowerInvariant();

        /// <summary>The file's field index behind a canonical column; -1 when absent.</summary>
        public static int FieldIndex(CsvColumnsMap map, int canonicalIndex)
        {
            if (map == null) return -1;
            switch (canonicalIndex)
            {
                case Name: return map.Name;
                case Type: return map.Type;
                case Logic: return map.FieldFor("logic") >= 0 ? map.FieldFor("logic") : map.FieldFor("ncnologic");
                case Node: return map.Node;
                case Offset: return map.Offset;
                case Bit: return map.Bit;
                case Card: return map.Board;
                case Description: return map.FieldFor("description");
                default: return -1;
            }
        }

        /// <summary>The file's REAL normalized headers per canonical column ("" when absent).</summary>
        public static string[] FieldHeaders(CsvColumnsMap map)
        {
            var result = new string[Labels.Length];
            for (int i = 0; i < Labels.Length; i++)
            {
                int field = FieldIndex(map, i);
                result[i] = field >= 0 && map != null && field < map.Headers.Count
                    ? IoListLogic.NormalizeHeader(map.Headers[field])
                    : "";
            }
            return result;
        }
    }

    /// <summary>Per-session join tuning, surfaced on the setup form.</summary>
    public sealed class IoBridgeJoinOptions
    {
        /// <summary>
        /// Regex over CSV mnemonics: matching VIRTUAL rows (types 5/6/9/10) join the comparison BY
        /// NAME against IO-list rows carrying that mnemonic — for virtuals that mirror real IO-list
        /// entries (the CLASSDIAGAL pattern). Empty = virtuals stay out, as always.
        /// </summary>
        public string IncludeVirtualsMatching = "";

        /// <summary>
        /// Regex over IO-LIST mnemonics: matching spine rows are shown but IGNORED by the comparison —
        /// they claim no CSV record and carry no verdict.
        /// </summary>
        public string ExcludeIoMnemonicsMatching = "";
    }

    /// <summary>
    /// Builds the workspace model: for every selected card sheet, the IO-list rows in their own order
    /// (the spine) joined to the CSV record at the same physical point, plus the CSV records of that
    /// card no spine row claimed. Same signal taxonomy as the pull commands: virtual (types 5/6/9/10)
    /// CSV rows never join by address, bus-diagnostic (type 11) rows join by node against module rows.
    ///
    /// Claiming is NAME-AWARE in two passes: when several IO rows share one address (a struct binding
    /// point above the actual DO — the SIREN pattern), each CSV record first goes to the IO row whose
    /// mnemonic carries its name; only records no row claims by name fall back to positional claiming.
    /// </summary>
    public static class IoBridgeJoin
    {
        /// <summary>First row (within 20) carrying both a "Bit" and a "Mnemonic" column.</summary>
        public static int FindHeaderRow(IReadOnlyList<string[]> rows)
        {
            if (rows == null) return -1;
            int limit = Math.Min(rows.Count, 20);
            for (int r = 0; r < limit; r++)
            {
                if (rows[r] == null) continue;
                if (IoListLogic.FindColumn(rows[r], "bit") >= 0 &&
                    IoListLogic.FindColumn(rows[r], "mnemonic") >= 0) return r;
            }
            return -1;
        }

        public static BridgeWorkspaceModel Build(
            IReadOnlyList<CardSheetGrid> cards, CsvDocument inputs, CsvDocument outputs,
            IoBridgeJoinOptions options = null)
        {
            options = options ?? new IoBridgeJoinOptions();
            Regex includeVirtuals = CompileOrNull(options.IncludeVirtualsMatching);
            Regex excludeIo = CompileOrNull(options.ExcludeIoMnemonicsMatching);

            var model = new BridgeWorkspaceModel
            {
                InputsMap = inputs != null ? CsvColumnsMap.Locate(inputs) : null,
                OutputsMap = outputs != null ? CsvColumnsMap.Locate(outputs) : null,
            };
            model.InputsFieldHeaders = CsvCanonical.FieldHeaders(model.InputsMap);
            model.OutputsFieldHeaders = CsvCanonical.FieldHeaders(model.OutputsMap);

            var inputsIndex = new CsvSideIndex(inputs, model.InputsMap, BridgeTarget.InputsCsv, 'I', includeVirtuals);
            var outputsIndex = new CsvSideIndex(outputs, model.OutputsMap, BridgeTarget.OutputsCsv, 'O', includeVirtuals);

            foreach (CardSheetGrid card in cards ?? new List<CardSheetGrid>())
            {
                BridgeSheetModel sheet = BuildSheet(card, model, inputsIndex, outputsIndex, excludeIo);
                if (sheet != null) model.Sheets.Add(sheet);
            }

            // CSV records of a selected card that no spine row claimed — the CSV-only bucket.
            CollectCsvOnly(model, inputsIndex, BridgeTarget.InputsCsv);
            CollectCsvOnly(model, outputsIndex, BridgeTarget.OutputsCsv);
            return model;
        }

        private static Regex CompileOrNull(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern)) return null;
            try { return new Regex(pattern.Trim(), RegexOptions.IgnoreCase); }
            catch (ArgumentException) { return null; }   // an invalid pattern matches nothing
        }

        /// <summary>One spine row in the making, held while the claim passes run.</summary>
        private sealed class PendingEntry
        {
            public BridgeRow Row;
            public CsvSideIndex Side;
            public string ClaimKey;          // the address/node key its records live under
            public List<string> NameKeys;    // the row's mnemonic keys
            public bool Excluded;
            public bool VirtualTwin;         // the claim is a virtual record carrying the row's name
            public List<int> Claimed = new List<int>();
        }

        private static BridgeSheetModel BuildSheet(
            CardSheetGrid card, BridgeWorkspaceModel model, CsvSideIndex inputs, CsvSideIndex outputs,
            Regex excludeIo)
        {
            if (card == null ||
                !IoListLogic.TryParseCardSheetName(card.SheetName, out string rack, out string slot, out string cardNo))
                return null;

            int headerIdx = FindHeaderRow(card.AllRows);
            if (headerIdx < 0) return null;

            string[] headers = card.AllRows[headerIdx];
            int iBit = IoListLogic.FindColumn(headers, "bit");
            int iMnemonic = IoListLogic.FindColumn(headers, "mnemonic");
            int iId = IoListLogic.FindColumn(headers, "id");

            // An IO point whose language-1 descriptions AND device are all empty is an UNUSED spare —
            // flagged on the row label so the view can fade it. Undecidable when the columns are absent.
            int iDesc1 = IoListLogic.FindColumn(headers, "Description language 1 (1° part", "description english language");
            int iDesc2 = IoListLogic.FindColumn(headers, "Description language 1 (2° part");
            int iDevice = IoListLogic.FindColumn(headers, "device");
            bool canJudgeUnused = iDesc1 >= 0 && iDevice >= 0;

            var sheet = new BridgeSheetModel
            {
                SheetName = card.SheetName,
                CardKey = SignalRef.MakeCardKey(rack, slot, cardNo),
                Card = IoListLogic.Norm(cardNo),
                IoHeaders = headers,
                IoHeaderSourceRow = card.FirstRowNumber + headerIdx,
                BitColumn = iBit,
                SourceColumnWidths = card.ColumnWidths,
            };

            // Phase 0 — parse the spine, no claiming yet.
            var entries = new List<PendingEntry>();
            for (int r = headerIdx + 1; r < card.AllRows.Count; r++)
            {
                string[] fields = card.AllRows[r] ?? new string[0];
                string rawBit = At(fields, iBit);
                string rawMnemonic = At(fields, iMnemonic).Trim();
                string id = At(fields, iId).Trim();

                // The exclude regex filters MNEMONICS, not rows. A cell may carry a generator key
                // beside a real mnemonic ("BUS_CABINETAL00159/SAFETY_NUMDOORS_16" — the key tells
                // another tool to generate 16 door blocks at that node, in plain mnemonic syntax);
                // the key drops out of the comparison, the mnemonic beside it still joins. Only a
                // cell whose EVERY part matches is excluded whole.
                FilterMnemonics(rawMnemonic, excludeIo, out string mnemonic, out string excludedParts, out bool allExcluded);

                var row = new BridgeRow
                {
                    IoSourceRow = card.FirstRowNumber + r,
                    RawBitText = rawBit,
                    IoValues = fields,
                    ExcludedMnemonics = excludedParts,
                };
                PendingEntry entry;

                if (IoListLogic.TryParseAddress(rawBit, out char dir, out string node, out string off, out string bit))
                {
                    row.Kind = BridgeRowKind.Signal;
                    row.RowLabel = "Signal";
                    row.Signal = new SignalRef
                    {
                        CardKey = sheet.CardKey,
                        Address = IoListLogic.BuildAddress(dir, node, off, bit),
                        Mnemonic = mnemonic,
                    };
                    CsvSideIndex side = dir == 'O' ? outputs : inputs;
                    entry = new PendingEntry
                    {
                        Row = row,
                        Side = side,
                        ClaimKey = side.AddressKeyOf(node, off, sheet.Card, bit),
                        NameKeys = IoListLogic.MnemonicKeys(mnemonic),
                    };
                }
                else if (id.Length > 0 && rawBit.Trim().Length == 0)
                {
                    row.Kind = BridgeRowKind.Module;
                    row.RowLabel = "Module";
                    row.Signal = new SignalRef
                    {
                        CardKey = sheet.CardKey,
                        Address = IoListLogic.ModuleAddress(id),
                        Mnemonic = IoListLogic.StripBusPrefix(mnemonic),
                    };
                    // Bus-diagnostic CSV signals (type 11) live in the Inputs file and join by node.
                    entry = new PendingEntry
                    {
                        Row = row,
                        Side = inputs,
                        ClaimKey = inputs.NodeKeyOf(sheet.Card, id),
                        NameKeys = IoListLogic.MnemonicKeys(row.Signal.Mnemonic),
                    };
                }
                else
                {
                    continue;   // decoration/blank row — not part of the joinable spine
                }

                if (allExcluded)
                {
                    entry.Excluded = true;
                    row.RowLabel += " · excluded";
                }
                else if (excludedParts.Length > 0)
                {
                    row.RowLabel += " · filtered";
                }

                if (canJudgeUnused &&
                    At(fields, iDesc1).Trim().Length == 0 &&
                    (iDesc2 < 0 || At(fields, iDesc2).Trim().Length == 0) &&
                    At(fields, iDevice).Trim().Length == 0)
                {
                    row.RowLabel += " · unused";
                }
                entries.Add(entry);
            }

            // Phase 1 — name-aware claims: each record goes to the IO row whose mnemonic carries its
            // name, so struct binding points above the real DO never steal the DO's record.
            foreach (PendingEntry entry in entries)
                if (!entry.Excluded)
                    entry.Claimed.AddRange(entry.Side.ClaimNameMatches(entry.ClaimKey, entry.NameKeys));

            // Phase 2 — positional fallback for rows the names left empty-handed.
            foreach (PendingEntry entry in entries)
                if (!entry.Excluded && entry.Claimed.Count == 0)
                    entry.Claimed.AddRange(entry.Side.ClaimAllRemaining(entry.ClaimKey));

            // Phase 3 — opted-in virtuals (the compare-virtuals regex: virtuals that are EXPECTED to
            // mirror IO-list entries, the CLASSDIAGAL pattern) attach BY NAME wherever the mnemonic
            // appears, and read as ordinary matches.
            foreach (PendingEntry entry in entries)
                if (!entry.Excluded)
                    entry.Claimed.AddRange(entry.Side.ClaimVirtualsByName(entry.NameKeys));

            // Phase 3b — a signal STILL without a record whose mnemonic some other VIRTUAL record
            // carries (type 5/6/9/10): that record is its twin. It binds FLAGGED — the point looks
            // present but acquires nothing from the field — so the row shows the twin and offers a
            // type change instead of integrating a double.
            foreach (PendingEntry entry in entries)
            {
                if (entry.Excluded || entry.Claimed.Count > 0 || entry.Row.Kind != BridgeRowKind.Signal) continue;
                List<int> twins = entry.Side.ClaimVirtualsByName(entry.NameKeys, anyVirtual: true);
                if (twins.Count == 0) continue;
                entry.Claimed.AddRange(twins);
                entry.VirtualTwin = true;
            }

            foreach (PendingEntry entry in entries)
            {
                if (!entry.Excluded) Bind(entry.Row, entry.Side, entry.Claimed, entry.Row.Signal.Mnemonic, entry.VirtualTwin);
                sheet.SpineRows.Add(entry.Row);
                Register(model, entry.Row);
            }
            return sheet;
        }

        private static void Bind(BridgeRow row, CsvSideIndex side, List<int> recordIdxs, string ioMnemonic,
            bool virtualTwin = false)
        {
            if (recordIdxs == null || recordIdxs.Count == 0)
            {
                // A point the CSV should carry but does not — the integration candidate. Module rows
                // count too: a named BUS module without its type-11 record is a real gap, not a match.
                if (row.Kind == BridgeRowKind.Signal ||
                    (row.Kind == BridgeRowKind.Module && row.Signal.Mnemonic.Trim().Length > 0))
                    row.MatchVerdict = "missing";
                return;
            }

            string[] firstFields = side.FieldsOf(recordIdxs[0]);
            var binding = new CsvBinding
            {
                File = side.Target,
                RecordIndex = recordIdxs[0],
                Fields = firstFields,
                Disabled = firstFields.Length > 0 && firstFields[0] != null &&
                           firstFields[0].StartsWith("#", StringComparison.Ordinal),
            };
            foreach (int idx in recordIdxs)
            {
                binding.RecordIndexes.Add(idx);
                binding.AllFields.Add(side.FieldsOf(idx));
            }
            row.Csv = binding;

            // The canonical block joins every record's value with "/", distinct, in file order — so a
            // point carried by two CSV rows reads "A/B", matching how the IO list writes its mnemonics.
            row.CsvValues = JoinCanonical(side, recordIdxs);
            if (binding.Disabled) row.RowLabel += " · disabled";
            if (binding.IsMulti) row.RowLabel += " · " + recordIdxs.Count + " rows";

            string csvNames = row.CsvValues[CsvCanonical.Name];
            row.MatchVerdict = IoListLogic.MatchRatio(csvNames, ioMnemonic, ioMnemonic);

            if (virtualTwin)
            {
                // Never reads as a clean match: the names agree, the acquisition does not exist.
                binding.Virtual = true;
                row.RowLabel += " · virtual twin";
                row.MatchVerdict = "virtual";
            }
        }

        private static string[] JoinCanonical(CsvSideIndex side, List<int> recordIdxs)
        {
            if (recordIdxs.Count == 1) return side.CanonicalValues(recordIdxs[0]);

            var joined = new string[CsvCanonical.Labels.Length];
            for (int c = 0; c < joined.Length; c++) joined[c] = "";

            var seen = new List<string>[CsvCanonical.Labels.Length];
            for (int c = 0; c < seen.Length; c++) seen[c] = new List<string>();

            foreach (int idx in recordIdxs)
            {
                string[] values = side.CanonicalValues(idx);
                for (int c = 0; c < joined.Length; c++)
                {
                    string v = values[c];
                    if (v.Length > 0 && !seen[c].Contains(v)) seen[c].Add(v);
                }
            }
            for (int c = 0; c < joined.Length; c++) joined[c] = string.Join("/", seen[c]);
            return joined;
        }

        private static void CollectCsvOnly(BridgeWorkspaceModel model, CsvSideIndex side, BridgeTarget target)
        {
            foreach (var leftover in side.Unclaimed())
            {
                BridgeSheetModel sheet = FindSheetByCard(model, leftover.Board);
                if (sheet == null) continue;   // a card outside this workspace's selection

                string[] fields = side.FieldsOf(leftover.RecordIndex);
                string[] canonical = side.CanonicalValues(leftover.RecordIndex);
                var row = new BridgeRow
                {
                    Kind = BridgeRowKind.CsvOnly,
                    RowLabel = "CSV-only" + (leftover.Virtual ? " · virtual" : ""),
                    Signal = new SignalRef
                    {
                        CardKey = sheet.CardKey,
                        Address = leftover.Address,
                        Mnemonic = canonical[CsvCanonical.Name],
                    },
                    CsvValues = canonical,
                    Csv = new CsvBinding
                    {
                        File = target,
                        RecordIndex = leftover.RecordIndex,
                        Fields = fields,
                        Disabled = fields.Length > 0 && fields[0] != null &&
                                   fields[0].StartsWith("#", StringComparison.Ordinal),
                    },
                };
                if (row.Csv.Disabled) row.RowLabel += " · disabled";

                sheet.CsvOnlyRows.Add(row);
                Register(model, row);
            }
        }

        private static void Register(BridgeWorkspaceModel model, BridgeRow row)
        {
            string key = row.Signal.IdentityKey();
            if (key.Length > 0 && !model.SignalIndex.ContainsKey(key)) model.SignalIndex[key] = row;

            string byName = row.Signal.MnemonicIdentityKey();
            if (byName.Length > 0 && !model.SignalIndex.ContainsKey(byName)) model.SignalIndex[byName] = row;

            // Each individual mnemonic also aliases the row, so a per-record action (identified by ONE
            // name of a multi-name point) still finds its spine row for highlighting and re-anchoring.
            foreach (string part in IoListLogic.MnemonicKeys(row.Signal.Mnemonic))
            {
                string alias = row.Signal.CardKey + "~" + part;
                if (!model.SignalIndex.ContainsKey(alias)) model.SignalIndex[alias] = row;
            }
        }

        private static BridgeSheetModel FindSheetByCard(BridgeWorkspaceModel model, string board)
        {
            string card = IoListLogic.Norm(board);
            foreach (BridgeSheetModel sheet in model.Sheets)
                if (sheet.Card == card) return sheet;
            return null;
        }

        private static string At(string[] fields, int index) =>
            index >= 0 && index < fields.Length && fields[index] != null ? fields[index] : "";

        /// <summary>
        /// Splits a "/"-joined mnemonic cell into the parts that enter the comparison and the parts
        /// the exclude regex filters out (original spelling, original order within each group).
        /// <paramref name="allExcluded"/> is true when the cell had parts and none survived — the
        /// row is then ignored whole and keeps its full name for display and identity.
        /// </summary>
        internal static void FilterMnemonics(string cell, Regex exclude,
            out string kept, out string excluded, out bool allExcluded)
        {
            var keptParts = new List<string>();
            var droppedParts = new List<string>();
            foreach (string raw in (cell ?? "").Split('/'))
            {
                string part = raw.Trim();
                if (part.Length == 0) continue;
                if (exclude != null && exclude.IsMatch(part)) droppedParts.Add(part);
                else keptParts.Add(part);
            }
            allExcluded = keptParts.Count == 0 && droppedParts.Count > 0;
            kept = allExcluded ? (cell ?? "").Trim() : string.Join("/", keptParts);
            excluded = string.Join("/", droppedParts);
        }

        /// <summary>
        /// One CSV file's records indexed for the join, tracking which records a spine row claimed.
        /// Virtual rows (type 5/6) are indexed for the CSV-only bucket but never join by address —
        /// their address fields are frequently stale copy-paste values.
        /// </summary>
        private sealed class CsvSideIndex
        {
            public readonly BridgeTarget Target;
            private readonly char _direction;   // 'I' or 'O' — prefixes CSV-only addresses
            private readonly CsvDocument _doc;
            private readonly CsvColumnsMap _map;
            private readonly Dictionary<string, List<int>> _byAddress = new Dictionary<string, List<int>>();
            private readonly Dictionary<string, List<int>> _byNode = new Dictionary<string, List<int>>();
            private readonly Dictionary<string, List<int>> _virtualByName = new Dictionary<string, List<int>>();
            private readonly Dictionary<string, List<int>> _anyVirtualByName = new Dictionary<string, List<int>>();
            private readonly List<Row> _rows = new List<Row>();

            private sealed class Row
            {
                public int RecordIndex;
                public string Board = "", Address = "", NameKey = "";
                public bool Virtual, Claimed;
            }

            public sealed class Leftover
            {
                public int RecordIndex;
                public string Board = "", Address = "";
                public bool Virtual;
            }

            public CsvSideIndex(
                CsvDocument doc, CsvColumnsMap map, BridgeTarget target, char direction, Regex includeVirtuals)
            {
                Target = target;
                _direction = direction;
                _doc = doc;
                _map = map;
                if (doc == null || map == null) return;

                for (int i = map.HeaderIndex + 1; i < doc.Records.Count; i++)
                {
                    string[] f = doc.Records[i].Fields;
                    if (f.Length == 0 || f[0] == null || f[0].StartsWith("!", StringComparison.Ordinal)) continue;

                    string node = At(f, map.Node), board = At(f, map.Board);
                    if (node.Trim().Length == 0 || board.Trim().Length == 0) continue;

                    string type = At(f, map.Type);
                    string name = At(f, map.Name).Trim();
                    var row = new Row
                    {
                        RecordIndex = i,
                        Board = IoListLogic.Norm(board),
                        NameKey = IoListLogic.MnemonicKey(name),
                    };

                    if (IoListLogic.IsVirtualType(type))
                    {
                        row.Virtual = true;
                        // Every virtual is a twin candidate by name (their address fields stay
                        // meaningless); opted-in ones (the CLASSDIAGAL pattern) also join BY NAME
                        // next to real records.
                        if (name.Length > 0) Add(_anyVirtualByName, row.NameKey, _rows.Count);
                        if (includeVirtuals != null && name.Length > 0 && includeVirtuals.IsMatch(name))
                            Add(_virtualByName, row.NameKey, _rows.Count);
                    }
                    else if (IoListLogic.IsBusDiagnosticType(type))
                    {
                        row.Address = IoListLogic.ModuleAddress(node);
                        Add(_byNode, IoListLogic.NodeKey(board, node), _rows.Count);
                    }
                    else
                    {
                        string off = At(f, map.Offset), bit = At(f, map.Bit);
                        if (off.Trim().Length > 0 && bit.Trim().Length > 0)
                        {
                            row.Address = IoListLogic.BuildAddress(_direction, node, off, bit);
                            Add(_byAddress, Key(node, off, board, bit), _rows.Count);
                        }
                    }
                    _rows.Add(row);
                }
            }

            public string AddressKeyOf(string node, string off, string board, string bit) =>
                Key(node, off, board, bit);

            public string NodeKeyOf(string board, string node) => IoListLogic.NodeKey(board, node);

            /// <summary>Pass 1: the unclaimed records at the key whose NAME is among the row's mnemonics.</summary>
            public List<int> ClaimNameMatches(string claimKey, List<string> nameKeys)
            {
                var taken = new List<int>();
                if (nameKeys == null || nameKeys.Count == 0) return taken;
                foreach (int rowIdx in RowsAt(claimKey))
                {
                    Row row = _rows[rowIdx];
                    if (row.Claimed || row.NameKey.Length == 0 || !nameKeys.Contains(row.NameKey)) continue;
                    row.Claimed = true;
                    taken.Add(row.RecordIndex);
                }
                return taken;
            }

            /// <summary>Pass 2: every record still unclaimed at the key — the positional fallback.</summary>
            public List<int> ClaimAllRemaining(string claimKey)
            {
                var taken = new List<int>();
                foreach (int rowIdx in RowsAt(claimKey))
                {
                    Row row = _rows[rowIdx];
                    if (row.Claimed) continue;
                    row.Claimed = true;
                    taken.Add(row.RecordIndex);
                }
                return taken;
            }

            /// <summary>Pass 3: opted-in virtual records carrying any of the row's mnemonics —
            /// or, with <paramref name="anyVirtual"/>, EVERY virtual record that does (the twins).</summary>
            public List<int> ClaimVirtualsByName(List<string> nameKeys, bool anyVirtual = false)
            {
                var taken = new List<int>();
                if (nameKeys == null) return taken;
                Dictionary<string, List<int>> index = anyVirtual ? _anyVirtualByName : _virtualByName;
                foreach (string key in nameKeys)
                {
                    if (!index.TryGetValue(key, out List<int> list)) continue;
                    foreach (int rowIdx in list)
                    {
                        Row row = _rows[rowIdx];
                        if (row.Claimed) continue;
                        row.Claimed = true;
                        taken.Add(row.RecordIndex);
                    }
                }
                return taken;
            }

            private IEnumerable<int> RowsAt(string claimKey)
            {
                // The key routes itself: node keys carry '#', address keys carry '|'.
                Dictionary<string, List<int>> map = claimKey != null && claimKey.IndexOf('#') >= 0 ? _byNode : _byAddress;
                if (claimKey != null && map.TryGetValue(claimKey, out List<int> list))
                    foreach (int rowIdx in list) yield return rowIdx;
            }

            public string[] FieldsOf(int recordIndex) => _doc.Records[recordIndex].Fields;

            public string[] CanonicalValues(int recordIndex)
            {
                var values = new string[CsvCanonical.Labels.Length];
                string[] fields = _doc.Records[recordIndex].Fields;
                for (int c = 0; c < values.Length; c++)
                {
                    int field = CsvCanonical.FieldIndex(_map, c);
                    values[c] = field >= 0 && field < fields.Length && fields[field] != null ? fields[field] : "";
                }
                return values;
            }

            public IEnumerable<Leftover> Unclaimed()
            {
                foreach (Row row in _rows)
                    if (!row.Claimed)
                        yield return new Leftover
                        {
                            RecordIndex = row.RecordIndex,
                            Board = row.Board,
                            Address = row.Address,
                            Virtual = row.Virtual,
                        };
            }

            private static string Key(string node, string off, string board, string bit) =>
                IoListLogic.Norm(node) + "|" + IoListLogic.Norm(off) + "|" +
                IoListLogic.Norm(board) + "|" + IoListLogic.Norm(bit);

            private static void Add(Dictionary<string, List<int>> map, string key, int rowIdx)
            {
                if (!map.TryGetValue(key, out List<int> list)) map[key] = list = new List<int>();
                list.Add(rowIdx);
            }
        }
    }
}
