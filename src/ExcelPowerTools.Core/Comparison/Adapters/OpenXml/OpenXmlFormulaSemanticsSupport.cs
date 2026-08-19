using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.Comparison.Adapters.OpenXml
{
    internal sealed class OpenXmlSharedFormulaCatalog
    {
        internal const long MaxSharedFormulaGroupCells = 1000000L;
        internal const long MaxSharedFormulaDerivedCharacters = 64L * 1024 * 1024;

        private readonly IDictionary<string, OpenXmlSharedFormulaResolution> _resolutions;

        private OpenXmlSharedFormulaCatalog(
            IDictionary<string, OpenXmlSharedFormulaResolution> resolutions)
        {
            _resolutions = resolutions;
        }

        internal static OpenXmlSharedFormulaCatalog Create(
            IEnumerable<KeyValuePair<string, Cell>> cells,
            long derivedCharacterLimit = MaxSharedFormulaDerivedCharacters)
        {
            if (derivedCharacterLimit < 0)
                throw new ArgumentOutOfRangeException(nameof(derivedCharacterLimit));
            var resolutions = new Dictionary<string, OpenXmlSharedFormulaResolution>(StringComparer.Ordinal);
            var groups = new Dictionary<uint, SharedGroupBuilder>();

            foreach (KeyValuePair<string, Cell> pair in cells)
            {
                CellFormula formula = pair.Value.CellFormula;
                if (formula == null || !string.Equals(Attribute(formula, "t"), "shared", StringComparison.Ordinal))
                    continue;

                string storedText = formula.Text ?? "";
                string rawIndex = Attribute(formula, "si");
                if (!uint.TryParse(rawIndex, NumberStyles.None, CultureInfo.InvariantCulture, out uint index))
                {
                    resolutions[pair.Key] = OpenXmlSharedFormulaResolution.Failed(
                        storedText,
                        null,
                        Attribute(formula, "ref"),
                        "The shared formula has a missing or malformed si group index.");
                    continue;
                }

                if (!groups.TryGetValue(index, out SharedGroupBuilder group))
                {
                    group = new SharedGroupBuilder(index);
                    groups.Add(index, group);
                }
                group.Members.Add(new SharedMember(pair.Key, formula, storedText));
            }

            var finalizedGroups = new List<SharedGroup>();
            foreach (SharedGroupBuilder builder in groups.Values.OrderBy(item => item.Index))
            {
                SharedGroup group = builder.FinalizeGroup(resolutions);
                if (group != null) finalizedGroups.Add(group);
            }

            // Complete rectangular membership plus unique physical cell references proves that two finalized
            // groups cannot overlap: an overlap coordinate would have to be a member of both si groups. Incomplete
            // groups were already rejected, so a quadratic pairwise range scan adds no evidence.
            long derivedCharacters = 0;
            bool scopeExhausted = false;
            foreach (SharedGroup group in finalizedGroups)
            {
                if (!group.TryGetDerivedCharacterUpperBound(out long groupCharacters) ||
                    groupCharacters > derivedCharacterLimit - derivedCharacters)
                {
                    scopeExhausted = true;
                    break;
                }

                derivedCharacters += groupCharacters;
            }

            if (scopeExhausted)
            {
                string reason =
                    "Reconstructing the shared-formula scope would exceed its deterministic workbook-partitioned " +
                    "derived-formula evidence limit of " +
                    derivedCharacterLimit.ToString(CultureInfo.InvariantCulture) +
                    " characters. No groups in this worksheet were translated.";
                foreach (SharedGroup group in finalizedGroups) group.FailAll(resolutions, reason);
            }
            else
            {
                foreach (SharedGroup group in finalizedGroups) group.Resolve(resolutions);
            }

            return new OpenXmlSharedFormulaCatalog(resolutions);
        }

        internal bool TryGet(string cellReference, out OpenXmlSharedFormulaResolution resolution) =>
            _resolutions.TryGetValue(cellReference, out resolution);

        private static string Attribute(OpenXmlElement element, string localName)
        {
            foreach (OpenXmlAttribute attribute in element.GetAttributes())
            {
                if (string.IsNullOrEmpty(attribute.NamespaceUri) &&
                    string.Equals(attribute.LocalName, localName, StringComparison.Ordinal))
                    return attribute.Value ?? "";
            }
            return "";
        }

        private sealed class SharedGroupBuilder
        {
            internal SharedGroupBuilder(uint index)
            {
                Index = index;
                Members = new List<SharedMember>();
            }

            internal uint Index { get; }
            internal IList<SharedMember> Members { get; }

            internal SharedGroup FinalizeGroup(
                IDictionary<string, OpenXmlSharedFormulaResolution> resolutions)
            {
                List<SharedMember> masters = Members
                    .Where(member =>
                        !string.IsNullOrEmpty(Attribute(member.Formula, "ref")) &&
                        !string.IsNullOrEmpty(member.StoredText))
                    .ToList();

                if (masters.Count != 1)
                {
                    string reason = masters.Count == 0
                        ? "The shared-formula group has no unique master with both ref and stored formula text."
                        : "The shared-formula group has more than one master with ref and stored formula text.";
                    FailAll(resolutions, reason);
                    return null;
                }

                SharedMember master = masters[0];
                string range = Attribute(master.Formula, "ref");
                foreach (SharedMember member in Members)
                {
                    if (ReferenceEquals(member, master)) continue;
                    if (!string.IsNullOrEmpty(member.StoredText) ||
                        !string.IsNullOrEmpty(Attribute(member.Formula, "ref")))
                    {
                        FailAll(
                            resolutions,
                            "A non-master shared-formula member contains stored formula text or a ref attribute. " +
                            "Only the empty-text/no-ref follower form is reconstructed.");
                        return null;
                    }
                }

                if (!OpenXmlSharedFormulaTranslator.TryParseRange(
                        range,
                        out int firstRow,
                        out int firstColumn,
                        out int lastRow,
                        out int lastColumn))
                {
                    FailAll(resolutions, "The shared-formula master ref is not a valid rectangular cell range.");
                    return null;
                }

                if (!Within(master.CellReference, firstRow, firstColumn, lastRow, lastColumn))
                {
                    FailAll(resolutions, "The shared-formula master cell is outside its declared ref range.");
                    return null;
                }

                long rowCount = (long)lastRow - firstRow + 1L;
                long columnCount = (long)lastColumn - firstColumn + 1L;
                long area = rowCount * columnCount;
                if (area > MaxSharedFormulaGroupCells)
                {
                    FailAll(
                        resolutions,
                        "The shared-formula master ref declares " +
                        area.ToString(CultureInfo.InvariantCulture) +
                        " cells, above the safe reconstruction limit of " +
                        MaxSharedFormulaGroupCells.ToString(CultureInfo.InvariantCulture) + ".");
                    return null;
                }

                foreach (SharedMember member in Members)
                {
                    if (!Within(member.CellReference, firstRow, firstColumn, lastRow, lastColumn))
                    {
                        FailAll(
                            resolutions,
                            "A shared-formula member is outside the master ref range; OOXML leaves this case " +
                            "implementation-defined.");
                        return null;
                    }
                }

                // The declared ref is the group's translation envelope, not a promise that every coordinate uses
                // this shared index. Excel legitimately emits sparse groups when cells inside that envelope are
                // blank, constants, ordinary formulas, or members of another shared group. The physical cells that
                // actually carry this si are authoritative; each one has already been proved unique and in-range.

                return new SharedGroup(
                    Index,
                    master,
                    range,
                    Members);
            }

            private void FailAll(
                IDictionary<string, OpenXmlSharedFormulaResolution> resolutions,
                string reason)
            {
                foreach (SharedMember member in Members)
                {
                    resolutions[member.CellReference] = OpenXmlSharedFormulaResolution.Failed(
                        member.StoredText,
                        Index,
                        Attribute(member.Formula, "ref"),
                        reason);
                }
            }
        }

        private sealed class SharedGroup
        {
            private readonly IList<SharedMember> _members;

            internal SharedGroup(
                uint index,
                SharedMember master,
                string range,
                IList<SharedMember> members)
            {
                Index = index;
                Master = master;
                Range = range;
                _members = members;
            }

            internal uint Index { get; }
            internal SharedMember Master { get; }
            internal string Range { get; }

            internal bool TryGetDerivedCharacterUpperBound(out long characters)
            {
                characters = 0;
                long masterLength = Master.StoredText.Length;
                long followerCount = _members.Count - 1L;

                // A translated A1/row/column token is at most a small fixed multiple of its source token. Nine
                // times the complete master plus 32 characters is a conservative per-follower lexical bound,
                // and is checked before any translated follower string is allocated.
                if (masterLength > (long.MaxValue - 32L) / 9L) return false;
                long perFollower = masterLength * 9L + 32L;
                if (followerCount > 0 && perFollower > (long.MaxValue - masterLength) / followerCount)
                    return false;
                characters = masterLength + followerCount * perFollower;
                long structuralOverhead = 2048L + Range.Length * 8L + Master.CellReference.Length * 8L;
                if (structuralOverhead > long.MaxValue - characters) return false;
                characters += structuralOverhead;
                return true;
            }

            internal void FailAll(
                IDictionary<string, OpenXmlSharedFormulaResolution> resolutions,
                string reason)
            {
                foreach (SharedMember member in _members)
                {
                    resolutions[member.CellReference] = OpenXmlSharedFormulaResolution.Failed(
                        member.StoredText,
                        Index,
                        Range,
                        reason);
                }
            }

            internal void Resolve(IDictionary<string, OpenXmlSharedFormulaResolution> resolutions)
            {
                foreach (SharedMember member in _members)
                {
                    if (resolutions.ContainsKey(member.CellReference)) continue;
                    if (string.Equals(member.CellReference, Master.CellReference, StringComparison.Ordinal))
                    {
                        resolutions[member.CellReference] = OpenXmlSharedFormulaResolution.ResolvedMaster(
                            member.StoredText,
                            Index,
                            Range,
                            Master.CellReference);
                        continue;
                    }

                    if (OpenXmlSharedFormulaTranslator.TryTranslate(
                            Master.StoredText,
                            Master.CellReference,
                            member.CellReference,
                            out string translated,
                            out string reason))
                    {
                        resolutions[member.CellReference] = OpenXmlSharedFormulaResolution.ResolvedFollower(
                            member.StoredText,
                            translated,
                            Index,
                            Range,
                            Master.CellReference);
                    }
                    else
                    {
                        resolutions[member.CellReference] = OpenXmlSharedFormulaResolution.Failed(
                            member.StoredText,
                            Index,
                            Range,
                            reason,
                            Master.CellReference);
                    }
                }
            }
        }

        private sealed class SharedMember
        {
            internal SharedMember(string cellReference, CellFormula formula, string storedText)
            {
                CellReference = cellReference;
                Formula = formula;
                StoredText = storedText;
            }

            internal string CellReference { get; }
            internal CellFormula Formula { get; }
            internal string StoredText { get; }
        }

        private static bool Within(
            string reference,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn) =>
            OpenXmlSharedFormulaTranslator.TryParseCellReference(reference, out int row, out int column) &&
            row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
    }

    internal sealed class OpenXmlSharedFormulaResolution
    {
        private OpenXmlSharedFormulaResolution(
            string storedText,
            string effectiveText,
            uint? sharedIndex,
            string range,
            string masterCellReference,
            bool reconstructed,
            string failureReason)
        {
            StoredText = storedText ?? "";
            EffectiveText = effectiveText;
            SharedIndex = sharedIndex;
            Range = range;
            MasterCellReference = masterCellReference;
            WasReconstructed = reconstructed;
            FailureReason = failureReason;
        }

        internal string StoredText { get; }
        internal string EffectiveText { get; }
        internal uint? SharedIndex { get; }
        internal string Range { get; }
        internal string MasterCellReference { get; }
        internal bool WasReconstructed { get; }
        internal string FailureReason { get; }
        internal bool IsResolved => FailureReason == null && EffectiveText != null;
        internal bool IsMaster => IsResolved && !WasReconstructed;

        internal static OpenXmlSharedFormulaResolution ResolvedMaster(
            string storedText,
            uint index,
            string range,
            string masterCellReference) =>
            new OpenXmlSharedFormulaResolution(
                storedText,
                storedText,
                index,
                range,
                masterCellReference,
                false,
                null);

        internal static OpenXmlSharedFormulaResolution ResolvedFollower(
            string storedText,
            string effectiveText,
            uint index,
            string range,
            string masterCellReference) =>
            new OpenXmlSharedFormulaResolution(
                storedText,
                effectiveText,
                index,
                range,
                masterCellReference,
                true,
                null);

        internal static OpenXmlSharedFormulaResolution Failed(
            string storedText,
            uint? index,
            string range,
            string reason,
            string masterCellReference = null) =>
            new OpenXmlSharedFormulaResolution(
                storedText,
                null,
                index,
                range,
                masterCellReference,
                false,
                reason ?? "The shared formula could not be reconstructed safely.");
    }

    internal sealed class OpenXmlFormulaMetadataLimits
    {
        internal static readonly OpenXmlFormulaMetadataLimits Default =
            new OpenXmlFormulaMetadataLimits(
                OpenXmlFormulaMetadataCatalog.MaxCatalogRecords,
                OpenXmlFormulaMetadataCatalog.MaxCatalogEvidenceCharacters,
                OpenXmlFormulaMetadataCatalog.MaxDerivedRecords,
                OpenXmlFormulaMetadataCatalog.MaxDerivedEvidenceCharacters,
                OpenXmlSharedFormulaCatalog.MaxSharedFormulaDerivedCharacters,
                OpenXmlWorkbookSnapshotAdapter.MaxTableFormulaObservations,
                OpenXmlFormulaMetadataCatalog.MaxDerivedEvidenceCharacters);

        internal OpenXmlFormulaMetadataLimits(
            long maxCatalogRecords,
            long maxCatalogEvidenceCharacters,
            long maxDerivedRecords,
            long maxDerivedEvidenceCharacters,
            long maxSharedFormulaDerivedCharacters = OpenXmlSharedFormulaCatalog.MaxSharedFormulaDerivedCharacters,
            long maxTableFormulaRecords = OpenXmlWorkbookSnapshotAdapter.MaxTableFormulaObservations,
            long maxTableFormulaEvidenceCharacters = OpenXmlFormulaMetadataCatalog.MaxDerivedEvidenceCharacters)
        {
            if (maxCatalogRecords <= 0) throw new ArgumentOutOfRangeException(nameof(maxCatalogRecords));
            if (maxCatalogEvidenceCharacters <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxCatalogEvidenceCharacters));
            if (maxDerivedRecords <= 0) throw new ArgumentOutOfRangeException(nameof(maxDerivedRecords));
            if (maxDerivedEvidenceCharacters <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxDerivedEvidenceCharacters));
            if (maxSharedFormulaDerivedCharacters <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxSharedFormulaDerivedCharacters));
            if (maxTableFormulaRecords <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxTableFormulaRecords));
            if (maxTableFormulaEvidenceCharacters <= 0)
                throw new ArgumentOutOfRangeException(nameof(maxTableFormulaEvidenceCharacters));
            MaxCatalogRecords = maxCatalogRecords;
            MaxCatalogEvidenceCharacters = maxCatalogEvidenceCharacters;
            MaxDerivedRecords = maxDerivedRecords;
            MaxDerivedEvidenceCharacters = maxDerivedEvidenceCharacters;
            MaxSharedFormulaDerivedCharacters = maxSharedFormulaDerivedCharacters;
            MaxTableFormulaRecords = maxTableFormulaRecords;
            MaxTableFormulaEvidenceCharacters = maxTableFormulaEvidenceCharacters;
        }

        internal long MaxCatalogRecords { get; }
        internal long MaxCatalogEvidenceCharacters { get; }
        internal long MaxDerivedRecords { get; }
        internal long MaxDerivedEvidenceCharacters { get; }
        internal long MaxSharedFormulaDerivedCharacters { get; }
        internal long MaxTableFormulaRecords { get; }
        internal long MaxTableFormulaEvidenceCharacters { get; }
    }

    internal sealed class OpenXmlFormulaMetadataCatalog
    {
        internal const long MaxCatalogRecords = 1000000L;
        internal const long MaxCatalogEvidenceCharacters = 64L * 1024 * 1024;
        internal const long MaxDerivedRecords = 1000000L;
        internal const long MaxDerivedEvidenceCharacters = 64L * 1024 * 1024;

        private const string DynamicArrayMetadataName = "XLDAPR";
        private const string DynamicArrayNamespace =
            "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";

        private readonly IList<MetadataTypeEvidence> _types;
        private readonly IDictionary<string, FutureMetadataEvidence> _futureMetadata;
        private readonly IList<CellMetadataEvidence> _cellMetadata;
        private readonly IDictionary<uint, OpenXmlFormulaMetadataResolution> _resolutionCache;
        private readonly OpenXmlFormulaMetadataLimits _limits;
        private readonly bool _hasMetadataPart;
        private readonly string _catalogFailure;
        private IDictionary<string, OpenXmlFormulaMetadataResolution> _scopeResolutions;

        internal OpenXmlFormulaMetadataCatalog(WorkbookPart workbookPart)
            : this(workbookPart, OpenXmlFormulaMetadataLimits.Default)
        {
        }

        internal OpenXmlFormulaMetadataCatalog(
            WorkbookPart workbookPart,
            OpenXmlFormulaMetadataLimits limits)
        {
            _limits = limits ?? throw new ArgumentNullException(nameof(limits));
            _types = new List<MetadataTypeEvidence>();
            _futureMetadata = new Dictionary<string, FutureMetadataEvidence>(StringComparer.Ordinal);
            _cellMetadata = new List<CellMetadataEvidence>();
            _resolutionCache = new Dictionary<uint, OpenXmlFormulaMetadataResolution>();
            _scopeResolutions = new Dictionary<string, OpenXmlFormulaMetadataResolution>(StringComparer.Ordinal);

            CellMetadataPart part = workbookPart == null ? null : workbookPart.CellMetadataPart;
            Metadata metadata = part == null ? null : part.Metadata;
            _hasMetadataPart = metadata != null;
            if (metadata == null)
            {
                _catalogFailure = null;
                return;
            }

            var failures = new List<string>();
            long records = 0;
            long characters = 0;
            bool budgetExhausted = false;

            List<MetadataTypes> typeCollections = metadata.Elements<MetadataTypes>().ToList();
            if (typeCollections.Count != 1)
                failures.Add("The metadata part does not contain exactly one metadataTypes collection.");
            if (typeCollections.Count == 1)
            {
                var names = new HashSet<string>(StringComparer.Ordinal);
                foreach (MetadataType type in typeCollections[0].Elements<MetadataType>())
                {
                    if (budgetExhausted) break;
                    string exact = type.OuterXml ?? "";
                    if (!TryChargeCatalog(ref records, ref characters, 1, exact.Length, failures))
                    {
                        budgetExhausted = true;
                        break;
                    }
                    string name = type.Name == null ? "" : type.Name.Value ?? "";
                    if (name.Length > 0 && !names.Add(name))
                        failures.Add("metadataType name '" + name + "' is duplicated.");
                    ((List<MetadataTypeEvidence>)_types).Add(new MetadataTypeEvidence(name, exact));
                }
            }

            var futureNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (FutureMetadata future in metadata.Elements<FutureMetadata>())
            {
                if (budgetExhausted) break;
                string name = future.Name == null ? "" : future.Name.Value ?? "";
                if (name.Length == 0)
                {
                    failures.Add("A futureMetadata collection has no name.");
                    continue;
                }
                if (!futureNames.Add(name))
                {
                    failures.Add("futureMetadata name '" + name + "' is duplicated.");
                    continue;
                }

                var blocks = new List<FutureMetadataBlockEvidence>();
                foreach (FutureMetadataBlock block in future.Elements<FutureMetadataBlock>())
                {
                    string exact = block.OuterXml ?? "";
                    if (!TryChargeCatalog(ref records, ref characters, 1, exact.Length, failures))
                    {
                        budgetExhausted = true;
                        break;
                    }
                    ReadDynamicArrayMarker(block, name, out bool dynamicArray, out string failure);
                    blocks.Add(new FutureMetadataBlockEvidence(exact, dynamicArray, failure));
                }
                if (!budgetExhausted)
                    _futureMetadata.Add(name, new FutureMetadataEvidence(blocks));
            }

            List<CellMetadata> cellCollections = metadata.Elements<CellMetadata>().ToList();
            if (cellCollections.Count != 1)
                failures.Add("The metadata part does not contain exactly one cellMetadata collection.");
            if (cellCollections.Count == 1 && !budgetExhausted)
            {
                foreach (MetadataBlock block in cellCollections[0].Elements<MetadataBlock>())
                {
                    List<MetadataRecord> sourceRecords = block.Elements<MetadataRecord>().ToList();
                    string exact = block.OuterXml ?? "";
                    if (!TryChargeCatalog(
                            ref records,
                            ref characters,
                            1L + sourceRecords.Count,
                            exact.Length,
                            failures))
                    {
                        budgetExhausted = true;
                        break;
                    }

                    var indexedRecords = new List<MetadataRecordEvidence>(sourceRecords.Count);
                    foreach (MetadataRecord record in sourceRecords)
                    {
                        indexedRecords.Add(new MetadataRecordEvidence(
                            Attribute(record, "t"),
                            Attribute(record, "v")));
                    }
                    ((List<CellMetadataEvidence>)_cellMetadata).Add(
                        new CellMetadataEvidence(exact, indexedRecords));
                }
            }

            _catalogFailure = failures.Count == 0 ? null : string.Join(" ", failures);
        }

        internal void PrepareScope(
            IList<KeyValuePair<string, Cell>> readableCells,
            long scopeRecordLimit,
            long scopeCharacterLimit)
        {
            if (readableCells == null) throw new ArgumentNullException(nameof(readableCells));
            if (scopeRecordLimit < 0) throw new ArgumentOutOfRangeException(nameof(scopeRecordLimit));
            if (scopeCharacterLimit < 0) throw new ArgumentOutOfRangeException(nameof(scopeCharacterLimit));
            // Exact cm evidence is a worksheet-scope working cache. Keeping it across scopes can retain rejected
            // strings outside the workbook-derived evidence budget.
            _resolutionCache.Clear();
            List<KeyValuePair<string, Cell>> formulaCells = readableCells
                .Where(pair => pair.Value != null && pair.Value.CellFormula != null &&
                               HasMetadataAssociation(pair.Value))
                .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .ToList();
            var prepared = new Dictionary<string, OpenXmlFormulaMetadataResolution>(StringComparer.Ordinal);
            long records = 0;
            long characters = 0;
            bool exhausted = false;
            foreach (KeyValuePair<string, Cell> pair in formulaCells)
            {
                OpenXmlFormulaMetadataResolution resolution = BuildResolution(
                    pair.Value,
                    scopeRecordLimit - records,
                    scopeCharacterLimit - characters,
                    out bool cellExhausted);
                if (cellExhausted)
                {
                    exhausted = true;
                    break;
                }

                prepared.Add(pair.Key, resolution);
                records += resolution.RecordCount;
                characters += resolution.ExactEvidence == null ? 0 : resolution.ExactEvidence.Length;
            }

            if (!exhausted)
            {
                _scopeResolutions = prepared;
                return;
            }

            const string budgetFailure =
                "The formula metadata derived-evidence record/character budget was exhausted for this " +
                "worksheet. Metadata projection failed as a whole so physical XML order cannot change which " +
                "logical cells retain evidence.";
            prepared.Clear();
            foreach (KeyValuePair<string, Cell> pair in formulaCells)
            {
                prepared.Add(
                    pair.Key,
                    new OpenXmlFormulaMetadataResolution(
                        null,
                        false,
                        budgetFailure,
                        0));
            }
            _scopeResolutions = prepared;
        }

        internal OpenXmlFormulaMetadataResolution Resolve(Cell cell, string cellReference)
        {
            if (!HasMetadataAssociation(cell)) return OpenXmlFormulaMetadataResolution.Empty;
            if (!string.IsNullOrEmpty(cellReference) &&
                _scopeResolutions.TryGetValue(cellReference, out OpenXmlFormulaMetadataResolution resolution))
                return resolution;

            return new OpenXmlFormulaMetadataResolution(
                null,
                false,
                "Formula metadata was not included in the deterministic worksheet preflight scope.",
                0);
        }

        internal int ResolvedEvidenceCacheCountForTesting => _resolutionCache.Count;

        private OpenXmlFormulaMetadataResolution BuildResolution(
            Cell cell,
            long remainingRecords,
            long remainingCharacters,
            out bool budgetExhausted)
        {
            budgetExhausted = false;
            string rawCellIndex = Attribute(cell, "cm");
            string rawValueIndex = Attribute(cell, "vm");
            DocumentFormat.OpenXml.Spreadsheet.ExtensionList extensionList = cell.ExtensionList;
            string extensionEvidence = extensionList == null ? null : extensionList.OuterXml ?? "";
            var evidence = new StringBuilder();
            long requiredCharacters =
                TokenLength("cell-cm") + TokenLength(rawCellIndex) +
                TokenLength("cell-vm") + TokenLength(rawValueIndex);
            long derivedRecords = 1;
            if (extensionList != null)
            {
                requiredCharacters += TokenLength("cell-extLst") + TokenLength(extensionEvidence);
                derivedRecords++;
            }

            // Reject before resolving/caching shared cm evidence when the cell-local payload alone already
            // exhausts the worksheet scope. This keeps the all-or-nothing fallback bounded as well.
            if (derivedRecords > remainingRecords || requiredCharacters > remainingCharacters ||
                requiredCharacters > int.MaxValue)
            {
                budgetExhausted = true;
                return new OpenXmlFormulaMetadataResolution(null, false, null, 0);
            }

            var failures = new List<string>();
            bool dynamicArray = false;
            OpenXmlFormulaMetadataResolution indexedResolution = null;
            if (!string.IsNullOrEmpty(rawValueIndex))
            {
                derivedRecords++;
                failures.Add(
                    "The formula cell has vm value metadata. Its exact index is retained, but rich-value " +
                    "semantics are not interpreted by this adapter.");
            }
            if (extensionList != null)
            {
                failures.Add(
                    "The formula cell has extension markup. Its exact XML is retained, but its semantics are " +
                    "not interpreted by this adapter.");
            }

            if (!string.IsNullOrEmpty(rawCellIndex))
            {
                if (!uint.TryParse(
                        rawCellIndex,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out uint cellIndex) ||
                    cellIndex == 0 || cellIndex > int.MaxValue)
                {
                    failures.Add("The formula cell has a malformed one-based cm metadata index.");
                }
                else if (!_hasMetadataPart || cellIndex > _cellMetadata.Count)
                {
                    failures.Add("The formula cell cm index does not resolve to a cellMetadata block.");
                }
                else if (_catalogFailure != null)
                {
                    failures.Add(_catalogFailure);
                }
                else
                {
                    OpenXmlFormulaMetadataResolution indexed = ResolveCellMetadata(
                        cellIndex,
                        Math.Max(0L, remainingCharacters - requiredCharacters));
                    indexedResolution = indexed;
                    if (indexed.ExactEvidence != null)
                    {
                        requiredCharacters +=
                            TokenLength("resolved-cell-metadata") + TokenLength(indexed.ExactEvidence);
                    }
                    derivedRecords += indexed.RecordCount;
                    dynamicArray = indexed.IsDynamicArray;
                    if (indexed.FailureReason != null) failures.Add(indexed.FailureReason);
                }
            }

            if (derivedRecords > remainingRecords || requiredCharacters > remainingCharacters ||
                requiredCharacters > int.MaxValue)
            {
                budgetExhausted = true;
                return new OpenXmlFormulaMetadataResolution(null, dynamicArray, null, 0);
            }

            evidence.EnsureCapacity((int)requiredCharacters);
            AppendToken(evidence, "cell-cm");
            AppendToken(evidence, rawCellIndex);
            AppendToken(evidence, "cell-vm");
            AppendToken(evidence, rawValueIndex);
            if (extensionList != null)
            {
                AppendToken(evidence, "cell-extLst");
                AppendToken(evidence, extensionEvidence);
            }
            if (!string.IsNullOrEmpty(rawCellIndex) &&
                uint.TryParse(
                    rawCellIndex,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out uint resolvedIndex) &&
                resolvedIndex > 0 && resolvedIndex <= int.MaxValue &&
                _hasMetadataPart && resolvedIndex <= _cellMetadata.Count &&
                _catalogFailure == null)
            {
                OpenXmlFormulaMetadataResolution indexed = indexedResolution;
                if (indexed.ExactEvidence != null)
                {
                    AppendToken(evidence, "resolved-cell-metadata");
                    AppendToken(evidence, indexed.ExactEvidence);
                }
            }

            string exactEvidence = evidence.ToString();
            return new OpenXmlFormulaMetadataResolution(
                exactEvidence,
                dynamicArray,
                failures.Count == 0 ? null : string.Join(" ", failures),
                derivedRecords);
        }

        private static bool HasMetadataAssociation(Cell cell) =>
            cell != null &&
            (!string.IsNullOrEmpty(Attribute(cell, "cm")) ||
             !string.IsNullOrEmpty(Attribute(cell, "vm")) ||
             cell.ExtensionList != null);

        private OpenXmlFormulaMetadataResolution ResolveCellMetadata(
            uint cellIndex,
            long evidenceCharacterLimit)
        {
            if (_resolutionCache.TryGetValue(cellIndex, out OpenXmlFormulaMetadataResolution cached))
                return cached;

            CellMetadataEvidence cell = _cellMetadata[(int)cellIndex - 1];
            var evidence = new StringBuilder();
            var failures = new List<string>();
            bool dynamicArray = false;
            int dynamicMarkers = 0;
            long records = cell.Records.Count;
            if (!TryAppendBounded(
                    evidence,
                    "cell-metadata-block",
                    cell.ExactEvidence,
                    evidenceCharacterLimit))
                failures.Add("One cell metadata resolution exceeds the derived-evidence character budget.");

            foreach (MetadataRecordEvidence record in cell.Records)
            {
                if (!uint.TryParse(
                        record.RawTypeIndex,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out uint typeIndex) ||
                    typeIndex == 0 || typeIndex > _types.Count)
                {
                    failures.Add("A cell metadata record has an unresolved one-based metadata type index.");
                    continue;
                }

                MetadataTypeEvidence type = _types[(int)typeIndex - 1];
                if (!TryAppendBounded(
                        evidence,
                        "metadata-type",
                        type.ExactEvidence,
                        evidenceCharacterLimit))
                {
                    failures.Add("One cell metadata resolution exceeds the derived-evidence character budget.");
                    break;
                }
                if (string.IsNullOrEmpty(type.Name))
                {
                    failures.Add("A referenced cell metadata type has no name.");
                    continue;
                }

                if (!uint.TryParse(
                        record.RawValueIndex,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out uint valueIndex) ||
                    valueIndex > int.MaxValue)
                {
                    failures.Add("A cell metadata record has a malformed zero-based value index.");
                    continue;
                }
                if (!_futureMetadata.TryGetValue(type.Name, out FutureMetadataEvidence future))
                {
                    failures.Add(
                        "A cell metadata record does not resolve to futureMetadata named '" + type.Name + "'.");
                    continue;
                }
                if (valueIndex >= future.Blocks.Count)
                {
                    failures.Add(
                        "A cell metadata value index is outside futureMetadata named '" + type.Name + "'.");
                    continue;
                }

                FutureMetadataBlockEvidence block = future.Blocks[(int)valueIndex];
                if (!TryAppendBounded(
                        evidence,
                        "future-metadata-block",
                        block.ExactEvidence,
                        evidenceCharacterLimit))
                {
                    failures.Add("One cell metadata resolution exceeds the derived-evidence character budget.");
                    break;
                }
                if (block.FailureReason != null) failures.Add(block.FailureReason);
                if (block.IsDynamicArray)
                {
                    dynamicMarkers++;
                    dynamicArray = true;
                }
            }

            if (dynamicMarkers > 1)
            {
                failures.Add(
                    "The cell metadata block resolves more than one fDynamic=true marker; dynamic-array " +
                    "association is ambiguous.");
                dynamicArray = false;
            }

            string exact = failures.Any(item =>
                item.IndexOf("exceeds the derived-evidence", StringComparison.Ordinal) >= 0)
                ? null
                : evidence.ToString();
            var resolved = new OpenXmlFormulaMetadataResolution(
                exact,
                dynamicArray,
                failures.Count == 0 ? null : string.Join(" ", failures),
                records);
            // Cache only bounded resolutions. An oversized/rejected entry must not accumulate across cm indices.
            if (exact != null) _resolutionCache.Add(cellIndex, resolved);
            return resolved;
        }

        private bool TryChargeCatalog(
            ref long records,
            ref long characters,
            long addedRecords,
            long addedCharacters,
            IList<string> failures)
        {
            if (addedRecords < 0 || addedCharacters < 0 ||
                addedRecords > _limits.MaxCatalogRecords - records ||
                addedCharacters > _limits.MaxCatalogEvidenceCharacters - characters)
            {
                if (!failures.Any(item => item.IndexOf("catalog budget", StringComparison.Ordinal) >= 0))
                {
                    failures.Add(
                        "The metadata catalog budget of " +
                        _limits.MaxCatalogRecords.ToString(CultureInfo.InvariantCulture) + " records and " +
                        _limits.MaxCatalogEvidenceCharacters.ToString(CultureInfo.InvariantCulture) +
                        " evidence characters was exceeded.");
                }
                return false;
            }

            records += addedRecords;
            characters += addedCharacters;
            return true;
        }

        private static bool TryAppendBounded(
            StringBuilder target,
            string role,
            string value,
            long evidenceCharacterLimit)
        {
            long required = TokenLength(role) + TokenLength(value);
            if (required > evidenceCharacterLimit - target.Length) return false;
            AppendToken(target, role);
            AppendToken(target, value);
            return true;
        }

        private static long TokenLength(string value)
        {
            int length = value == null ? 0 : value.Length;
            return length + 1L + length.ToString(CultureInfo.InvariantCulture).Length;
        }

        private static void ReadDynamicArrayMarker(
            FutureMetadataBlock block,
            string metadataName,
            out bool dynamicArray,
            out string failureReason)
        {
            dynamicArray = false;
            failureReason = null;
            if (!string.Equals(metadataName, DynamicArrayMetadataName, StringComparison.Ordinal)) return;

            List<OpenXmlElement> properties = block.Descendants()
                .Where(element =>
                    string.Equals(element.LocalName, "dynamicArrayProperties", StringComparison.Ordinal) &&
                    string.Equals(element.NamespaceUri, DynamicArrayNamespace, StringComparison.Ordinal))
                .ToList();
            if (properties.Count != 1)
            {
                failureReason =
                    "XLDAPR metadata does not contain exactly one dynamicArrayProperties marker.";
                return;
            }

            string rawDynamic = Attribute(properties[0], "fDynamic");
            if (rawDynamic == "1" || string.Equals(rawDynamic, "true", StringComparison.OrdinalIgnoreCase))
            {
                dynamicArray = true;
                return;
            }
            if (rawDynamic == "0" || string.Equals(rawDynamic, "false", StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrEmpty(rawDynamic))
                return;
            failureReason = "The XLDAPR fDynamic marker has an invalid Boolean lexical value.";
        }

        private static string Attribute(OpenXmlElement element, string localName)
        {
            foreach (OpenXmlAttribute attribute in element.GetAttributes())
            {
                if (string.IsNullOrEmpty(attribute.NamespaceUri) &&
                    string.Equals(attribute.LocalName, localName, StringComparison.Ordinal))
                    return attribute.Value ?? "";
            }
            return "";
        }

        private static void AppendToken(StringBuilder target, string value)
        {
            value = value ?? "";
            target.Append(value.Length.ToString(CultureInfo.InvariantCulture));
            target.Append(':');
            target.Append(value);
        }

        private sealed class MetadataTypeEvidence
        {
            internal MetadataTypeEvidence(string name, string exactEvidence)
            {
                Name = name;
                ExactEvidence = exactEvidence;
            }
            internal string Name { get; }
            internal string ExactEvidence { get; }
        }

        private sealed class MetadataRecordEvidence
        {
            internal MetadataRecordEvidence(string rawTypeIndex, string rawValueIndex)
            {
                RawTypeIndex = rawTypeIndex;
                RawValueIndex = rawValueIndex;
            }
            internal string RawTypeIndex { get; }
            internal string RawValueIndex { get; }
        }

        private sealed class CellMetadataEvidence
        {
            internal CellMetadataEvidence(string exactEvidence, IList<MetadataRecordEvidence> records)
            {
                ExactEvidence = exactEvidence;
                Records = records;
            }
            internal string ExactEvidence { get; }
            internal IList<MetadataRecordEvidence> Records { get; }
        }

        private sealed class FutureMetadataEvidence
        {
            internal FutureMetadataEvidence(IList<FutureMetadataBlockEvidence> blocks)
            {
                Blocks = blocks;
            }
            internal IList<FutureMetadataBlockEvidence> Blocks { get; }
        }

        private sealed class FutureMetadataBlockEvidence
        {
            internal FutureMetadataBlockEvidence(
                string exactEvidence,
                bool isDynamicArray,
                string failureReason)
            {
                ExactEvidence = exactEvidence;
                IsDynamicArray = isDynamicArray;
                FailureReason = failureReason;
            }
            internal string ExactEvidence { get; }
            internal bool IsDynamicArray { get; }
            internal string FailureReason { get; }
        }
    }

    internal sealed class OpenXmlFormulaMetadataResolution
    {
        internal static readonly OpenXmlFormulaMetadataResolution Empty =
            new OpenXmlFormulaMetadataResolution(null, false, null, 0);

        internal OpenXmlFormulaMetadataResolution(
            string exactEvidence,
            bool isDynamicArray,
            string failureReason,
            long recordCount)
        {
            ExactEvidence = exactEvidence;
            IsDynamicArray = isDynamicArray;
            FailureReason = failureReason;
            RecordCount = recordCount;
        }

        internal string ExactEvidence { get; }
        internal bool IsDynamicArray { get; }
        internal string FailureReason { get; }
        internal long RecordCount { get; }
    }

    internal static class OpenXmlFormulaReferenceClassifier
    {
        internal static bool HasExternalWorkbookReference(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return false;

            int index = 0;
            while (index < formula.Length)
            {
                if (formula[index] == '"')
                {
                    SkipDoubleQuoted(formula, ref index);
                    continue;
                }

                if (formula[index] == '\'')
                {
                    int contentStart = index + 1;
                    if (!SkipSingleQuoted(formula, ref index, out int contentEnd)) return false;
                    if (index < formula.Length && formula[index] == '!' &&
                        ContainsBracketPair(formula, contentStart, contentEnd))
                        return true;
                    continue;
                }

                if (formula[index] == '[')
                {
                    int close = formula.IndexOf(']', index + 1);
                    if (close >= 0)
                    {
                        int probe = close + 1;
                        while (probe < formula.Length &&
                               formula[probe] != '!' &&
                               formula[probe] != '+' && formula[probe] != '-' &&
                               formula[probe] != '*' && formula[probe] != '/' &&
                               formula[probe] != ',' && formula[probe] != ';' &&
                               formula[probe] != ')' && formula[probe] != '(')
                            probe++;
                        if (probe < formula.Length && formula[probe] == '!') return true;
                        index = close + 1;
                        continue;
                    }
                }

                index++;
            }
            return false;
        }

        private static bool ContainsBracketPair(string text, int start, int end)
        {
            int open = text.IndexOf('[', start, end - start);
            return open >= 0 && text.IndexOf(']', open + 1, end - open - 1) >= 0;
        }

        private static void SkipDoubleQuoted(string text, ref int index)
        {
            index++;
            while (index < text.Length)
            {
                if (text[index++] != '"') continue;
                if (index < text.Length && text[index] == '"')
                {
                    index++;
                    continue;
                }
                return;
            }
        }

        private static bool SkipSingleQuoted(string text, ref int index, out int contentEnd)
        {
            index++;
            while (index < text.Length)
            {
                if (text[index++] != '\'') continue;
                if (index < text.Length && text[index] == '\'')
                {
                    index++;
                    continue;
                }
                contentEnd = index - 1;
                return true;
            }
            contentEnd = text.Length;
            return false;
        }
    }
}
