using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    public sealed class SwpDiffOptions
    {
        public SwpPlantType PlantType = SwpPlantType.Msc;
        public int FirstLineId = 1;
        public int GroupIndex = 1;

        public bool TrustConfigPlant;               // generated/P-type rows write only the label addition
        public bool UpdateDescriptionFormulas;      // wholesale Q/R rewrite (legacy A40 = "y")

        /// <summary>What a FU/Location mismatch does to the writes (the warning always fires,
        /// suppressible per row): append the compacted identity — device blanks, template-only
        /// formulas get the extension — or overwrite with the IO values as-is.</summary>
        public SwpMismatchMode MismatchMode = SwpMismatchMode.AppendToDescription;

        /// <summary>The bus-failure constants: language 1 (the SWP's Description 1) and the
        /// workspace's chosen second language, resolved from the _Languages table.</summary>
        public string BusFailureLang1 = SwpDescriptions.DefaultBusFailureText;
        public string BusFailureLang2 = "";

        /// <summary>The merged + tweaked glossary pairs (null = no glossary in this workspace).</summary>
        public IReadOnlyList<SwpGlossaryPair> GlossaryPairs;
        public bool GlossaryTranslateAlways;
        public IReadOnlyList<SwpDictionaryEntry> Dictionary1;
        public IReadOnlyList<SwpDictionaryEntry> Dictionary2;
        public IReadOnlyDictionary<string, int> ChuteMap;                 // MSC chutes only
        public IReadOnlyDictionary<string, SwpRowOverride> Overrides;    // keyed by mnemonic

        /// <summary>The per-object index rules (_Offsets table; null = the official v93 values).</summary>
        public SwpIndexRules IndexRules;
    }

    /// <summary>One diagnostic mnemonic joined to its SWP row, with everything the spine shows.</summary>
    public sealed class SwpSpineRow
    {
        public DiagnosticEntry Entry;
        public SwpMappedMnemonic Mapped;
        public SwpRowSnapshot Snapshot;        // null unless matched
        public SwpRowVerdict Verdict = SwpRowVerdict.UpToDate;
        public bool FUnitMismatch;
        public bool LocationMismatch;

        /// <summary>This row's effective FLD treatment is the brutal overwrite (H/I cells) rather
        /// than the append — the workspace default, or the row's own [Overwrite FLD] flip.</summary>
        public bool FldOverwriteMode;

        /// <summary>The row's mismatch is marked deliberate ([Mismatch OK] override, or a
        /// generated row): the warning and the orange verdict stay silent; the configured
        /// mismatch treatment still applies.</summary>
        public bool MismatchSuppressed;

        /// <summary>An IO-list alarm ends up with NO T.S. Ref at all — neither the IO list nor
        /// the glossary offers a card and the SWP holds none. Mandatory data: shown red.</summary>
        public bool TsRefMissing;
        public string Desc1Proposed = "";
        public string Desc2Proposed = "";
        public List<SwpProposedEdit> Edits = new List<SwpProposedEdit>();
        public SwpCsvVerdict CsvVerdict = SwpCsvVerdict.NotChecked;
        public SwpBusfieldVerdict BusfieldVerdict = SwpBusfieldVerdict.NotChecked;
        public SwpAlarmVerdict AlarmVerdict = SwpAlarmVerdict.NotChecked;
        public SwpBusExtVerdict BusExtVerdict = SwpBusExtVerdict.NotChecked;

        /// <summary>Where each coherence verdict came from — the 1-based line of the checked
        /// document (0 = nothing matched), so a double-click can open the source right there.</summary>
        public int CsvLine;
        public int BusfieldLine;
        public int BusExtLine;
        public int AlarmLine;
        public string AlarmSource = "";   // the Alarms_*.txt file name the match came from

        /// <summary>The phrase this row keys the glossary with (full permanent part).</summary>
        public string GlossaryKey = "";
    }

    public sealed class SwpSheetTally
    {
        public string SheetName = "";
        public int UpToDate, Pending, LabelMismatch, RowMissing, Duplicate, Invalid, SheetMissing;
        public int Total => UpToDate + Pending + LabelMismatch + RowMissing + Duplicate + Invalid + SheetMissing;
    }

    public sealed class SwpWorkspaceModel
    {
        public List<SwpSpineRow> Rows = new List<SwpSpineRow>();
        public List<SwpProposedEdit> Edits = new List<SwpProposedEdit>();
        public List<SwpSheetTally> Tallies = new List<SwpSheetTally>();
        public List<string> MissingSheets = new List<string>();
    }

    /// <summary>
    /// The join + proposal engine: each filtered mnemonic maps to its SWP coordinate, joins the
    /// row snapshot, and yields the per-cell edits the legacy update would have written — but as
    /// a reviewable diff (already-correct cells produce nothing). The FUnit/Location mismatch
    /// rules, TrustConfigPlant and the wholesale formula rewrite are ported verbatim from
    /// a_EditSWP; duplicate targets are surfaced instead of silently last-wins overwritten.
    /// </summary>
    public static class SwpDiff
    {
        public static SwpWorkspaceModel Build(IReadOnlyList<DiagnosticEntry> entries,
            SwpWorkbookIndex swp, SwpDiffOptions options, List<SwpLogEntry> log)
        {
            if (swp == null) throw new ArgumentNullException(nameof(swp));
            options = options ?? new SwpDiffOptions();
            var model = new SwpWorkspaceModel();
            model.MissingSheets.AddRange(swp.MissingSheets);

            var composeContext = new SwpDescriptions.Context
            {
                Dictionary1 = options.Dictionary1,
                Dictionary2 = options.Dictionary2,
                BusFailureLang1 = options.BusFailureLang1,
                BusFailureLang2 = options.BusFailureLang2,
                GlossaryPairs = options.GlossaryPairs,
                GlossaryTranslateAlways = options.GlossaryTranslateAlways,
            };

            var dashLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (DiagnosticEntry entry in entries ?? new List<DiagnosticEntry>())
            {
                if (entry == null) continue;
                var row = new SwpSpineRow { Entry = entry };
                model.Rows.Add(row);

                SwpRowOverride overrides = OverrideFor(options, entry.Mnemonic);

                row.Mapped = SwpAddressMap.Map(entry.Mnemonic, options.PlantType,
                    options.FirstLineId, options.GroupIndex,
                    overrides?.SwpDbIndex, options.ChuteMap, options.IndexRules);

                // Official-rules sanity on the IO side: the last bit of a DWORD is forbidden
                // (§3.3), and location names must not contain '-' (§3.8, once per location).
                if (row.Mapped.Ok &&
                    (row.Mapped.Coordinate.Bit == 31 || row.Mapped.Coordinate.Bit == 63))
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "bit " + row.Mapped.Coordinate.Bit + " is forbidden — the last bit of a " +
                        "DWORD must stay unused (rules §3.3)",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                if (entry.Location.IndexOf('-') >= 0 && dashLocations.Add(entry.Location))
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "location \"" + entry.Location + "\" contains '-' — forbidden in " +
                        "location names (rules §3.8)",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));

                SwpDescriptions.Result descriptions = SwpDescriptions.Compose(entry, composeContext, overrides);
                row.Desc1Proposed = descriptions.Desc1;
                row.Desc2Proposed = descriptions.Desc2;
                row.GlossaryKey = descriptions.GlossaryKey;
                if (descriptions.GlossaryDiffers)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the IO list's language 2 differs from the glossary: \"" +
                        descriptions.IoLang2Text + "\" vs \"" + descriptions.GlossaryText + "\"",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                if (descriptions.BusDescMismatch)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the bus module's description \"" + descriptions.BusIoText +
                        "\" is not an official glossary phrase — writing \"" +
                        descriptions.BusChosenEn + "\" instead",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                if (descriptions.TsRefOverridden)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the IO list's T.S. Ref \"" + descriptions.TsRefFromIo +
                        "\" differs from the glossary's \"" + descriptions.TsRef +
                        "\" for this phrase — the glossary wins",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));

                if (!row.Mapped.Ok)
                {
                    row.Verdict = SwpRowVerdict.Invalid;
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "cannot resolve an SWP coordinate: " + row.Mapped.Error,
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                    continue;
                }

                SwpCoordinate coordinate = row.Mapped.Coordinate;
                if (!swp.Sheets.TryGetValue(coordinate.SheetName, out SwpSheetIndex sheet))
                {
                    row.Verdict = SwpRowVerdict.SheetMissing;
                    continue;
                }
                if (!sheet.RowsByKey.TryGetValue(coordinate.RowKey, out SwpRowSnapshot snapshot))
                {
                    row.Verdict = SwpRowVerdict.RowMissing;
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                        "tag not found in the SWP: " + coordinate.FullKey,
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                    continue;
                }

                row.Snapshot = snapshot;
                ProposeEdits(row, coordinate, snapshot, descriptions, overrides, options, log);
                row.Verdict = (row.FUnitMismatch || row.LocationMismatch) && !row.MismatchSuppressed
                    ? SwpRowVerdict.LabelMismatch
                    : row.Edits.Count > 0 ? SwpRowVerdict.Pending : SwpRowVerdict.UpToDate;
            }

            DemoteDuplicateTargets(model, log);

            foreach (SwpSpineRow row in model.Rows)
                model.Edits.AddRange(row.Edits);
            Tally(model);
            return model;
        }

        private static void ProposeEdits(SwpSpineRow row, SwpCoordinate coordinate,
            SwpRowSnapshot snapshot, SwpDescriptions.Result descriptions, SwpRowOverride overrides,
            SwpDiffOptions options, List<SwpLogEntry> log)
        {
            DiagnosticEntry entry = row.Entry;
            // The mismatch is DELIBERATE for rows flagged [Mismatch OK] and for generated rows
            // (photocells, strobes — their labels never match the block): the warning and the
            // orange verdict stay silent, while the configured treatment still applies.
            bool deliberate = (overrides != null && overrides.SkipLabelCheck) ||
                entry.GeneratedBy.Length > 0;
            bool forceElectricalIds = overrides != null && overrides.ForceElectricalIds;
            bool isBeltSheet = coordinate.SheetName.EndsWith("BELT", StringComparison.Ordinal);

            // The FLD rule: the IO list's Functional unit / Location / Device must ALWAYS be
            // visible in the SWP — either in the row's own cells (H/I/J, "brutal overwrite") or
            // appended to the descriptions. Append mode compares against the BLOCK header (the
            // row's H/I are the header's formulas; broken ones get restored first); overwrite
            // mode compares against the row's OWN cells — after an overwrite they hold the IO
            // identity and the row reads up to date.
            // The treatment is per-row: [Overwrite FLD] / [Append FLD] flips THIS row away from
            // the workspace default. The FLD still always lands somewhere — only where changes.
            SwpMismatchMode fldMode = overrides?.FldMode ?? options.MismatchMode;
            bool overwriteMode = fldMode == SwpMismatchMode.Overwrite;
            row.FldOverwriteMode = overwriteMode;
            string swpFUnit = overwriteMode && snapshot.HasCellH ? snapshot.CachedH : snapshot.BlockFUnit;
            string swpLocation = overwriteMode && snapshot.HasCellI ? snapshot.CachedI : snapshot.BlockLocation;
            row.FUnitMismatch = !SameTrimmed(swpFUnit, entry.FUnit);
            row.LocationMismatch = !SameTrimmed(swpLocation, entry.Location);
            row.MismatchSuppressed = deliberate;
            bool mismatch = row.FUnitMismatch || row.LocationMismatch;

            // Brutal overwrite puts the identity in the row's own cells — unless a cell is a
            // shared formula (not writable headlessly): that row falls back to the append, so
            // the FLD is never lost.
            bool overwriteCells = mismatch && overwriteMode;
            if (overwriteCells &&
                ((row.FUnitMismatch && snapshot.FormulaHShared) || (row.LocationMismatch && snapshot.FormulaIShared)))
            {
                overwriteCells = false;
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the row's FU/Location cells are shared formulas — not overwritable headlessly; " +
                    "the identity is appended to the descriptions instead",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            }
            bool appendIdentity = mismatch && !overwriteCells;

            if (mismatch && !deliberate)
            {
                // The wording follows the row's H/I formula health: intact formulas mean a REAL
                // configuration difference; broken ones already have a restoration queued (the
                // formula cleaner), and the comparison stands on the restored block values.
                bool healthy =
                    SwpFormulaCheck.RowFormulaIsHealthy(snapshot.FormulaH, snapshot.FormulaHShared,
                        snapshot.BlockHeaderRow, "H") &&
                    SwpFormulaCheck.RowFormulaIsHealthy(snapshot.FormulaI, snapshot.FormulaIShared,
                        snapshot.BlockHeaderRow, "I");
                string what = row.FUnitMismatch && row.LocationMismatch ? "functional unit and location"
                    : row.FUnitMismatch ? "functional unit" : "location";
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    what + " differ from the block at " + coordinate.FullKey +
                    " [SWP \"" + snapshot.BlockFUnit + "\" \"" + snapshot.BlockLocation +
                    "\" | IO \"" + entry.FUnit + "\" \"" + entry.Location + "\"]" +
                    (healthy
                        ? " — deliberate? mark [Mismatch OK] on the row"
                        : " — the row's FU/Location formulas are broken (restoration queued); " +
                          "compared against the restored block values"),
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            }

            string additionToDescription = "";
            string deviceValue = entry.Device;
            if (appendIdentity)
            {
                additionToDescription = " " + (entry.FUnit + entry.Location + entry.Device).Replace(" ", "");
                deviceValue = "";
            }

            // TrustConfigPlant: generated alarms (and "P*"-typed module rows) trust the template
            // texts already in CONFIG_PLANT — only the label addition is written into T/U.
            // "Never overwrite CONFIG_PLANT descriptions": the row's final SCADA text is
            // CONFIG's template ⊕ T, so writing the composed IO text into T stacks a second
            // sentence onto a template that already says it. Trusted rows contribute the FLD
            // label alone (and nothing at all when there is no mismatch — see the guard below).
            // A per-row tick answers outright; only the workspace default carries the automatic
            // scope (rows this bridge generated + the "P*"-typed module rows).
            bool trustTemplates = overrides?.TrustConfigDescriptions ??
                (options.TrustConfigPlant &&
                 (entry.GeneratedBy.Length > 0 || entry.TypeCell.StartsWith("P", StringComparison.Ordinal)));
            string desc1Value = trustTemplates ? additionToDescription : descriptions.Desc1 + additionToDescription;
            string desc2Value = trustTemplates ? additionToDescription : descriptions.Desc2 + additionToDescription;

            void Edit(SwpField field, string current, string proposed)
            {
                row.Edits.Add(new SwpProposedEdit
                {
                    Sheet = coordinate.SheetName,
                    CoordinateKey = coordinate.RowKey,
                    Field = field,
                    Mnemonic = entry.Mnemonic,
                    Current = current ?? "",
                    Proposed = proposed ?? "",
                    RowNumber = snapshot.RowNumber,
                });
            }

            // Rules §3.1 writes the export flag as "X"; the comparison stays case-insensitive so
            // the fleet's legacy lowercase "x" never churns.
            if (!SwpFieldCompare.Same(SwpField.ExportB, snapshot.ExportB, "X"))
                Edit(SwpField.ExportB, snapshot.ExportB, "X");

            // An addressless entry (a virtual CLASSDIAG row) never touches G — proposing a blank
            // would erase whatever the SWP already holds.
            string swpAddress = entry.SwpAddress();
            if (swpAddress.Length > 0 && !SwpFieldCompare.Same(SwpField.AddressG, snapshot.AddressG, swpAddress))
                Edit(SwpField.AddressG, snapshot.AddressG, swpAddress);

            // Belt alarms keep their SWP device id (a belt has no single electrical device)
            // unless the per-row force flag insists — the legacy rule verbatim.
            if ((forceElectricalIds || !isBeltSheet) &&
                !SwpFieldCompare.Same(SwpField.DeviceJ, snapshot.DeviceJ, deviceValue))
                Edit(SwpField.DeviceJ, snapshot.DeviceJ, deviceValue);

            // Brutal overwrite: the IO identity replaces the block formula in the row's own
            // H/I cells (plain values — the formula cleaner reads them as deliberate).
            if (overwriteCells)
            {
                if (row.FUnitMismatch && !SwpFieldCompare.Same(SwpField.FUnitH, snapshot.CachedH, entry.FUnit))
                    Edit(SwpField.FUnitH, snapshot.CachedH, entry.FUnit);
                if (row.LocationMismatch && !SwpFieldCompare.Same(SwpField.LocationI, snapshot.CachedI, entry.Location))
                    Edit(SwpField.LocationI, snapshot.CachedI, entry.Location);
            }

            // T.S. Ref: the SCADA launches the troubleshooting manual straight from this code,
            // so an empty composed value NEVER blanks an existing one — a wrong section beats
            // no action. And IO-list alarms MUST carry one: a row left with none anywhere
            // (IO list, glossary card, SWP) is flagged red on the workspace.
            string tsRef = descriptions.TsRef.Trim();
            if (tsRef.Length > 0 &&
                !SwpFieldCompare.Same(SwpField.TsRefL, snapshot.TsRefL, descriptions.TsRef))
                Edit(SwpField.TsRefL, snapshot.TsRefL, descriptions.TsRef);
            row.TsRefMissing = entry.GeneratedBy.Length == 0 &&
                tsRef.Length == 0 && snapshot.TsRefL.Trim().Length == 0;
            if (row.TsRefMissing)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "no T.S. Ref anywhere — mandatory for IO-list alarms; fill it in the IO " +
                    "list, or add the phrase's card to the glossary",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));

            // Trust rows carry their texts in the CONFIG templates: with no label addition
            // there is nothing to write — never propose blanking what the templates feed.
            if (!(trustTemplates && additionToDescription.Trim().Length == 0))
            {
                if (!SwpFieldCompare.Same(SwpField.Desc1T, snapshot.Desc1T, desc1Value))
                    Edit(SwpField.Desc1T, snapshot.Desc1T, desc1Value);
                if (!SwpFieldCompare.Same(SwpField.Desc2U, snapshot.Desc2U, desc2Value))
                    Edit(SwpField.Desc2U, snapshot.Desc2U, desc2Value);
            }

            bool appendPending = appendIdentity;
            ProposeFormulaEdit(row, coordinate, snapshot, options, appendPending, SwpField.FormulaQ,
                snapshot.FormulaQ, snapshot.FormulaQShared, "T", log);
            ProposeFormulaEdit(row, coordinate, snapshot, options, appendPending, SwpField.FormulaR,
                snapshot.FormulaR, snapshot.FormulaRShared, "U", log);
        }

        private static void ProposeFormulaEdit(SwpSpineRow row, SwpCoordinate coordinate,
            SwpRowSnapshot snapshot, SwpDiffOptions options, bool appendPending, SwpField field,
            string formula, bool isShared, string descriptionColumn, List<SwpLogEntry> log)
        {
            string proposed = null;

            if (options.UpdateDescriptionFormulas)
                proposed = WholesaleDescriptionFormula(options.PlantType, coordinate.SheetName,
                    snapshot.BlockHeaderRow, snapshot.RowNumber, descriptionColumn);
            else if (appendPending &&
                     formula.Length > 0 && formula.EndsWith(")", StringComparison.Ordinal) &&
                     formula.Replace("$", "").IndexOf(
                         descriptionColumn + snapshot.RowNumber.ToString(CultureInfo.InvariantCulture),
                         StringComparison.OrdinalIgnoreCase) < 0)
            {
                // In APPEND mode the extension is automatic for TEMPLATE-ONLY formulas — the
                // only way the appended identity reaches the displayed text. A formula already
                // referencing T{row}/U{row} (the standard SUBSTITUTE(template&T,…) shape) flows
                // the addition by itself — extending it would print the description twice.
                proposed = formula + "&\" \"&" + descriptionColumn + snapshot.RowNumber;
            }

            if (proposed == null || proposed == formula) return;

            if (isShared)
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the " + descriptionColumn + "-description formula at " + coordinate.FullKey +
                    " is a shared formula — left untouched",
                    row.Entry.CardSheet, row.Entry.SourceRow, row.Entry.Mnemonic));
                return;
            }

            row.Edits.Add(new SwpProposedEdit
            {
                Sheet = coordinate.SheetName,
                CoordinateKey = coordinate.RowKey,
                Field = field,
                Mnemonic = row.Entry.Mnemonic,
                Current = formula,
                Proposed = proposed,
                RowNumber = snapshot.RowNumber,
            });
        }

        /// <summary>
        /// The legacy MakeAlarmDescriptionFormula: a CONCAT over the sheet description (B6), the
        /// block's electrical label and the T/U cell. The label row is the enclosing block header
        /// (the nearest row above whose F holds a real label — which is exactly that row).
        /// </summary>
        internal static string WholesaleDescriptionFormula(SwpPlantType plantType, string sheetName,
            int labelRow, int rowNumber, string descriptionColumn)
        {
            string cell = descriptionColumn + rowNumber.ToString(CultureInfo.InvariantCulture);
            string label = "$F" + labelRow.ToString(CultureInfo.InvariantCulture);
            if (plantType == SwpPlantType.Ups)
            {
                int underscore = sheetName.IndexOf('_');
                string sheetType = underscore >= 0 ? sheetName.Substring(underscore + 1) : sheetName;
                return "CONCAT(\"UPSTREAM \",$B$6,\" - " + sheetType + " \"," + label + ",\" - \"," + cell + ")";
            }
            return "CONCAT(UPPER($B$6),\" \"," + label + ",\" - \"," + cell + ")";
        }

        /// <summary>
        /// What the Q/R cell will DISPLAY after Excel recalculates, simulated from the cached
        /// text. A formula REFERENCING the description cell displays template &amp; desc — the
        /// cached final ends with the CURRENT desc and the simulation swaps that suffix for the
        /// proposed one (never concatenating old and new). A TEMPLATE-ONLY formula shows the
        /// description only through a pending append ( &amp;" "&amp;T ) — exactly once; without
        /// one, a description edit does not change the final text at all. A pending wholesale
        /// rewrite composes from scratch. Both field bugs of 2026-08-21 (the old-plus-new mix,
        /// then the doubled append) are pinned by tests.
        /// </summary>
        public static string SimulateRecalculatedFinal(string cachedFinal, string currentDesc,
            string proposedDesc, bool formulaReferencesDesc, bool formulaAppendPending,
            bool wholesalePending, SwpPlantType plantType, string sheetName,
            string sheetDescription, string blockLabel)
        {
            string desc = proposedDesc ?? "";

            if (wholesalePending)
            {
                if (plantType == SwpPlantType.Ups)
                {
                    int underscore = (sheetName ?? "").IndexOf('_');
                    string sheetType = underscore >= 0 ? sheetName.Substring(underscore + 1) : sheetName ?? "";
                    return "UPSTREAM " + sheetDescription + " - " + sheetType + " " + blockLabel + " - " + desc;
                }
                return (sheetDescription ?? "").ToUpperInvariant() + " " + blockLabel + " - " + desc;
            }

            string cached = cachedFinal ?? "";
            if (formulaReferencesDesc)
            {
                string current = currentDesc ?? "";
                return current.Length > 0 && cached.EndsWith(current, StringComparison.Ordinal)
                    ? cached.Substring(0, cached.Length - current.Length) + desc
                    : cached + desc;   // desc was empty so far: template & "" displayed the prefix alone
            }
            return formulaAppendPending ? cached + " " + desc : cached;
        }

        /// <summary>
        /// Rows resolving to the same SWP row would overwrite each other (the legacy tool silently
        /// let the last one win). All of them are demoted and none of their edits survive.
        /// </summary>
        private static void DemoteDuplicateTargets(SwpWorkspaceModel model, List<SwpLogEntry> log)
        {
            var bySwpRow = new Dictionary<string, List<SwpSpineRow>>();
            foreach (SwpSpineRow row in model.Rows)
            {
                if (row.Snapshot == null || !row.Mapped.Ok) continue;
                string key = row.Mapped.Coordinate.FullKey;
                if (!bySwpRow.TryGetValue(key, out List<SwpSpineRow> list))
                    bySwpRow[key] = list = new List<SwpSpineRow>();
                list.Add(row);
            }

            foreach (var pair in bySwpRow)
            {
                if (pair.Value.Count < 2) continue;
                var sources = new List<string>();
                foreach (SwpSpineRow row in pair.Value)
                    sources.Add(row.Entry.Mnemonic + " (" + row.Entry.CardSheet + " row " + row.Entry.SourceRow + ")");
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Error,
                    pair.Value.Count + " rows target " + pair.Key + ": " + string.Join(", ", sources) +
                    " — nothing is queued for them"));
                foreach (SwpSpineRow row in pair.Value)
                {
                    row.Verdict = SwpRowVerdict.DuplicateTarget;
                    row.Edits.Clear();
                }
            }
        }

        private static void Tally(SwpWorkspaceModel model)
        {
            var bySheet = new Dictionary<string, SwpSheetTally>(StringComparer.OrdinalIgnoreCase);
            SwpSheetTally For(string name)
            {
                if (!bySheet.TryGetValue(name, out SwpSheetTally tally))
                {
                    bySheet[name] = tally = new SwpSheetTally { SheetName = name };
                    model.Tallies.Add(tally);
                }
                return tally;
            }

            foreach (SwpSpineRow row in model.Rows)
            {
                string sheet = row.Mapped != null && row.Mapped.Ok ? row.Mapped.Coordinate.SheetName : "(unmapped)";
                SwpSheetTally tally = For(sheet);
                switch (row.Verdict)
                {
                    case SwpRowVerdict.UpToDate: tally.UpToDate++; break;
                    case SwpRowVerdict.Pending: tally.Pending++; break;
                    case SwpRowVerdict.LabelMismatch: tally.LabelMismatch++; break;
                    case SwpRowVerdict.RowMissing: tally.RowMissing++; break;
                    case SwpRowVerdict.DuplicateTarget: tally.Duplicate++; break;
                    case SwpRowVerdict.Invalid: tally.Invalid++; break;
                    case SwpRowVerdict.SheetMissing: tally.SheetMissing++; break;
                }
            }
            model.Tallies.Sort((a, b) => string.CompareOrdinal(a.SheetName, b.SheetName));
        }

        private static SwpRowOverride OverrideFor(SwpDiffOptions options, string mnemonic)
        {
            if (options.Overrides == null) return null;
            return options.Overrides.TryGetValue(mnemonic, out SwpRowOverride result) ? result : null;
        }

        private static bool SameTrimmed(string a, string b, bool caseInsensitive = false)
        {
            a = (a ?? "").Trim();
            b = (b ?? "").Trim();
            return string.Equals(a, b, caseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }
    }
}
