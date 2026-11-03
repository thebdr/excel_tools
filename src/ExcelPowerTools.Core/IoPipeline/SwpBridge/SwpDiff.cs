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

        /// <summary>
        /// What a row does with CONFIG_PLANT's template text. KeepConfigText leaves the Q/R
        /// formula alone and writes only the FLD label into T/U (a cell that formula already shows
        /// after a template with no words past its placeholders still takes the description — it
        /// is the alarm's only name); ReplaceTemplateTail rebuilds the
        /// formula so the row's description lands where the template's generic tail was (the
        /// legacy A40 = "y") instead of stacking after it. Exclusive by nature — a row cannot do
        /// both — which is why it is one mode and not two switches.
        /// </summary>
        public SwpConfigMode ConfigMode = SwpConfigMode.ReplaceTemplateTail;

        /// <summary>The destination-sheet exception to <see cref="ConfigMode"/>: rows whose SWP
        /// sheet name matches take the rule's mode instead (null = no exception). A per-row
        /// [Keep CONFIG text] / [Write IO text] tick still wins over it.</summary>
        public SwpConfigSheetRule ConfigSheetRule;

        /// <summary>What a FU/Location mismatch does to the writes (the warning always fires,
        /// suppressible per row): append the compacted identity — device blanks, template-only
        /// formulas get the extension — or overwrite with the IO values as-is.</summary>
        public SwpMismatchMode MismatchMode = SwpMismatchMode.AppendToDescription;

        /// <summary>The destination-sheet exception to <see cref="MismatchMode"/>: rows whose SWP
        /// sheet name matches the rule take the rule's mode instead (null = no exception). A
        /// per-row [Append FLD] / [Overwrite FLD] tick still wins over it.</summary>
        public SwpFldSheetRule FldSheetRule;

        /// <summary>The bus-failure constants: language 1 (the SWP's Description 1) and the
        /// workspace's chosen second language, resolved from the _Languages table.</summary>
        public string BusFailureLang1 = SwpDescriptions.DefaultBusFailureText;
        public string BusFailureLang2 = "";

        /// <summary>The workspace's second-language code ("FR"; empty = none configured). With one
        /// configured — other than EN, language 1 itself — a Description 2 falling back to language-1
        /// text is reported: on a SCADA that offers only that language, the fallback is text the
        /// operator cannot read.</summary>
        public string SecondLanguage = "";

        /// <summary>The merged + tweaked glossary pairs (null = no glossary in this workspace).</summary>
        public IReadOnlyList<SwpGlossaryPair> GlossaryPairs;
        public bool GlossaryTranslateAlways;
        public IReadOnlyList<SwpDictionaryEntry> Dictionary1;
        public IReadOnlyList<SwpDictionaryEntry> Dictionary2;
        public IReadOnlyDictionary<string, int> ChuteMap;                 // MSC chutes only
        public IReadOnlyDictionary<string, SwpRowOverride> Overrides;    // keyed by mnemonic

        /// <summary>The per-object index rules (_Offsets table; null = the official v93 values).</summary>
        public SwpIndexRules IndexRules;

        /// <summary>The mnemonic rule table — what a mnemonic IS and where it lands
        /// (null = the shipped default).</summary>
        public SwpMnemonicMap MnemonicMap;
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
        /// than the append — the workspace default, the destination-sheet rule, or the row's own
        /// [Overwrite FLD] flip.</summary>
        public bool FldOverwriteMode;

        /// <summary>This row's treatment came from the workspace's destination-sheet rule, not from
        /// the default and not from a per-row tick — what makes the rule's reach reportable.</summary>
        public bool FldModeFromSheetRule;

        /// <summary>The same, for the CONFIG_PLANT destination-sheet rule.</summary>
        public bool ConfigModeFromSheetRule;

        /// <summary>The row's mismatch is marked deliberate ([Mismatch OK] override, or a
        /// generated row): the warning and the orange verdict stay silent; the configured
        /// mismatch treatment still applies.</summary>
        public bool MismatchSuppressed;

        /// <summary>An IO-list alarm ends up with NO T.S. Ref at all — neither the IO list nor
        /// the glossary offers a card and the SWP holds none. Mandatory data: shown red.</summary>
        public bool TsRefMissing;

        /// <summary>The workspace replaces the template's trailing text, but this row was excepted
        /// — by the destination-sheet rule or by its own [Keep CONFIG text] tick — and keeps
        /// CONFIG_PLANT's formula, because a rebuild cuts off the text it was told to keep.</summary>
        public bool ConfigFormulaKept;

        /// <summary>A kept CONFIG_PLANT formula already shows T/U after a template with nothing past
        /// its last placeholder, so the description still lands in that cell (the formula stays).</summary>
        public bool KeptFormulaShowsDescription;
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
    /// rules are ported verbatim from a_EditSWP; duplicate targets are surfaced instead of silently last-wins overwritten.
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

            // A destination-sheet rule that does not compile is inert, never fatal — but it is
            // announced, because silently falling back to the default is how a whole sheet ends
            // up with the treatment the workspace did not ask for.
            if (options.FldSheetRule != null && options.FldSheetRule.Error.Length > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the destination-sheet FLD rule /" + options.FldSheetRule.Pattern +
                    "/ does not parse (" + options.FldSheetRule.Error +
                    ") — every sheet follows the workspace default instead"));
            if (options.ConfigSheetRule != null && options.ConfigSheetRule.Error.Length > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the destination-sheet CONFIG_PLANT rule /" + options.ConfigSheetRule.Pattern +
                    "/ does not parse (" + options.ConfigSheetRule.Error +
                    ") — every sheet follows the workspace default instead"));

            var composeContext = new SwpDescriptions.Context
            {
                Dictionary1 = options.Dictionary1,
                Dictionary2 = options.Dictionary2,
                BusFailureLang1 = options.BusFailureLang1,
                BusFailureLang2 = options.BusFailureLang2,
                GlossaryPairs = options.GlossaryPairs,
                GlossaryTranslateAlways = options.GlossaryTranslateAlways,
            };

            // Resolved once: compiling the shipped table per row would dominate a 30k-row build.
            SwpMnemonicMap mnemonicMap = options.MnemonicMap ?? SwpMnemonicMapDefaults.Official();
            var indexSources = new Dictionary<string, IReadOnlyDictionary<string, int>>(
                StringComparer.OrdinalIgnoreCase);
            if (options.ChuteMap != null) indexSources["Destination_Parameters"] = options.ChuteMap;

            var dashLocations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (DiagnosticEntry entry in entries ?? new List<DiagnosticEntry>())
            {
                if (entry == null) continue;
                var row = new SwpSpineRow { Entry = entry };
                model.Rows.Add(row);

                SwpRowOverride overrides = OverrideFor(options, entry.Mnemonic);

                row.Mapped = mnemonicMap.Map(entry.Mnemonic, options.PlantType, new SwpMapContext
                {
                    FirstLineId = options.FirstLineId,
                    GroupIndex = options.GroupIndex,
                    DeviceIndexOverride = overrides?.SwpDbIndex,
                    Rules = options.IndexRules,
                    Sources = indexSources,
                });

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
                ProposeEdits(row, coordinate, snapshot, sheet.Description, descriptions, overrides, options, log);
                row.Verdict = (row.FUnitMismatch || row.LocationMismatch) && !row.MismatchSuppressed
                    ? SwpRowVerdict.LabelMismatch
                    : row.Edits.Count > 0 ? SwpRowVerdict.Pending : SwpRowVerdict.UpToDate;
            }

            DemoteDuplicateTargets(model, log);

            int configFormulasKept = 0;
            int keptFormulasShowingDescription = 0;
            int sheetRuleRows = 0;
            int configSheetRuleRows = 0;
            foreach (SwpSpineRow row in model.Rows)
            {
                model.Edits.AddRange(row.Edits);
                if (row.ConfigFormulaKept) configFormulasKept++;
                if (row.KeptFormulaShowsDescription) keptFormulasShowingDescription++;
                if (row.FldModeFromSheetRule) sheetRuleRows++;
                if (row.ConfigModeFromSheetRule) configSheetRuleRows++;
            }

            // Like the FLD one, the CONFIG rule moves rows with no visible gesture ON them, so
            // its reach is always reported.
            SwpConfigSheetRule configRule = options.ConfigSheetRule;
            if (configRule != null && configRule.IsActive)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    configSheetRuleRows + " row(s) follow the destination-sheet CONFIG_PLANT rule /" +
                    configRule.Pattern + "/ — on the SWP sheets it matches, rows " +
                    (configRule.Mode == SwpConfigMode.KeepConfigText
                        ? "keep CONFIG_PLANT's text (only the FLD label is written)"
                        : "have the template's trailing text replaced by their description") +
                    " instead of the workspace default"));
            // The sheet rule moves rows without any visible gesture on them, so its reach is
            // always reported — silence would hide which sheets took the other treatment.
            SwpFldSheetRule fldRule = options.FldSheetRule;
            if (fldRule != null && fldRule.IsActive)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    sheetRuleRows + " row(s) follow the destination-sheet FLD rule /" + fldRule.Pattern +
                    "/ — on the SWP sheets it matches, a FU/Location mismatch " +
                    (fldRule.Mode == SwpMismatchMode.Overwrite
                        ? "overwrites the row's FU/Loc/Device cells"
                        : "appends the identity to the descriptions") +
                    " instead of the workspace default"));
            if (configFormulasKept > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    configFormulasKept + " row(s) keep their CONFIG_PLANT description formulas while the " +
                    "workspace replaces the template's trailing text — the destination-sheet rule and the " +
                    "per-row [Keep CONFIG text] ticks are what excepted them"));
            if (keptFormulasShowingDescription > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    keptFormulasShowingDescription + " row(s) keep CONFIG_PLANT's formula and still receive their " +
                    "description: the formula already shows T/U after a template with no text past its last " +
                    "placeholder, so that cell is the alarm's only name"));
            Tally(model);
            return model;
        }

        private static void ProposeEdits(SwpSpineRow row, SwpCoordinate coordinate,
            SwpRowSnapshot snapshot, string sheetDescription, SwpDescriptions.Result descriptions,
            SwpRowOverride overrides, SwpDiffOptions options, List<SwpLogEntry> log)
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
            // The treatment resolves in three steps, most specific first: the per-row
            // [Overwrite FLD] / [Append FLD] tick, then the workspace's destination-sheet rule
            // (a regex on the SWP sheet name), then the workspace default. The FLD still always
            // lands somewhere — only where changes.
            // An EMPTY IO-side member (FU, Location or Device) is missing data, not a value: it
            // can neither match nor mismatch, and it never writes — the SWP's own member stays,
            // the rule G and L already follow (an empty value never erases one). A generated
            // strobe row, or a photocell whose PHT row was not found, carries no identity at all
            // and leaves the SWP row's identity alone; an IO-list row with none is reported,
            // because its FLD cannot be carried anywhere.
            string ioFUnit = (entry.FUnit ?? "").Trim();
            string ioLocation = (entry.Location ?? "").Trim();
            string ioDevice = (entry.Device ?? "").Trim();
            bool hasIdentity = ioFUnit.Length > 0 || ioLocation.Length > 0 || ioDevice.Length > 0;
            if (!hasIdentity && entry.GeneratedBy.Length == 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the IO list carries no functional unit, location or device for this row — " +
                    "nothing to carry into the SWP; the row's own identity is kept untouched",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));

            SwpMismatchMode? sheetRuleMode = options.FldSheetRule?.ModeFor(coordinate.SheetName);
            bool rowChoseMode = overrides != null && overrides.FldMode.HasValue;
            SwpMismatchMode fldMode = overrides?.FldMode ?? sheetRuleMode ?? options.MismatchMode;
            bool overwriteMode = fldMode == SwpMismatchMode.Overwrite;
            row.FldOverwriteMode = overwriteMode;
            row.FldModeFromSheetRule = !rowChoseMode && sheetRuleMode.HasValue;
            string swpFUnit = overwriteMode && snapshot.HasCellH ? snapshot.CachedH : snapshot.BlockFUnit;
            string swpLocation = overwriteMode && snapshot.HasCellI ? snapshot.CachedI : snapshot.BlockLocation;
            row.FUnitMismatch = ioFUnit.Length > 0 && !SameTrimmed(swpFUnit, ioFUnit);
            row.LocationMismatch = ioLocation.Length > 0 && !SameTrimmed(swpLocation, ioLocation);
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
            string deviceValue = ioDevice;
            if (appendIdentity)
            {
                additionToDescription = " " + (ioFUnit + ioLocation + ioDevice).Replace(" ", "");
                deviceValue = "";
            }

            // The CONFIG_PLANT treatment, resolved most-specific-first like the FLD one: the
            // per-row tick, then the destination-sheet rule, then the workspace default.
            //
            // "Never overwrite CONFIG_PLANT descriptions" — the row's final SCADA text is CONFIG's
            // template ⊕ T, so writing the composed IO text into T stacks a second sentence onto a
            // template that already says it. Kept rows contribute the FLD label alone (and nothing
            // at all when there is no mismatch — see the guard below).
            //
            // The mode is the WHOLE answer — no invisible scope. The old setting exempted the rows
            // this bridge generated and the "P*"-typed module rows, because composing a
            // description into T/U stacked it after a template that already said the same thing.
            // ReplaceTemplateTail dissolves that: the description lands WHERE the tail was, so
            // there is nothing left to duplicate and nothing left to exempt. What departs from
            // the default is what a person can see — the destination-sheet rule and the per-row
            // tick, in that order of specificity.
            SwpConfigMode? configSheetMode = options.ConfigSheetRule?.ModeFor(coordinate.SheetName);
            SwpConfigMode configMode = configSheetMode ?? options.ConfigMode;
            row.ConfigModeFromSheetRule = overrides?.TrustConfigDescriptions == null && configSheetMode.HasValue;
            bool trustTemplates = overrides?.TrustConfigDescriptions ?? configMode == SwpConfigMode.KeepConfigText;
            bool rebuildFormulas = !trustTemplates;

            // Keeping CONFIG_PLANT's text keeps its WORDS — it never starves the cell a kept formula
            // shows. Where the formula already displays T/U after a template that says nothing past
            // its last placeholder (CABINET: "INTRODUCTION %1 - ARMOIRE DE GROUPE %2 - " & " " & U),
            // there is no CONFIG sentence to stack onto: that cell is the alarm's only name, and
            // withholding the description leaves whatever an older tool put there — the English of a
            // French-only SCADA (FVZ, 2026-09-10). The formula stays untouched; the text lands.
            string rowNumber = snapshot.RowNumber.ToString(CultureInfo.InvariantCulture);
            bool desc1Shown = trustTemplates && SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                snapshot.FormulaQ, snapshot.CachedQ, "T" + rowNumber, snapshot.Desc1T,
                sheetDescription, snapshot.BlockLabel);
            bool desc2Shown = trustTemplates && SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                snapshot.FormulaR, snapshot.CachedR, "U" + rowNumber, snapshot.Desc2U,
                sheetDescription, snapshot.BlockLabel);
            row.KeptFormulaShowsDescription = desc1Shown || desc2Shown;
            bool writesDesc1 = !trustTemplates || desc1Shown;
            bool writesDesc2 = !trustTemplates || desc2Shown;
            string desc1Value = writesDesc1 ? descriptions.Desc1 + additionToDescription : additionToDescription;
            string desc2Value = writesDesc2 ? descriptions.Desc2 + additionToDescription : additionToDescription;

            // Description 2 is all a single-language SCADA shows, so its falling back to language-1
            // text is never silent once a second language is configured — one other than English,
            // language 1 itself, where repeating Description 1 is the design.
            string secondLanguage = (options.SecondLanguage ?? "").Trim();
            if (secondLanguage.Length > 0 && writesDesc2 &&
                !string.Equals(secondLanguage, "EN", StringComparison.OrdinalIgnoreCase))
            {
                if (descriptions.Desc2FromLanguage1)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "no " + secondLanguage + " text for Description 2 — neither the IO list's language 2 " +
                        "nor the glossary provides one, so it repeats language 1: \"" + descriptions.Desc2.Trim() + "\"",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
                else if (descriptions.Desc2VariablePartFromLanguage1)
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the glossary translates the permanent part, but the IO list has no " + secondLanguage +
                        " variable part — \"" + descriptions.Desc2RidingText + "\" stays in language 1 in Description 2",
                        entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            }

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
            // unless the per-row force flag insists — the legacy rule verbatim. The blank is
            // proposed only in append mode, where the device moves into the label: an IO row
            // WITHOUT a device has nothing to move, so the SWP's device stays (and is reported,
            // because the IO list is supposed to carry one).
            bool deviceWritable = forceElectricalIds || !isBeltSheet;
            if (deviceWritable && ioDevice.Length > 0 &&
                !SwpFieldCompare.Same(SwpField.DeviceJ, snapshot.DeviceJ, deviceValue))
                Edit(SwpField.DeviceJ, snapshot.DeviceJ, deviceValue);
            else if (deviceWritable && ioDevice.Length == 0 && hasIdentity &&
                     entry.GeneratedBy.Length == 0 && snapshot.DeviceJ.Trim().Length > 0)
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the IO list carries no device for this row — the SWP's \"" + snapshot.DeviceJ.Trim() +
                    "\" is kept (an empty value never erases one)",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));

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

            // Descriptions: like the T.S. Ref, an EMPTY composed text never blanks what the SWP
            // holds — a stale description still names the alarm, a blank one names nothing. The
            // label addition alone is a text (the FLD lands even when the IO list says nothing);
            // trust rows with no addition compose to "" by design and never touch what the
            // CONFIG templates feed.
            if (desc1Value.Trim().Length == 0 && writesDesc1 &&
                (snapshot.Desc1T.Trim().Length > 0 || snapshot.Desc2U.Trim().Length > 0))
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the composed description is empty — the SWP's own text is kept " +
                    "(an empty value never erases one)",
                    entry.CardSheet, entry.SourceRow, entry.Mnemonic));
            if (desc1Value.Trim().Length > 0 &&
                !SwpFieldCompare.Same(SwpField.Desc1T, snapshot.Desc1T, desc1Value))
                Edit(SwpField.Desc1T, snapshot.Desc1T, desc1Value);
            if (desc2Value.Trim().Length > 0 &&
                !SwpFieldCompare.Same(SwpField.Desc2U, snapshot.Desc2U, desc2Value))
                Edit(SwpField.Desc2U, snapshot.Desc2U, desc2Value);

            // Does the suffix cell CARRY A DESCRIPTION once this batch lands? The truncating form
            // drops the template's trailing text ("… - SPARE ALARM.63") and shows T/U in its
            // place, which is right only where T/U names the bit. Three ways it does not: a
            // trusted row (T/U takes the FLD label alone), a mismatch row whose IO list had no
            // description (the label alone again), and a cell that stays empty. Those keep the
            // plain append — dropping the template's tail there would erase the alarm's own name.
            bool desc1Describes = !trustTemplates &&
                (descriptions.Desc1.Trim().Length > 0 ||
                 (desc1Value.Trim().Length == 0 && snapshot.Desc1T.Trim().Length > 0));
            bool desc2Describes = !trustTemplates &&
                (descriptions.Desc2.Trim().Length > 0 ||
                 (desc2Value.Trim().Length == 0 && snapshot.Desc2U.Trim().Length > 0));

            // A row the workspace default would have rebuilt but a rule or a tick kept: that is the
            // exception worth counting, and it is what the log reports.
            if (trustTemplates && options.ConfigMode == SwpConfigMode.ReplaceTemplateTail)
                row.ConfigFormulaKept = true;

            bool appendPending = appendIdentity;
            ProposeFormulaEdit(row, coordinate, snapshot, rebuildFormulas, appendPending,
                desc1Describes, SwpField.FormulaQ, snapshot.FormulaQ, snapshot.FormulaQShared,
                snapshot.CachedQ, "T", log);
            ProposeFormulaEdit(row, coordinate, snapshot, rebuildFormulas, appendPending,
                desc2Describes, SwpField.FormulaR, snapshot.FormulaR, snapshot.FormulaRShared,
                snapshot.CachedR, "U", log);

            // The FLD group: in append mode the device blank (J) is safe only once the label it
            // moved into is in the file AND displayed — so the blank, the T/U writes and the Q/R
            // extension of one row toggle together on the queue and apply together (one decision,
            // never half of it: a blank J next to an un-appended label is the device existing nowhere).
            if (appendIdentity && row.Edits.Exists(e => e.Field == SwpField.DeviceJ && e.Proposed.Length == 0))
                foreach (SwpProposedEdit edit in row.Edits)
                    if (edit.Field == SwpField.DeviceJ || edit.Field == SwpField.Desc1T ||
                        edit.Field == SwpField.Desc2U || edit.Field == SwpField.FormulaQ ||
                        edit.Field == SwpField.FormulaR)
                        edit.FldGroup = coordinate.FullKey;
        }

        private static void ProposeFormulaEdit(SwpSpineRow row, SwpCoordinate coordinate,
            SwpRowSnapshot snapshot, bool rebuildFormulas, bool appendPending,
            bool suffixDescribes, SwpField field, string formula, bool isShared, string cachedText,
            string descriptionColumn, List<SwpLogEntry> log)
        {
            string proposed = null;
            string descriptionCell = descriptionColumn + snapshot.RowNumber.ToString(CultureInfo.InvariantCulture);
            string formulaColumn = descriptionColumn == "T" ? "Q" : "R";   // the final-text cell the log names

            if (rebuildFormulas && suffixDescribes)
            {
                // The rebuild works from the row's OWN template reference: truncated at its last
                // placeholder, so the description replaces the template's generic tail instead of
                // stacking onto it. A row whose reference is gone (an older wholesale rewrite
                // destroyed it) cannot be rebuilt from itself — the reference-SWP formula audit
                // is what restores a lost template, and the row is named rather than guessed at.
                if (SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed template))
                    proposed = SwpDescriptionFormula.Build(template, descriptionCell);
                else if (!isShared && formula.Length > 0 && !SwpDescriptionFormula.IsTruncatingForm(formula))
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the " + formulaColumn + " final-text formula at " + coordinate.FullKey +
                        " carries no CONFIG_PLANT template to rebuild from — left untouched; " +
                        "run the audit against a reference SWP to restore it",
                        row.Entry.CardSheet, row.Entry.SourceRow, row.Entry.Mnemonic));
                else if (!isShared && formula.Length == 0 && !appendPending)
                    // A cell with NO formula cannot show the description at all, and nothing can
                    // rebuild it from itself — the same dead end, and it must not be silent either:
                    // it is exactly what a formula Excel removed while repairing a file leaves
                    // behind (the correct cached text, frozen). A mismatch row is not named here,
                    // because the append path below rewrites its value as a formula anyway.
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the " + formulaColumn + " final-text cell at " + coordinate.FullKey +
                        ((cachedText ?? "").Trim().Length > 0 ? " holds a plain value" : " is empty") +
                        ", not a formula — there is no CONFIG_PLANT template to rebuild it from, so " +
                        descriptionCell + " never reaches the displayed text (a paste-over, or a formula " +
                        "Excel removed while repairing the file); restore it from the reference SWP or the " +
                        "file's .bak",
                        row.Entry.CardSheet, row.Entry.SourceRow, row.Entry.Mnemonic));
            }

            if (proposed == null && appendPending)
            {
                if (isShared)
                {
                    // A shared-formula member carries no text of its own (a child) or anchors
                    // its siblings (the master): neither can be extended headlessly. Said now,
                    // not silently skipped — the appended identity displays only if the shared
                    // formula already references the description cell.
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the " + formulaColumn + " final-text formula at " + coordinate.FullKey +
                        " is a shared formula — left untouched; the appended identity is displayed only " +
                        "if that formula already references " + descriptionCell,
                        row.Entry.CardSheet, row.Entry.SourceRow, row.Entry.Mnemonic));
                    return;
                }
                if (formula.Length == 0)
                {
                    // A value-only Q/R cell (a pasted-over literal) displays nothing of T/U: the
                    // appended identity could never reach SCADA, and J is blanked on the strength
                    // of that label. The literal becomes the head of a formula that appends the
                    // description cell — the same text, plus the label.
                    string literal = (cachedText ?? "").Trim();
                    proposed = literal.Length == 0
                        ? descriptionCell
                        : "\"" + literal.Replace("\"", "\"\"") + "\"&\" \"&" + descriptionCell;
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "the " + formulaColumn + " final-text cell at " + coordinate.FullKey +
                        " is a plain value, not a formula — rewritten as a formula over the same text " +
                        "with " + descriptionCell + " appended, so the appended identity is displayed",
                        row.Entry.CardSheet, row.Entry.SourceRow, row.Entry.Mnemonic));
                }
                else if (!SwpDescriptionFormula.ReferencesCell(formula, descriptionCell))
                {
                    // In APPEND mode the extension is automatic for TEMPLATE-ONLY formulas — the
                    // only way the appended identity reaches the displayed text. A formula already
                    // referencing T{row}/U{row} (the standard SUBSTITUTE(template&T,…) shape) flows
                    // the addition by itself — extending it would print the description twice.
                    //
                    // Which extension depends on what T/U will hold. A description belongs where
                    // the template's generic tail is, so it TRUNCATES; a bare FLD label belongs
                    // after that tail, so it is appended and the alarm's own name survives.
                    if (suffixDescribes &&
                        SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed template))
                        proposed = SwpDescriptionFormula.Build(template, descriptionCell);
                    else
                        proposed = formula + "&\" \"&" + descriptionCell;
                }
            }

            if (proposed == null || proposed == formula) return;

            if (isShared)
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the " + formulaColumn + " final-text formula at " + coordinate.FullKey +
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
        /// What the Q/R cell will DISPLAY after Excel recalculates, simulated from the cached
        /// text. A formula REFERENCING the description cell displays template &amp; desc — the
        /// cached final ends with the CURRENT desc and the simulation swaps that suffix for the
        /// proposed one (never concatenating old and new). A TEMPLATE-ONLY formula shows the
        /// description only through a pending append ( &amp;" "&amp;T ) — exactly once; without
        /// one, a description edit does not change the final text at all. A pending REBUILD cuts
        /// the template's trailing text off first, so the description lands where that tail was.
        /// Both field bugs of 2026-08-21 (the old-plus-new mix, then the doubled append) are
        /// pinned by tests.
        /// </summary>
        public static string SimulateRecalculatedFinal(string cachedFinal, string currentDesc,
            string proposedDesc, bool formulaReferencesDesc, bool formulaAppendPending,
            bool rebuildPending, string sheetDescription, string blockLabel)
        {
            string desc = proposedDesc ?? "";

            // The rebuild reads the same cached text: its head ends at the template's last
            // placeholder — the block label, or the sheet description on a one-placeholder
            // template — and everything after it is the tail the truncation drops.
            if (rebuildPending)
                return SwpDescriptionFormula.RenderedFinal(cachedFinal, sheetDescription, blockLabel, desc);

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
