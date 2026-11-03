using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using ExcelPowerTools.Core.Csv;
using ExcelPowerTools.Core.Files;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// Build/refresh orchestration. The long phases (SWP streaming read, filter, diff) are
    /// headless and run curtain-free with the dashboard as the progress surface; only the
    /// COM-visible rendering happens behind the state guard the commands hold. The dashboard's
    /// Cancel flag is polled between phases — a phase in flight always completes.
    /// </summary>
    internal static class SwpBridgeBuilder
    {
        /// <summary>Everything from documents to model. NO COM except the IO-list read.</summary>
        public static void BuildModel(dynamic app, SwpBridgeSession session)
        {
            SwpBridgeSettings settings = session.Settings;
            session.BuildLog = new List<SwpLogEntry>();
            List<SwpLogEntry> log = session.BuildLog;
            SwpBridgeDashboard dashboard = session.Dashboard;

            ResolveProductionFiles(session);

            dashboard?.PostPhase("Reading the IO list…");
            session.IoRows = SwpBridgeCardReader.Read(app, settings.IoListPath, settings.CardSheets, log);
            if (session.IoRows.Count == 0)
                throw new InvalidOperationException("None of the selected card sheets yielded any rows.");
            ThrowIfCancelled(session);

            dashboard?.PostPhase("Reading the SWP (" + settings.SwpFileName + ")…");
            session.SwpBytes = ReadFileBytes(session.SwpPath, "SWP workbook");
            ThrowIfCancelled(session);

            session.BaseIndex = null;
            if (File.Exists(settings.ReferenceSwpPath))
            {
                dashboard?.PostPhase("Reading the reference SWP…");
                session.BaseIndex = SwpSheetReader.Read(ReadFileBytes(settings.ReferenceSwpPath, "reference SWP"),
                    new string[0], session.BuildLog, includeAllDiagnosticSheets: true);
                ThrowIfCancelled(session);
            }

            session.AlarmStringsFiles.Clear();
            foreach (string path in settings.AlarmStringsPaths)
            {
                if (!File.Exists(path))
                {
                    session.BuildLog.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "alarm-strings file not found: " + path));
                    continue;
                }
                try { session.AlarmStringsFiles.Add((path, File.ReadAllLines(path))); }
                catch (Exception ex)
                {
                    session.BuildLog.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "alarm-strings file could not be read: " + path + " — " + ex.Message));
                }
            }

            LoadGlossaries(session);
            RecomputeFromMemory(session);
        }

        /// <summary>
        /// One PRODUCTION root → every plant input, by the realtime tree layout. The resolved
        /// paths land in the settings (and so in the sidecar), so what was actually used is on
        /// record; a missing input keeps its check off, like an empty per-file path always did.
        /// </summary>
        private static void ResolveProductionFiles(SwpBridgeSession session)
        {
            SwpBridgeSettings settings = session.Settings;
            if (string.IsNullOrWhiteSpace(settings.ProductionRoot)) return;

            SwpProductionFiles files = SwpProductionTree.Resolve(
                settings.ProductionRoot, settings.PlantType, settings.GroupIndex);
            settings.BusfieldCsvPath = files.Busfield.FullPath;
            settings.PlantInputsCsvPath = files.PlantInputs.FullPath;
            settings.DestinationParamsCsvPath = files.DestinationParams.FullPath;
            settings.SensorCsvPath = files.Sensor.FullPath;
            settings.StrobeDeviceCsvPath = files.StrobeDevice.FullPath;
            settings.BusExtendedDiagnosticCsvPath = files.BusExtendedDiagnostic.FullPath;
            settings.AlarmStringsPaths = new List<string>(files.AlarmStrings);

            if (files.Root.Length == 0)
            {
                session.BuildLog.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "production root not found: " + settings.ProductionRoot + " — every plant check stays off"));
                return;
            }
            session.BuildLog.Add(SwpLogEntry.For(SwpLogSeverity.Info, "production root: " + files.Root));
            foreach (string line in files.ReportLines())
                session.BuildLog.Add(SwpLogEntry.For(SwpLogSeverity.Info, "    " + line));
        }

        /// <summary>Parses the optional glossary workbooks into the session baselines.</summary>
        private static void LoadGlossaries(SwpBridgeSession session)
        {
            SwpBridgeSettings settings = session.Settings;
            session.GlossaryStandard = new List<SwpGlossaryPair>();
            session.GlossaryShadow = new List<SwpGlossaryPair>();
            bool anyPath = File.Exists(settings.GlossaryPath1) || File.Exists(settings.GlossaryPath2);
            if (!anyPath) return;

            string code = ResolveSecondLanguageCode(settings);
            if (code.Length == 0)
            {
                session.BuildLog.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "glossaries are selected but no 2nd-language code is set (setup takes the " +
                    "2-letter glossary code, e.g. PT) — nothing can translate"));
                return;
            }

            session.Dashboard?.PostPhase("Reading the glossaries (" + code + ")…");
            if (File.Exists(settings.GlossaryPath1))
                session.GlossaryStandard = SwpGlossary.Parse(
                    ReadFileBytes(settings.GlossaryPath1, "glossary workbook"), code, "standard", session.BuildLog);
            if (File.Exists(settings.GlossaryPath2))
                session.GlossaryShadow = SwpGlossary.Parse(
                    ReadFileBytes(settings.GlossaryPath2, "glossary ombra workbook"), code, "shadow", session.BuildLog);
        }

        internal static string ResolveSecondLanguageCode(SwpBridgeSettings settings) =>
            NormalizeLanguageCode(settings.SecondLanguage);

        /// <summary>
        /// The setting is a free 2-letter glossary code ("PT"); sidecars from before the free
        /// input carry a language NAME ("Português") and resolve through the _Languages table.
        /// </summary>
        internal static string NormalizeLanguageCode(string value)
        {
            string trimmed = (value ?? "").Trim();
            if (trimmed.Length == 2) return trimmed.ToUpperInvariant();
            foreach ((string name, string code, _) in SwpBridgeRegistry.GetLanguages())
                if (string.Equals(name, trimmed, StringComparison.OrdinalIgnoreCase))
                    return (code ?? "").Trim().ToUpperInvariant();
            return "";
        }

        /// <summary>
        /// Filter → generators → map/diff → coherence → queue reconcile, all from the session's
        /// in-memory data. Override and dictionary edits re-enter here without touching a file.
        /// </summary>
        public static void RecomputeFromMemory(SwpBridgeSession session)
        {
            SwpBridgeSettings settings = session.Settings;
            SwpBridgeDashboard dashboard = session.Dashboard;
            var log = new List<SwpLogEntry>();

            dashboard?.PostPhase("Filtering diagnostic mnemonics…");
            // ONE table for the whole build: the filter, the producers and the mapping all read
            // the same registry copy, so an edit on _Mnemonics reaches every stage at once.
            SwpMnemonicMap mnemonicMap = SwpBridgeRegistry.GetMnemonicMap();
            var filterOptions = new SwpFilterOptions
            {
                PlantType = settings.PlantType,
                IncludeMnemonics = TryRegex(settings.IncludeRegex, log),
                ExcludeMnemonics = TryRegex(settings.ExcludeRegex, log),
                Map = mnemonicMap,
            };
            session.Entries = SwpMnemonicFilter.Run(session.IoRows, filterOptions, log);

            // WHICH producers run is the mnemonic map's business, not this method's: a plant with
            // no generate rows generates nothing, and unticking a row switches one off. The
            // manager reports a row naming a producer that does not exist rather than skipping it.
            dashboard?.PostPhase("Generating mnemonics…");
            session.Entries.AddRange(CustomMnemonicsManager.RunAll(
                mnemonicMap,
                new SwpGeneratorContext
                {
                    PlantType = settings.PlantType,
                    IoRows = session.IoRows,
                    StrobeTemplates = session.Sidecar.StrobeTemplates,
                    Log = log,
                    ReadSource = name =>
                    {
                        string path =
                            name.Equals("Sensor", StringComparison.OrdinalIgnoreCase) ? settings.SensorCsvPath :
                            name.Equals("StrobeDevice", StringComparison.OrdinalIgnoreCase) ? settings.StrobeDeviceCsvPath :
                            "";
                        return path.Length > 0 && File.Exists(path) ? File.ReadAllLines(path) : null;
                    },
                }));
            ThrowIfCancelled(session);

            dashboard?.PostPhase("Indexing the SWP sheets…");
            session.IndexRules = SwpBridgeRegistry.GetIndexRules();
            var targetSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var chuteMap = LoadChuteMap(settings, log);
            var indexSources = new Dictionary<string, IReadOnlyDictionary<string, int>>(
                StringComparer.OrdinalIgnoreCase);
            if (chuteMap != null) indexSources["Destination_Parameters"] = chuteMap;
            var overrides = session.OverridesByMnemonic();
            foreach (DiagnosticEntry entry in session.Entries)
            {
                SwpRowOverride o = overrides.TryGetValue(entry.Mnemonic, out SwpRowOverride v) ? v : null;
                SwpMappedMnemonic mapped = mnemonicMap.Map(entry.Mnemonic, settings.PlantType,
                    new SwpMapContext
                    {
                        FirstLineId = settings.FirstLineId,
                        GroupIndex = settings.GroupIndex,
                        DeviceIndexOverride = o?.SwpDbIndex,
                        Rules = session.IndexRules,
                        Sources = indexSources,
                    });
                if (mapped.Ok) targetSheets.Add(mapped.Coordinate.SheetName);
            }
            // Target sheets for the join, plus EVERY '#'-marked diagnostic sheet for the
            // formula-integrity check — hand-edit damage hides anywhere, STROBE included.
            session.Index = SwpSheetReader.Read(session.SwpBytes, targetSheets, log,
                includeAllDiagnosticSheets: true);
            WarnOnParameterShapedMisses(session.Index, log);
            ThrowIfCancelled(session);

            dashboard?.PostPhase("Computing the diff…");
            ResolveBusFailureConstants(settings, out string busFailure1, out string busFailure2);

            session.GlossaryMerged = SwpGlossary.Merge(
                session.GlossaryStandard, session.GlossaryShadow, session.Sidecar.GlossaryTweaks);
            List<SwpGlossaryPair> glossary = session.GlossaryMerged.Count > 0 ? session.GlossaryMerged : null;

            session.Model = SwpDiff.Build(session.Entries, session.Index, new SwpDiffOptions
            {
                PlantType = settings.PlantType,
                FirstLineId = settings.FirstLineId,
                GroupIndex = settings.GroupIndex,
                ConfigMode = settings.ConfigMode,
                ConfigSheetRule = SwpConfigSheetRule.Create(
                    settings.ConfigSheetRulePattern, settings.ConfigSheetRuleMode),
                MismatchMode = settings.MismatchMode,
                FldSheetRule = SwpFldSheetRule.Create(settings.FldSheetRulePattern, settings.FldSheetRuleMode),
                BusFailureLang1 = busFailure1,
                BusFailureLang2 = busFailure2,
                GlossaryPairs = glossary,
                GlossaryTranslateAlways = settings.GlossaryTranslateAlways,
                SecondLanguage = ResolveSecondLanguageCode(settings),
                Dictionary1 = session.Sidecar.Dictionary1,
                Dictionary2 = session.Sidecar.Dictionary2,
                ChuteMap = chuteMap,
                Overrides = overrides,
                IndexRules = session.IndexRules,
            }, log);

            dashboard?.PostPhase("Checking the formulas…");
            // Brutal-overwrite mode writes the IO identity into H/I as plain values: the cleaner
            // reads those as deliberate, and never queues a repair over a cell this build is
            // about to overwrite.
            var deliberate = new Dictionary<(string Sheet, int Row), (string FUnit, string Location)>();
            var overwriteTargets = new HashSet<string>(StringComparer.Ordinal);
            // Per ROW, not per workspace: [Overwrite FLD] / [Append FLD] flips individual rows
            // away from the default, so only the rows actually in overwrite mode are exempted.
            foreach (SwpSpineRow row in session.Model.Rows)
            {
                if (!row.FldOverwriteMode) continue;
                if (row.Snapshot == null || row.Mapped == null || !row.Mapped.Ok) continue;
                deliberate[(row.Mapped.Coordinate.SheetName, row.Snapshot.RowNumber)] = (row.Entry.FUnit, row.Entry.Location);
                foreach (SwpProposedEdit edit in row.Edits)
                {
                    if (edit.Field == SwpField.FUnitH) overwriteTargets.Add(edit.Sheet + "|" + edit.RowNumber + "|H");
                    else if (edit.Field == SwpField.LocationI) overwriteTargets.Add(edit.Sheet + "|" + edit.RowNumber + "|I");
                }
            }
            List<SwpProposedEdit> repairs = SwpFormulaCheck.ProposeRepairs(session.Index, log,
                deliberate.Count > 0 ? deliberate : null);
            repairs.RemoveAll(r =>
                (r.Field == SwpField.FormulaH && overwriteTargets.Contains(r.Sheet + "|" + r.RowNumber + "|H")) ||
                (r.Field == SwpField.FormulaI && overwriteTargets.Contains(r.Sheet + "|" + r.RowNumber + "|I")));
            if (repairs.Count > 0)
            {
                session.Model.Edits.AddRange(repairs);
                log.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    repairs.Count + " formula repair(s) queued — broken H/I references, header " +
                    "locations or strobe formulas (see the lines above)"));
            }
            if (session.BaseIndex != null)
            {
                dashboard?.PostPhase("Auditing against the reference SWP…");
                SwpFormulaCheck.AuditAgainstBase(session.Index, session.BaseIndex, log);
            }

            // The official-rules content audits (duplicated strings, forbidden bits, priority
            // column, dashed locations) — log findings over the whole indexed SWP.
            //
            // The forbidden-bit check needs to know where the mnemonics actually LAND: an export
            // tick on bit 31/63 is ignored by the plant and worth nothing to report, while a
            // mnemonic landing there is a signal that will never be seen.
            var targetedRows = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (session.Model != null)
                foreach (SwpSpineRow row in session.Model.Rows)
                    if (row.Mapped != null && row.Mapped.Ok && row.Snapshot != null)
                        targetedRows.Add(SwpContentChecks.RowKeyOf(
                            row.Mapped.Coordinate.SheetName, row.Snapshot.RowNumber));
            SwpContentChecks.Run(session.Index, log, targetedRows);

            AnnotateCoherence(session, log);

            // Session relevance: pairs whose EN phrase IS some row's full permanent part.
            var glossaryKeys = new List<string>();
            foreach (SwpSpineRow row in session.Model.Rows) glossaryKeys.Add(row.GlossaryKey);
            session.GlossaryRelevant = SwpGlossary.RelevantPairs(
                session.GlossaryMerged, glossaryKeys, out session.GlossaryOccurrences);

            dashboard?.PostPhase("Reconciling the queue…");
            session.Queue.ReconcileFromDiff(session.Model.Edits,
                action => ResolveCurrent(session.Index, action),
                DateTime.UtcNow.ToString("O"));

            session.BuildLog.AddRange(log);
            dashboard?.PostModel(session);
        }

        /// <summary>The file's current value for an action's target cell; null when the row is gone.</summary>
        internal static string ResolveCurrent(SwpWorkbookIndex index, SwpAction action)
        {
            if (index == null || !index.Sheets.TryGetValue(action.Sheet, out SwpSheetIndex sheet)) return null;

            if (action.CoordinateKey.StartsWith("header>", StringComparison.Ordinal))
            {
                if (!int.TryParse(action.CoordinateKey.Substring("header>".Length), out int headerRow) ||
                    !sheet.BlockHeaderByRow.TryGetValue(headerRow, out SwpBlockHeader header))
                    return null;
                return action.Field == SwpField.HeaderFormulaI ? header.LocationIFormula : header.LocationI;
            }

            if (!sheet.RowsByKey.TryGetValue(action.CoordinateKey, out SwpRowSnapshot row)) return null;
            switch (action.Field)
            {
                case SwpField.ExportB: return row.ExportB;
                case SwpField.AddressG: return row.AddressG;
                case SwpField.DeviceJ: return row.DeviceJ;
                case SwpField.TsRefL: return row.TsRefL;
                case SwpField.Desc1T: return row.Desc1T;
                case SwpField.Desc2U: return row.Desc2U;
                case SwpField.FormulaQ: return row.FormulaQ;
                case SwpField.FormulaR: return row.FormulaR;
                case SwpField.FormulaH: return row.FormulaH;
                case SwpField.FormulaI: return row.FormulaI;
                case SwpField.FUnitH: return row.CachedH;
                case SwpField.LocationI: return row.CachedI;
                default: return "";
            }
        }

        /// <summary>
        /// The first language (row 1 of the global table — the SWP's Description 1 language)
        /// and the workspace's chosen second language, resolved to their bus-failure texts.
        /// The second language is a code ("PT"); rows match by Code, with the language NAME
        /// still accepted for pre-free-input sidecars. No matching row → empty text (a bus
        /// module's description 2 then falls back to description 1).
        /// </summary>
        internal static void ResolveBusFailureConstants(SwpBridgeSettings settings,
            out string lang1, out string lang2)
        {
            List<(string Name, string Code, string Text)> languages = SwpBridgeRegistry.GetLanguages();
            lang1 = languages.Count > 0 ? languages[0].Text : SwpDescriptions.DefaultBusFailureText;
            lang2 = "";
            string code = NormalizeLanguageCode(settings.SecondLanguage);
            foreach ((string name, string rowCode, string text) in languages)
                if ((code.Length > 0 && string.Equals((rowCode ?? "").Trim(), code, StringComparison.OrdinalIgnoreCase)) ||
                    string.Equals(name, settings.SecondLanguage, StringComparison.OrdinalIgnoreCase))
                {
                    lang2 = text;
                    break;
                }
        }

        /// <summary>Renders every workspace sheet. Call inside an ExcelStateGuard.</summary>
        public static void RenderWorkspace(dynamic workbook, SwpBridgeSession session)
        {
            session.Dashboard?.PostPhase("Rendering the workspace…");

            RemovePreviewSheet(workbook);
            dynamic spine = SwpBridgeQueueSheet.FindSheet(workbook, SwpBridgeWorkspace.SpineSheetName)
                ?? workbook.Worksheets(1);
            SwpBridgeSpineSheet.Render(spine, session);
            SwpBridgeQueueSheet.Render(workbook, session);
            SwpBridgeAuxSheets.RenderLog(workbook, session);
            SwpBridgeAuxSheets.RenderDictionary(workbook, session);
            SwpBridgeAuxSheets.RenderGlossary(workbook, session);
            SwpBridgeAuxSheets.RenderOffsets(workbook);
            SwpBridgeAuxSheets.RenderMnemonics(workbook);
            SwpBridgeAuxSheets.RenderStrobeTemplates(workbook, session);
            SwpBridgeAuxSheets.RenderHome(workbook, session);

            // The _Preview sheet opens with the workspace (the first previewable row, or the row
            // the user was previewing across a refresh) — no double-click needed to see it exists.
            int previewRow = session.PreviewContext?.ModelRow ?? FirstPreviewableRow(session);
            SwpBridgePreview.Prepare(session, previewRow);

            ColorWorkspaceTabs(workbook);
            try { workbook.Worksheets(SwpBridgeWorkspace.SpineSheetName).Activate(); } catch { }

            session.Dashboard?.PostPhase("Ready");
            session.Dashboard?.PostModel(session);
        }

        /// <summary>
        /// The override/dictionary/language fast path: recompute in memory + re-render the
        /// generated sheets. Runs INSIDE a SheetChange handler, so deliberately NO
        /// ExcelStateGuard — toggling ScreenUpdating there defers the repaint until the next
        /// full refresh (the io-bridge lesson); the caller's reentry gate already stops the
        /// event cascade.
        /// </summary>
        public static void RecomputeAndRerender(SwpBridgeSession session)
        {
            dynamic app = ExcelApp.App;
            dynamic workbook = SwpBridgeWorkspace.FindLiveWorkbook(app);
            if (workbook == null) return;

            RecomputeFromMemory(session);
            dynamic spine = SwpBridgeQueueSheet.FindSheet(workbook, SwpBridgeWorkspace.SpineSheetName);
            if (spine != null) SwpBridgeSpineSheet.Render(spine, session);
            SwpBridgeQueueSheet.Render(workbook, session);
            SwpBridgeAuxSheets.RenderLog(workbook, session);
            SwpBridgeAuxSheets.RenderGlossary(workbook, session);
            session.SaveSidecarNow();
        }

        internal static void RemovePreviewSheet(dynamic workbook)
        {
            dynamic preview = SwpBridgeQueueSheet.FindSheet(workbook, SwpBridgeWorkspace.PreviewSheetName);
            if (preview == null) return;
            dynamic app = ExcelApp.App;
            bool prevAlerts = true;
            try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
            try { preview.Delete(); } catch { }
            finally { try { app.DisplayAlerts = prevAlerts; } catch { } }
        }

        /// <summary>The first row that actually has an SWP block to preview (matched + mapped).</summary>
        private static int FirstPreviewableRow(SwpBridgeSession session)
        {
            if (session?.Model == null) return -1;
            for (int i = 0; i < session.Model.Rows.Count; i++)
            {
                SwpSpineRow row = session.Model.Rows[i];
                if (row.Snapshot != null && row.Mapped != null && row.Mapped.Ok) return i;
            }
            return -1;
        }

        /// <summary>
        /// Sheet-tab colours so the three sheets a user moves between are found by colour, not by
        /// reading the strip: Home light blue, Workspace orange, _Preview yellow (the same three
        /// hues the sheets use for their own headers / SWP side / pending cells).
        /// </summary>
        private static void ColorWorkspaceTabs(dynamic workbook)
        {
            SetTabColor(workbook, SwpBridgeWorkspace.HomeSheetName, 180, 210, 240);   // light blue
            SetTabColor(workbook, SwpBridgeWorkspace.SpineSheetName, 250, 200, 140);  // orange
            SetTabColor(workbook, SwpBridgeWorkspace.PreviewSheetName, 255, 235, 156);// yellow
        }

        private static void SetTabColor(dynamic workbook, string sheetName, int r, int g, int b)
        {
            dynamic sheet = SwpBridgeQueueSheet.FindSheet(workbook, sheetName);
            if (sheet == null) return;   // _Preview may be absent when nothing is previewable
            try { sheet.Tab.Color = r | (g << 8) | (b << 16); } catch { }
        }

        // ---- pieces -------------------------------------------------------------------------

        private static void AnnotateCoherence(SwpBridgeSession session, List<SwpLogEntry> log)
        {
            SwpBridgeSettings settings = session.Settings;
            CsvDocument plantInputs = null, busfield = null, busExtended = null;
            if (File.Exists(settings.PlantInputsCsvPath))
            {
                session.Dashboard?.PostPhase("Checking Plant_Inputs…");
                plantInputs = LoadCsv(settings.PlantInputsCsvPath, log);
            }
            if (File.Exists(settings.BusfieldCsvPath))
            {
                session.Dashboard?.PostPhase("Checking the busfield…");
                busfield = LoadCsv(settings.BusfieldCsvPath, log);
            }
            if (File.Exists(settings.BusExtendedDiagnosticCsvPath))
            {
                session.Dashboard?.PostPhase("Checking BUS_ExtendedDiagnostic…");
                busExtended = LoadCsv(settings.BusExtendedDiagnosticCsvPath, log);
            }

            List<SwpAlarmString> alarms = null;
            if (session.AlarmStringsFiles.Count > 0)
            {
                session.Dashboard?.PostPhase("Checking the alarm strings…");
                alarms = new List<SwpAlarmString>();

                // A production folder often holds the lists for the whole plant; this SWP reads
                // only the ones its own sheets export (row 5), so a sibling controller's list is
                // skipped rather than parsed into a page of malformed-mnemonic warnings.
                var offered = new List<string>();
                foreach ((string path, string[] _) in session.AlarmStringsFiles)
                    offered.Add(Path.GetFileName(path));
                var owned = new HashSet<string>(
                    SwpAlarmStrings.OwnedFiles(session.Index, offered, log),
                    StringComparer.OrdinalIgnoreCase);

                int read = 0;
                foreach ((string path, string[] lines) in session.AlarmStringsFiles)
                {
                    string name = Path.GetFileName(path);
                    if (!owned.Contains(name)) continue;
                    read++;
                    alarms.AddRange(SwpAlarmStrings.Parse(lines, settings.PlantType, name, log));
                }
                log.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                    alarms.Count + " exported alarm string(s) loaded from " + read + " file(s)" +
                    (read == session.AlarmStringsFiles.Count
                        ? ""
                        : " (" + (session.AlarmStringsFiles.Count - read) +
                          " not part of this SWP)")));
            }

            if (plantInputs != null || busfield != null || alarms != null || busExtended != null)
                SwpCoherence.Annotate(session.Model, plantInputs, busfield, settings.PlantType, log,
                    alarms, settings.FirstLineId, settings.GroupIndex, session.IndexRules, busExtended);
        }

        private static CsvDocument LoadCsv(string path, List<SwpLogEntry> log)
        {
            SourceReadState state = AtomicFileStore.ReadSource(path, out byte[] bytes, out string error);
            if (state != SourceReadState.Present)
            {
                log.Add(SwpLogEntry.For(SwpLogSeverity.Warning, "could not read " + path + ": " + error));
                return null;
            }
            return CsvFile.Load(bytes);
        }

        private static IReadOnlyDictionary<string, int> LoadChuteMap(SwpBridgeSettings settings, List<SwpLogEntry> log)
        {
            if (settings.PlantType != SwpPlantType.Msc || !File.Exists(settings.DestinationParamsCsvPath))
                return null;
            try
            {
                return SwpDestinationParams.ParseChuteMap(File.ReadAllLines(settings.DestinationParamsCsvPath), log);
            }
            catch (Exception ex)
            {
                log.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "Destination_Parameters could not be read: " + ex.Message));
                return null;
            }
        }

        /// <summary>
        /// A resolved sheet name that does not exist usually means a wrong first-line/group
        /// parameter — one prominent hint beats a hundred per-row verdicts.
        /// </summary>
        private static void WarnOnParameterShapedMisses(SwpWorkbookIndex index, List<SwpLogEntry> log)
        {
            if (index.MissingSheets.Count == 0) return;
            log.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                index.MissingSheets.Count + " resolved sheet name(s) do not exist in the SWP (" +
                string.Join(", ", index.MissingSheets) + ") — the workbook has: " +
                string.Join(", ", index.AllSheetNames) +
                ". A wrong plant type / first line / group parameter produces exactly this."));
        }

        internal static byte[] ReadFileBytes(string path) => ReadFileBytes(path, "file");

        /// <summary>
        /// Reads a source workbook, and when it cannot be read (typically: open in Excel, which
        /// locks it) offers close-it-then-Retry instead of failing the whole build. Cancel
        /// throws OperationCanceledException — the commands' cancelled path, which keeps an
        /// existing workspace alive on refresh.
        /// </summary>
        internal static byte[] ReadFileBytes(string path, string label)
        {
            while (true)
            {
                SourceReadState state = AtomicFileStore.ReadSource(path, out byte[] bytes, out string error);
                if (state == SourceReadState.Present) return bytes;

                // The prompt must surface ABOVE the topmost build curtain — an ownerless
                // MessageBox would hide beneath it and the build would look hung.
                System.Windows.Forms.DialogResult choice;
                using (var owner = new System.Windows.Forms.Form
                {
                    TopMost = true,
                    ShowInTaskbar = false,
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                    StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                    Size = new System.Drawing.Size(1, 1),
                    Opacity = 0,
                })
                {
                    owner.Show();
                    choice = Dialogs.RetryCancel(
                        "The " + label + " could not be read:\n" + path +
                        (string.IsNullOrEmpty(error) ? "" : "\n\n" + error) +
                        "\n\nIt may be open in Excel or locked by another program — close it " +
                        "there, then press Retry.",
                        "SWP Bridge", owner)
                        ? System.Windows.Forms.DialogResult.Retry
                        : System.Windows.Forms.DialogResult.Cancel;
                }
                if (choice != System.Windows.Forms.DialogResult.Retry)
                    throw new OperationCanceledException("the " + label + " stayed unreadable");
            }
        }

        private static Regex TryRegex(string pattern, List<SwpLogEntry> log)
        {
            if (string.IsNullOrWhiteSpace(pattern)) return null;
            try { return new Regex(pattern.Trim(), RegexOptions.CultureInvariant); }
            catch (Exception ex)
            {
                log.Add(SwpLogEntry.For(SwpLogSeverity.Warning, "regex \"" + pattern + "\" does not parse: " + ex.Message));
                return null;
            }
        }

        private static void ThrowIfCancelled(SwpBridgeSession session)
        {
            if (session.CancelRequested)
            {
                session.CancelRequested = false;
                throw new OperationCanceledException("cancelled from the dashboard");
            }
        }
    }
}
