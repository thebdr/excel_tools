using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The three supporting sheets: _Log (findings with jump links, generated), _Dictionary and
    /// _StrobeTemplates (both freely editable — their content parses back into the sidecar on
    /// every change and takes effect on the next refresh).
    /// </summary>
    internal static class SwpBridgeAuxSheets
    {
        private const int MaxLogRows = 2000;
        internal const int LogHeaderRow = 2;
        internal const int DictionaryHeaderRow = 3;
        internal const int StrobeFirstRow = 3;

        private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);
        private static readonly int HeaderBlue = Rgb(180, 210, 240);
        private static readonly int ErrorRed = Rgb(255, 199, 206);
        private static readonly int WarningYellow = Rgb(255, 235, 156);

        // ---- _Log ---------------------------------------------------------------------------

        public static void RenderLog(dynamic workbook, SwpBridgeSession session)
        {
            dynamic sheet = EnsureSheet(workbook, SwpBridgeWorkspace.LogSheetName);
            sheet.Cells.Clear();
            try { sheet.Cells.FormatConditions.Delete(); } catch { }

            List<SwpLogEntry> log = session.BuildLog;
            sheet.Cells(1, 1).Value = "SWP Bridge findings · " + log.Count +
                (log.Count > MaxLogRows ? " (first " + MaxLogRows + " shown)" : "") +
                " · click a Go to link to jump to the spine row";
            sheet.Cells(1, 1).Font.Bold = true;

            string[] headers = { "Severity", "Card", "Row", "Mnemonic", "Message", "Go to" };
            var headerData = new object[1, headers.Length];
            for (int c = 0; c < headers.Length; c++) headerData[0, c] = headers[c];
            dynamic headerRange = sheet.Range(sheet.Cells(LogHeaderRow, 1), sheet.Cells(LogHeaderRow, headers.Length));
            headerRange.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, LogHeaderRow, 1, headerData);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;

            int rows = Math.Min(log.Count, MaxLogRows);
            if (rows > 0)
            {
                // First spine worksheet row per mnemonic, for the jump links.
                var spineRowOf = new Dictionary<string, int>(StringComparer.Ordinal);
                if (session.Model != null)
                    for (int i = 0; i < session.Model.Rows.Count; i++)
                    {
                        string m = session.Model.Rows[i].Entry.Mnemonic;
                        if (!spineRowOf.ContainsKey(m) && session.WorksheetRowByModelRow.TryGetValue(i, out int wsRow))
                            spineRowOf[m] = wsRow;
                    }

                var data = new object[rows, 5];
                var links = new string[rows];
                for (int r = 0; r < rows; r++)
                {
                    SwpLogEntry entry = log[r];
                    data[r, 0] = entry.Severity.ToString();
                    data[r, 1] = entry.CardSheet;
                    data[r, 2] = entry.SourceRow > 0 ? entry.SourceRow.ToString() : "";
                    data[r, 3] = entry.Mnemonic;
                    data[r, 4] = entry.Message;
                    links[r] = entry.Mnemonic.Length > 0 && spineRowOf.TryGetValue(entry.Mnemonic, out int row)
                        ? "=HYPERLINK(\"#" + SwpBridgeWorkspace.SpineSheetName + "!C" + row + "\",\"go\")"
                        : "";
                }

                dynamic block = sheet.Range(sheet.Cells(LogHeaderRow + 1, 1), sheet.Cells(LogHeaderRow + rows, 5));
                block.NumberFormat = "@";
                SwpBridgeSpineSheet.WriteBlock(sheet, LogHeaderRow + 1, 1, data);

                var linkData = new object[rows, 1];
                for (int r = 0; r < rows; r++) linkData[r, 0] = links[r];
                sheet.Range(sheet.Cells(LogHeaderRow + 1, 6), sheet.Cells(LogHeaderRow + rows, 6)).Formula = linkData;

                ApplySeverityFormats(sheet, rows, headers.Length);
            }

            try
            {
                int lastRow = LogHeaderRow + Math.Max(1, rows);
                sheet.Range(sheet.Cells(LogHeaderRow, 1), sheet.Cells(lastRow, headers.Length)).Columns.AutoFit();
                dynamic message = sheet.Columns(5);
                if ((double)message.ColumnWidth > 90) message.ColumnWidth = 90;
            }
            catch { }
        }

        private static void ApplySeverityFormats(dynamic sheet, int rows, int columns)
        {
            const int XlExpression = 2;
            dynamic block = sheet.Range(
                sheet.Cells(LogHeaderRow + 1, 1), sheet.Cells(LogHeaderRow + rows, columns));
            try { block.FormatConditions.Delete(); } catch { }
            try { sheet.Cells(LogHeaderRow + 1, 1).Select(); } catch { }

            string severityRef = "$A" + (LogHeaderRow + 1);
            dynamic error = block.FormatConditions.Add(XlExpression, Type.Missing, "=" + severityRef + "=\"Error\"");
            error.Interior.Color = ErrorRed;
            error.StopIfTrue = false;
            dynamic warning = block.FormatConditions.Add(XlExpression, Type.Missing, "=" + severityRef + "=\"Warning\"");
            warning.Interior.Color = WarningYellow;
            warning.StopIfTrue = false;
        }

        /// <summary>The findings as the exported .txt lines (written next to the SWP on apply).</summary>
        public static List<string> LogLines(SwpBridgeSession session)
        {
            var lines = new List<string>();
            foreach (SwpLogEntry entry in session.BuildLog)
            {
                string severity = entry.Severity == SwpLogSeverity.Error ? "[ER]"
                    : entry.Severity == SwpLogSeverity.Warning ? "[WR]" : "[I]";
                string where = entry.CardSheet.Length > 0
                    ? " (" + entry.CardSheet + (entry.SourceRow > 0 ? " row " + entry.SourceRow : "") + ")"
                    : "";
                string mnemonic = entry.Mnemonic.Length > 0 ? " " + entry.Mnemonic : "";
                lines.Add(severity + mnemonic + where + " " + entry.Message);
            }
            return lines;
        }

        // ---- _Dictionary --------------------------------------------------------------------

        public static void RenderDictionary(dynamic workbook, SwpBridgeSession session)
        {
            dynamic sheet = EnsureSheet(workbook, SwpBridgeWorkspace.DictionarySheetName);
            sheet.Cells.Clear();

            sheet.Cells(1, 1).Value =
                "Description dictionary · sequential find→replace over the IO-list texts, in row " +
                "order · edit freely — saved immediately, applied on the next refresh";
            sheet.Cells(1, 1).Font.Bold = true;

            var headers = new object[1, 5];
            headers[0, 0] = "Language 1 — find";
            headers[0, 1] = "Language 1 — replace";
            headers[0, 2] = "";
            headers[0, 3] = "Language 2 — find";
            headers[0, 4] = "Language 2 — replace";
            dynamic headerRange = sheet.Range(sheet.Cells(DictionaryHeaderRow, 1), sheet.Cells(DictionaryHeaderRow, 5));
            headerRange.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, DictionaryHeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;

            int rows = Math.Max(Math.Max(session.Sidecar.Dictionary1.Count, session.Sidecar.Dictionary2.Count), 1);
            var data = new object[rows, 5];
            for (int r = 0; r < session.Sidecar.Dictionary1.Count; r++)
            {
                data[r, 0] = session.Sidecar.Dictionary1[r].Find;
                data[r, 1] = session.Sidecar.Dictionary1[r].Replace;
            }
            for (int r = 0; r < session.Sidecar.Dictionary2.Count; r++)
            {
                data[r, 3] = session.Sidecar.Dictionary2[r].Find;
                data[r, 4] = session.Sidecar.Dictionary2[r].Replace;
            }
            dynamic block = sheet.Range(sheet.Cells(DictionaryHeaderRow + 1, 1), sheet.Cells(DictionaryHeaderRow + rows, 5));
            block.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, DictionaryHeaderRow + 1, 1, data);

            try
            {
                sheet.Range(sheet.Cells(DictionaryHeaderRow, 1), sheet.Cells(DictionaryHeaderRow + rows, 5))
                    .Columns.AutoFit();
            }
            catch { }
        }

        /// <summary>Reads the dictionary sheet back into the sidecar. True when something parsed.</summary>
        public static bool ParseDictionary(dynamic sheet, SwpBridgeSession session)
        {
            try
            {
                long lastRow = Math.Max(RangeHelpers.LastRow(sheet), DictionaryHeaderRow);
                var one = new List<SwpDictionaryEntry>();
                var two = new List<SwpDictionaryEntry>();
                for (int r = DictionaryHeaderRow + 1; r <= lastRow; r++)
                {
                    string find1 = CellText(sheet, r, 1), replace1 = CellText(sheet, r, 2);
                    string find2 = CellText(sheet, r, 4), replace2 = CellText(sheet, r, 5);
                    if (find1.Length > 0) one.Add(new SwpDictionaryEntry { Find = find1, Replace = replace1 });
                    if (find2.Length > 0) two.Add(new SwpDictionaryEntry { Find = find2, Replace = replace2 });
                }
                session.Sidecar.Dictionary1 = one;
                session.Sidecar.Dictionary2 = two;
                return true;
            }
            catch { return false; }
        }

        // ---- _Glossary ----------------------------------------------------------------------

        internal const int GlossaryHeaderRow = 3;
        internal const int GlossaryUseColumn = 1;

        public static void RenderGlossary(dynamic workbook, SwpBridgeSession session)
        {
            dynamic sheet = SwpBridgeQueueSheet.FindSheet(workbook, SwpBridgeWorkspace.GlossarySheetName);
            bool haveGlossary = session.GlossaryMerged != null && session.GlossaryMerged.Count > 0;
            if (!haveGlossary)
            {
                // No glossary in this workspace — no sheet either.
                if (sheet != null)
                {
                    dynamic app = ExcelApp.App;
                    bool prevAlerts = true;
                    try { prevAlerts = (bool)app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
                    try { sheet.Delete(); } catch { }
                    finally { try { app.DisplayAlerts = prevAlerts; } catch { } }
                }
                session.GlossaryDisplayedKeys.Clear();
                return;
            }

            sheet = EnsureSheet(workbook, SwpBridgeWorkspace.GlossarySheetName);
            sheet.Cells.Clear();
            try { sheet.Cells.FormatConditions.Delete(); } catch { }

            sheet.Cells(1, 1).Value =
                "Glossary · the session-relevant pairs (" + session.GlossaryRelevant.Count + " of " +
                session.GlossaryMerged.Count + " translate something here) · double-click Use to " +
                "disable a pair · translations are editable, new EN+translation rows become custom " +
                "pairs · longest EN phrase wins · decisions persist in the sidecar";
            sheet.Cells(1, 1).Font.Bold = true;

            string[] headers = { "Use", "EN", "Translation", "Source", "Rows hit" };
            var headerData = new object[1, headers.Length];
            for (int c = 0; c < headers.Length; c++) headerData[0, c] = headers[c];
            dynamic headerRange = sheet.Range(sheet.Cells(GlossaryHeaderRow, 1), sheet.Cells(GlossaryHeaderRow, headers.Length));
            headerRange.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, GlossaryHeaderRow, 1, headerData);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;

            session.GlossaryDisplayedKeys.Clear();
            int rows = session.GlossaryRelevant.Count;
            if (rows > 0)
            {
                var data = new object[rows, headers.Length];
                for (int r = 0; r < rows; r++)
                {
                    SwpGlossaryPair pair = session.GlossaryRelevant[r];
                    session.GlossaryDisplayedKeys.Add(pair.En);
                    data[r, 0] = pair.Enabled ? SwpBridgeQueueSheet.IncludedGlyph : SwpBridgeQueueSheet.ExcludedGlyph;
                    data[r, 1] = pair.En;
                    data[r, 2] = pair.Translation;
                    data[r, 3] = pair.Source;
                    data[r, 4] = session.GlossaryOccurrences.TryGetValue(pair.En, out int count)
                        ? count.ToString() : "0";
                }
                dynamic block = sheet.Range(
                    sheet.Cells(GlossaryHeaderRow + 1, 1), sheet.Cells(GlossaryHeaderRow + rows, headers.Length));
                block.NumberFormat = "@";
                SwpBridgeSpineSheet.WriteBlock(sheet, GlossaryHeaderRow + 1, 1, data);
                SwpBridgeSpineSheet.MarkDoubleClickable(sheet.Range(
                    sheet.Cells(GlossaryHeaderRow + 1, GlossaryUseColumn),
                    sheet.Cells(GlossaryHeaderRow + rows, GlossaryUseColumn)));
            }

            try
            {
                int lastRow = GlossaryHeaderRow + Math.Max(1, rows);
                sheet.Range(sheet.Cells(GlossaryHeaderRow, 1), sheet.Cells(lastRow, headers.Length)).Columns.AutoFit();
                sheet.Columns(GlossaryUseColumn).ColumnWidth = 6;
                for (int c = 2; c <= 3; c++)
                {
                    dynamic column = sheet.Columns(c);
                    if ((double)column.ColumnWidth > 70) column.ColumnWidth = 70;
                }
            }
            catch { }
        }

        /// <summary>
        /// Reads the _Glossary sheet back into the sidecar tweaks. The sheet shows only the
        /// session-relevant pairs, so only THEIR tweaks are rewritten — decisions about pairs the
        /// current workspace never displays survive untouched.
        /// </summary>
        public static bool ParseGlossary(dynamic sheet, SwpBridgeSession session)
        {
            try
            {
                // The file baseline WITHOUT tweaks — what a sheet row must differ from to matter.
                var baseline = new Dictionary<string, SwpGlossaryPair>(StringComparer.OrdinalIgnoreCase);
                foreach (SwpGlossaryPair pair in SwpGlossary.Merge(
                    session.GlossaryStandard, session.GlossaryShadow, null))
                    baseline[pair.En] = pair;

                var tweaks = new List<SwpGlossaryTweak>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                long lastRow = Math.Max(RangeHelpers.LastRow(sheet), GlossaryHeaderRow);
                for (int r = GlossaryHeaderRow + 1; r <= lastRow; r++)
                {
                    string en = CellText(sheet, r, 2);
                    if (en.Length == 0 || !seen.Add(en)) continue;
                    string glyph = CellText(sheet, r, 1).ToLowerInvariant();
                    bool enabled = glyph != SwpBridgeQueueSheet.ExcludedGlyph && glyph != "0" &&
                                   glyph != "no" && glyph != "false" && glyph != "-";
                    string translation = CellText(sheet, r, 3);

                    if (baseline.TryGetValue(en, out SwpGlossaryPair basePair))
                    {
                        bool translationChanged = translation.Length > 0 && translation != basePair.Translation;
                        if (!enabled || translationChanged)
                            tweaks.Add(new SwpGlossaryTweak
                            {
                                En = en,
                                Enabled = enabled,
                                Translation = translationChanged ? translation : "",
                            });
                    }
                    else if (translation.Length > 0)
                    {
                        tweaks.Add(new SwpGlossaryTweak { En = en, Enabled = enabled, Translation = translation });
                    }
                }

                // Keep the decisions about pairs this workspace does not display.
                foreach (SwpGlossaryTweak old in session.Sidecar.GlossaryTweaks)
                    if (!session.GlossaryDisplayedKeys.Contains(old.En) && !seen.Contains(old.En))
                        tweaks.Add(old);

                session.Sidecar.GlossaryTweaks = tweaks;
                return true;
            }
            catch { return false; }
        }

        // ---- _Offsets -----------------------------------------------------------------------

        internal const int OffsetsHeaderRow = 4;

        public static void RenderOffsets(dynamic workbook)
        {
            dynamic sheet = EnsureSheet(workbook, SwpBridgeWorkspace.OffsetsSheetName);
            sheet.Cells.Clear();

            sheet.Cells(1, 1).Value =
                "Index rules per plant type and diagnosable object — seeded from \"SWP and " +
                "diagnostic configuration rules\" v93 (§4.9 sorter, §5.2 induct, §7.10 upstream) · " +
                "edits persist for ALL systems and sessions";
            sheet.Cells(1, 1).Font.Bold = true;
            sheet.Cells(2, 1).Value =
                "Mnemonic start = the IO-list numbering base · SWP entry start = the table-entry " +
                "base (the shift between them re-addresses the join) · OPC tag start = the " +
                "alarm-export base (drives the Alarm-strings check) · Alarm check x = the object's " +
                "exported tags are verified (only cabinet/belt/zone tag families can be)";

            var headers = new object[1, 6];
            headers[0, 0] = "Plant";
            headers[0, 1] = "Object";
            headers[0, 2] = "Mnemonic start";
            headers[0, 3] = "SWP entry start";
            headers[0, 4] = "OPC tag start";
            headers[0, 5] = "Alarm check";
            dynamic headerRange = sheet.Range(sheet.Cells(OffsetsHeaderRow, 1), sheet.Cells(OffsetsHeaderRow, 6));
            headerRange.NumberFormat = "@";
            SwpBridgeSpineSheet.WriteBlock(sheet, OffsetsHeaderRow, 1, headers);
            headerRange.Interior.Color = HeaderBlue;
            headerRange.Font.Bold = true;

            List<SwpIndexRule> rules = SwpBridgeRegistry.GetIndexRules().Rules;
            var data = new object[rules.Count, 6];
            for (int r = 0; r < rules.Count; r++)
            {
                data[r, 0] = rules[r].PlantType.ToString();
                data[r, 1] = rules[r].ObjectKey;
                data[r, 2] = rules[r].MnemonicStart.ToString(CultureInfo.InvariantCulture);
                data[r, 3] = rules[r].SwpEntryStart.ToString(CultureInfo.InvariantCulture);
                data[r, 4] = rules[r].OpcTagStart.ToString(CultureInfo.InvariantCulture);
                data[r, 5] = rules[r].AlarmCheckEnabled ? "x" : "-";
            }
            dynamic block = sheet.Range(
                sheet.Cells(OffsetsHeaderRow + 1, 1), sheet.Cells(OffsetsHeaderRow + Math.Max(rules.Count, 1), 6));
            block.NumberFormat = "@";
            if (rules.Count > 0) SwpBridgeSpineSheet.WriteBlock(sheet, OffsetsHeaderRow + 1, 1, data);

            try
            {
                sheet.Range(sheet.Cells(OffsetsHeaderRow, 1),
                    sheet.Cells(OffsetsHeaderRow + rules.Count, 6)).Columns.AutoFit();
            }
            catch { }
        }

        /// <summary>Reads the offsets sheet back into the GLOBAL registry table.</summary>
        public static bool ParseOffsets(dynamic sheet)
        {
            try
            {
                long lastRow = Math.Max(RangeHelpers.LastRow(sheet), OffsetsHeaderRow);
                var lines = new List<string>();
                for (int r = OffsetsHeaderRow + 1; r <= lastRow; r++)
                {
                    string plant = CellText(sheet, r, 1);
                    if (plant.Length == 0) continue;
                    lines.Add(plant + "|" + CellText(sheet, r, 2) + "|" + CellText(sheet, r, 3) +
                        "|" + CellText(sheet, r, 4) + "|" + CellText(sheet, r, 5) +
                        "|" + (CellText(sheet, r, 6).Equals("x", StringComparison.OrdinalIgnoreCase) ? "x" : "-"));
                }
                SwpIndexRules rules = SwpIndexRules.FromLines(lines);
                if (rules == null) return false;
                SwpBridgeRegistry.SetIndexRules(rules);
                return true;
            }
            catch { return false; }
        }

        // ---- _TSRef (the glossario card lookup) ---------------------------------------------

        internal const int TsRefHeaderRow = 2;

        /// <summary>
        /// The quick card lookup: one row per TS_nnn, its EN descriptors across the row —
        /// rendered on demand (double-click a T.S. ref cell on the Workspace) and jumped to
        /// the requested card. A scratch view like _Preview; rebuilt on every open.
        /// </summary>
        public static void OpenTsRefLookup(dynamic workbook, SwpBridgeSession session, string card)
        {
            if (session.GlossaryMerged == null || session.GlossaryMerged.Count == 0)
            {
                try { ExcelApp.App.StatusBar = "SWP Bridge: no glossary in this workspace — nothing to look up."; } catch { }
                return;
            }

            // Group the merged pairs by card, keeping the glossary's order inside each.
            var byCard = new SortedDictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (SwpGlossaryPair pair in session.GlossaryMerged)
            {
                string tsRef = pair.TsRef.Trim();
                if (tsRef.Length == 0) continue;
                if (!byCard.TryGetValue(tsRef, out List<string> descriptors))
                    byCard[tsRef] = descriptors = new List<string>();
                descriptors.Add(pair.En);
            }

            dynamic sheet = EnsureSheet(workbook, SwpBridgeWorkspace.TsRefSheetName);
            sheet.Cells.Clear();
            sheet.Cells(1, 1).Value =
                "T.S. card lookup · one row per card, its English descriptors alongside · from " +
                "the loaded glossario (" + byCard.Count + " cards) · scratch view, rebuilt on every open";
            sheet.Cells(1, 1).Font.Bold = true;

            int maxDescriptors = 1;
            foreach (List<string> descriptors in byCard.Values)
                if (descriptors.Count > maxDescriptors) maxDescriptors = descriptors.Count;

            var data = new object[byCard.Count, maxDescriptors + 1];
            int jumpRow = 0, r = 0;
            foreach (KeyValuePair<string, List<string>> entry in byCard)
            {
                data[r, 0] = entry.Key;
                for (int d = 0; d < entry.Value.Count; d++) data[r, d + 1] = entry.Value[d];
                if (card != null && entry.Key.Equals(card.Trim(), StringComparison.OrdinalIgnoreCase))
                    jumpRow = TsRefHeaderRow + 1 + r;
                r++;
            }

            if (byCard.Count > 0)
            {
                dynamic block = sheet.Range(
                    sheet.Cells(TsRefHeaderRow + 1, 1),
                    sheet.Cells(TsRefHeaderRow + byCard.Count, maxDescriptors + 1));
                block.NumberFormat = "@";
                SwpBridgeSpineSheet.WriteBlock(sheet, TsRefHeaderRow + 1, 1, data);
                sheet.Range(sheet.Cells(TsRefHeaderRow + 1, 1), sheet.Cells(TsRefHeaderRow + byCard.Count, 1))
                    .Font.Bold = true;
            }

            try
            {
                sheet.Columns(1).ColumnWidth = 10;
                for (int c = 2; c <= maxDescriptors + 1; c++)
                {
                    dynamic column = sheet.Columns(c);
                    column.AutoFit();
                    if ((double)column.ColumnWidth > 55) column.ColumnWidth = 55;
                }
            }
            catch { }

            try
            {
                sheet.Activate();
                if (jumpRow > 0) ExcelApp.App.Goto(sheet.Cells(jumpRow, 1), true);
            }
            catch { }
        }

        // ---- Home ---------------------------------------------------------------------------

        public static void RenderHome(dynamic workbook, SwpBridgeSession session)
        {
            dynamic sheet = EnsureSheet(workbook, SwpBridgeWorkspace.HomeSheetName);
            sheet.Cells.Clear();
            try { sheet.Buttons().Delete(); } catch { }
            try { sheet.Move(workbook.Worksheets(1)); } catch { }   // first tab

            var lines = new (int Row, int Col, string Text, bool Bold)[]
            {
                (1, 1, "SWP Bridge", true),
                (2, 1, session.Settings.SwpFileName + " · " + session.Settings.PlantType +
                       " · version " + session.Settings.CurrentVersionFolder +
                       " · " + (session.Settings.ApplyToVersionFolder
                           ? "apply creates a NEW dated version folder"
                           : "apply writes IN PLACE (with .bak)"), false),

                (8, 1, "The workflow", true),
                (9, 1, "1.  The Workspace sheet holds one row per diagnostic mnemonic: the IO-list side, the computed SWP coordinate,", false),
                (10, 1, "     the SWP's current values and — only where something would change — the proposed values (▸ columns).", false),
                (11, 1, "2.  Review: green = up to date, red = did not resolve, orange = FU/Location mismatch; the Plant_Inputs / Busfield / Alarm strings columns carry the coherence verdicts.", false),
                (12, 1, "     Blue-bordered cells react to a double-click: Verdict/Coordinate preview the SWP block, Row opens the IO list at that row, SWP row opens the SWP there.", false),
                (13, 1, "3.  Steer with the [bracketed] override columns (they persist per mnemonic): mark a FU/Location mismatch", false),
                (14, 1, "     deliberate ([Mismatch OK] silences its warning), reroute the DB index, force device ids onto belts, or override a description outright.", false),
                (15, 1, "4.  The _Queue lists every pending write — double-click Include (☑/☐) to keep one out of the apply.", false),
                (16, 1, "     Broken formulas (pasted values over H/I, lost header locations) arrive there too, as repairs.", false),
                (17, 1, "5.  Apply patches the SWP file headlessly: three-way checked, verified before written, atomic with .bak;", false),
                (18, 1, "     the workspace then rebuilds to show the new reality.", false),

                (19, 1, "Colors", true),
                (20, 1, "yellow ▸ cell = pending write · red = conflict (the file moved) · red T.S. ref = missing (mandatory for IO-list alarms) · blue = already there · orange headers = SWP side · blue borders = double-clickable", false),

                (22, 1, "Sheets", true),
                (23, 1, "_Queue the pending writes · _Log every finding with jump links · _Dictionary description find→replace", false),
                (24, 1, "_Glossary the session-relevant translation pairs (disable ☑/☐, edit, add) — the glossary drives Description 2, the bus-failure texts and the T.S. Ref cards · _TSRef the card lookup (double-click a T.S. ref cell)", false),
                (25, 1, "_Offsets the index rules per plant type and object — mnemonic → table entry → exported OPC tag (global, seeded from the official rules v93) · _StrobeTemplates the UPS strobe blocks", false),
                (26, 1, "_Preview the SWP block of the last double-clicked row (toggle \"show full block\" for every pending update)", false),

                (28, 1, "Nothing touches the SWP until Apply. Decisions live in swp-bridge.json at the system root.", false),
            };
            foreach (var line in lines)
            {
                sheet.Cells(line.Row, line.Col).Value = line.Text;
                if (line.Bold) sheet.Cells(line.Row, line.Col).Font.Bold = true;
            }
            sheet.Cells(1, 1).Font.Size = 16;

            // Buttons lay out in fixed pixel slots from the left margin — anchoring them to
            // cells made them overlap as soon as the column widths disagreed.
            AddButton(sheet, 4, 0, "Refresh", "PT_SwpBridgeRefresh");
            AddButton(sheet, 4, 1, "Apply queue", "PT_SwpBridgeApply");
            AddButton(sheet, 4, 2, "Preview block", "PT_SwpBridgePreview");
            AddButton(sheet, 4, 3, "Dashboard", "PT_SwpBridgeDashboard");
            AddButton(sheet, 6, 0, "Clear history", "PT_SwpBridgeClearHistory");
            AddButton(sheet, 6, 1, "Setup…", "PT_SwpBridgeOpen");
        }

        private const double ButtonWidth = 118.0;
        private const double ButtonGap = 10.0;

        private static void AddButton(dynamic sheet, int anchorRow, int slot, string caption, string macro)
        {
            try
            {
                dynamic anchor = sheet.Cells(anchorRow, 1);
                dynamic button = sheet.Buttons().Add(
                    4.0 + slot * (ButtonWidth + ButtonGap), (double)anchor.Top + 2, ButtonWidth, 26.0);
                button.Caption = caption;
                button.OnAction = macro;
            }
            catch { }
        }

        // ---- _StrobeTemplates ---------------------------------------------------------------

        public static void RenderStrobeTemplates(dynamic workbook, SwpBridgeSession session)
        {
            dynamic sheet = EnsureSheet(workbook, SwpBridgeWorkspace.StrobeSheetName);
            sheet.Cells.Clear();

            sheet.Cells(1, 1).Value =
                "Strobe device templates · \">Name | type id\" opens a block, mnemonic/description " +
                "rows follow (placeholders %sys% %belt% %cabinet%), \"/\" closes it · saved " +
                "immediately, applied on the next refresh";
            sheet.Cells(1, 1).Font.Bold = true;

            var lines = new List<(string A, string B)>();
            foreach (SwpStrobeTemplate template in session.Sidecar.StrobeTemplates)
            {
                lines.Add((">" + template.TypeDescription, template.TypeId));
                for (int i = 0; i < template.Mnemonics.Count; i++)
                    lines.Add((template.Mnemonics[i], i < template.Descriptions.Count ? template.Descriptions[i] : ""));
                lines.Add(("/", ""));
            }

            if (lines.Count > 0)
            {
                var data = new object[lines.Count, 2];
                for (int r = 0; r < lines.Count; r++)
                {
                    data[r, 0] = lines[r].A;
                    data[r, 1] = lines[r].B;
                }
                dynamic block = sheet.Range(sheet.Cells(StrobeFirstRow, 1), sheet.Cells(StrobeFirstRow + lines.Count - 1, 2));
                block.NumberFormat = "@";
                SwpBridgeSpineSheet.WriteBlock(sheet, StrobeFirstRow, 1, data);
            }

            try
            {
                sheet.Range(sheet.Cells(StrobeFirstRow, 1),
                    sheet.Cells(StrobeFirstRow + Math.Max(1, lines.Count), 2)).Columns.AutoFit();
            }
            catch { }
        }

        /// <summary>Reads the strobe-template sheet back into the sidecar (the legacy block shape).</summary>
        public static bool ParseStrobeTemplates(dynamic sheet, SwpBridgeSession session)
        {
            try
            {
                long lastRow = RangeHelpers.LastRow(sheet);
                var templates = new List<SwpStrobeTemplate>();
                SwpStrobeTemplate current = null;
                for (int r = StrobeFirstRow; r <= lastRow; r++)
                {
                    string a = CellText(sheet, r, 1);
                    string b = CellText(sheet, r, 2);
                    if (a.StartsWith(">", StringComparison.Ordinal))
                    {
                        current = new SwpStrobeTemplate { TypeDescription = a.Substring(1).Trim(), TypeId = b.Trim() };
                        templates.Add(current);
                    }
                    else if (a == "/")
                    {
                        current = null;
                    }
                    else if (current != null && a.Length > 0)
                    {
                        current.Mnemonics.Add(a);
                        current.Descriptions.Add(b);
                    }
                }
                session.Sidecar.StrobeTemplates = templates;
                return true;
            }
            catch { return false; }
        }

        // ---- shared -------------------------------------------------------------------------

        internal static dynamic EnsureSheet(dynamic workbook, string name)
        {
            dynamic sheet = SwpBridgeQueueSheet.FindSheet(workbook, name);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add(Type.Missing, workbook.Worksheets(workbook.Worksheets.Count));
                sheet.Name = name;
            }
            return sheet;
        }

        private static string CellText(dynamic sheet, int row, int column)
        {
            object raw = sheet.Cells(row, column).Value2;
            return ((string)CsvSheet.CellText(raw, sheet, row, column)).Trim();
        }
    }
}
