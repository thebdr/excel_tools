using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The formula cleaner, ported from the Cleanup VBA. Two machines:
    ///
    /// REPAIRS (self-contained, become queue actions): a healthy bit row's FUnit (H) and
    /// Location (I) are formulas pointing at their own block header — pasted values, deleted
    /// cells and dragged formulas break that and the repair proposes =H{header} / =I{header}.
    /// A block header whose Location lost its "+" prefix is rebuilt from the electrical label,
    /// and a STROBE sheet's first header Location must be =MID(F15,4,12).
    ///
    /// AUDIT (needs a reference SWP, findings only): every target row folds onto the base's
    /// canonical block; a column where the base holds a formula but the target does not is
    /// damage (MISSING_F — the classic paste-over), and an external CONFIG_PLANT reference that
    /// differs from the base's (names normalized) points at the wrong template row.
    /// </summary>
    public static class SwpFormulaCheck
    {
        private const int MaxAuditLogPerSheet = 50;

        // ---- self-contained repairs ---------------------------------------------------------

        /// <summary>
        /// <paramref name="deliberateIdentities"/> — (sheet, row) → the IO list's (FUnit, Location)
        /// of the signal living on that SWP row — marks the plain values brutal-overwrite mode
        /// writes into H/I as DELIBERATE: a value equal to that identity is not a pasted accident
        /// and gets no repair. Null outside overwrite mode.
        /// </summary>
        public static List<SwpProposedEdit> ProposeRepairs(SwpWorkbookIndex index, List<SwpLogEntry> log,
            IReadOnlyDictionary<(string Sheet, int Row), (string FUnit, string Location)> deliberateIdentities = null)
        {
            var edits = new List<SwpProposedEdit>();
            if (index == null) return edits;

            foreach (SwpSheetIndex sheet in index.Sheets.Values)
            {
                if (!sheet.IsDiagnosticSheet) continue;
                bool isStrobe = sheet.SheetName.IndexOf("STROBE", StringComparison.OrdinalIgnoreCase) >= 0;

                foreach (SwpRowSnapshot row in sheet.BitRows)
                {
                    (string FUnit, string Location) identity = ("", "");
                    bool deliberate = deliberateIdentities != null &&
                        deliberateIdentities.TryGetValue((sheet.SheetName, row.RowNumber), out identity);
                    CheckRowFormula(sheet, row, edits, log,
                        SwpField.FormulaH, row.FormulaH, row.FormulaHShared, row.HasCellH,
                        row.CachedH, "H", row.BlockFUnit, deliberate ? identity.FUnit : null);
                    CheckRowFormula(sheet, row, edits, log,
                        SwpField.FormulaI, row.FormulaI, row.FormulaIShared, row.HasCellI,
                        row.CachedI, "I", row.BlockLocation, deliberate ? identity.Location : null);
                }

                foreach (SwpBlockHeader header in sheet.BlockHeaders)
                {
                    if (isStrobe)
                    {
                        // The legacy cleaner pinned only the FIRST strobe header to the MID shape.
                        if (header.RowNumber != sheet.BlockHeaders[0].RowNumber) continue;
                        string wanted = "MID(F" + header.RowNumber.ToString(CultureInfo.InvariantCulture) + ",4,12)";
                        if (NormalizeFormula(header.LocationIFormula) == NormalizeFormula(wanted)) continue;
                        edits.Add(new SwpProposedEdit
                        {
                            Sheet = sheet.SheetName,
                            CoordinateKey = header.Key,
                            Field = SwpField.HeaderFormulaI,
                            Mnemonic = "(formula repair)",
                            Current = header.LocationIFormula,
                            Proposed = wanted,
                            RowNumber = header.RowNumber,
                        });
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "'" + sheet.SheetName + "' row " + header.RowNumber +
                            ": the strobe header Location is not the MID(F…) formula — repair queued"));
                        continue;
                    }

                    // Virtual objects (SORTERSTATE, TRAIN, …) have no location at all — an empty
                    // header cell is their healthy state, not a lost prefix.
                    if (header.LocationI.Trim().Length == 0) continue;
                    if (header.LocationI.TrimStart().StartsWith("+", StringComparison.Ordinal)) continue;

                    int plus = header.LabelF.IndexOf('+');
                    if (plus < 0)
                    {
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "'" + sheet.SheetName + "' row " + header.RowNumber +
                            ": the header Location has no '+' and the label \"" + header.LabelF +
                            "\" offers none to rebuild from"));
                        continue;
                    }
                    edits.Add(new SwpProposedEdit
                    {
                        Sheet = sheet.SheetName,
                        CoordinateKey = header.Key,
                        Field = SwpField.HeaderLocationI,
                        Mnemonic = "(formula repair)",
                        Current = header.LocationI,
                        Proposed = header.LabelF.Substring(plus),
                        RowNumber = header.RowNumber,
                    });
                    log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                        "'" + sheet.SheetName + "' row " + header.RowNumber +
                        ": the header Location \"" + header.LocationI + "\" lost its '+' — repair queued"));
                }
            }
            return edits;
        }

        /// <summary>
        /// A bit row's FUnit/Location cell is healthy when it is the shared template formula or
        /// an unshared formula still pointing at its own block header. Shared with the diff's
        /// mismatch warning, which words itself differently over a broken cell.
        /// </summary>
        public static bool RowFormulaIsHealthy(string formula, bool isShared, int blockHeaderRow,
            string columnLetter)
        {
            if (isShared) return true;
            string wanted = columnLetter + blockHeaderRow.ToString(CultureInfo.InvariantCulture);
            return (formula ?? "").Length > 0 &&
                formula.Replace("$", "").IndexOf(wanted, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void CheckRowFormula(SwpSheetIndex sheet, SwpRowSnapshot row,
            List<SwpProposedEdit> edits, List<SwpLogEntry> log, SwpField field,
            string formula, bool isShared, bool hasCell, string cached,
            string columnLetter, string headerValue, string deliberateValue = null)
        {
            // A block whose header carries no FUnit/Location is a virtual object (SORTERSTATE,
            // TRAIN, …): its rows legitimately hold nothing in that column, and a repair formula
            // would only display a spurious 0 over the empty header cell.
            if ((headerValue ?? "").Trim().Length == 0) return;

            if (RowFormulaIsHealthy(formula, isShared, row.BlockHeaderRow, columnLetter)) return;

            // Brutal-overwrite mode writes the IO identity into the row's own H/I: a plain value
            // equal to that identity is deliberate, not a pasted accident.
            if (!isShared && hasCell && (formula ?? "").Length == 0 && deliberateValue != null &&
                string.Equals((cached ?? "").Trim(), deliberateValue.Trim(), StringComparison.Ordinal))
                return;

            string wanted = columnLetter + row.BlockHeaderRow.ToString(CultureInfo.InvariantCulture);

            edits.Add(new SwpProposedEdit
            {
                Sheet = sheet.SheetName,
                CoordinateKey = row.RowKey,
                Field = field,
                Mnemonic = "(formula repair)",
                Current = formula,
                Proposed = wanted,
                RowNumber = row.RowNumber,
            });
            log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                "'" + sheet.SheetName + "' row " + row.RowNumber + " " + columnLetter + ": " +
                (hasCell
                    ? formula.Length == 0
                        ? "a pasted value (\"" + cached + "\") replaced the formula"
                        : "the formula \"" + formula + "\" does not reference the block header"
                    : "the cell is gone") +
                " — repair to =" + wanted + " queued (header holds \"" + headerValue + "\")"));
        }

        private static string NormalizeFormula(string formula) =>
            (formula ?? "").Replace("$", "").Replace(" ", "").ToUpperInvariant();

        // ---- the base-SWP audit -------------------------------------------------------------

        /// <summary>
        /// Compares the target's formula structure against a reference SWP. Sheets map by exact
        /// name first, then by the part after the first underscore (UP03_BELT audits against the
        /// base's UP01_BELT); rows fold onto the base's canonical block by its block length.
        /// Findings only — the wholesale formula content is the reference file's business.
        /// </summary>
        public static void AuditAgainstBase(SwpWorkbookIndex target, SwpWorkbookIndex baseIndex,
            List<SwpLogEntry> log)
        {
            if (target == null || baseIndex == null || log == null) return;

            foreach (SwpSheetIndex sheet in target.Sheets.Values)
            {
                if (!sheet.IsDiagnosticSheet) continue;
                SwpSheetIndex baseSheet = FindBaseSheet(baseIndex, sheet.SheetName);
                if (baseSheet == null)
                {
                    log.Add(SwpLogEntry.For(SwpLogSeverity.Info,
                        "audit: no reference sheet for '" + sheet.SheetName + "'"));
                    continue;
                }
                AuditSheet(sheet, baseSheet, log);
            }
        }

        private static SwpSheetIndex FindBaseSheet(SwpWorkbookIndex baseIndex, string name)
        {
            if (baseIndex.Sheets.TryGetValue(name, out SwpSheetIndex exact)) return exact;

            int underscore = name.IndexOf('_');
            if (underscore < 0) return null;
            string suffix = name.Substring(underscore);
            foreach (SwpSheetIndex candidate in baseIndex.Sheets.Values)
                if (candidate.SheetName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    return candidate;
            return null;
        }

        private static void AuditSheet(SwpSheetIndex sheet, SwpSheetIndex baseSheet, List<SwpLogEntry> log)
        {
            if (baseSheet.BlockHeaders.Count == 0 || sheet.BlockHeaders.Count == 0) return;

            int baseFirst = baseSheet.BlockHeaders[0].RowNumber;
            int blockLength = baseSheet.BlockHeaders.Count > 1
                ? baseSheet.BlockHeaders[1].RowNumber - baseFirst
                : 166;   // the canonical block

            // The base block's rows by their offset inside the block.
            var baseByOffset = new Dictionary<int, SwpRowSnapshot>();
            foreach (SwpRowSnapshot row in baseSheet.BitRows)
            {
                int offset = (row.RowNumber - baseFirst) % blockLength;
                if (!baseByOffset.ContainsKey(offset)) baseByOffset[offset] = row;
            }

            int targetFirst = sheet.BlockHeaders[0].RowNumber;
            int logged = 0, suppressed = 0;
            foreach (SwpRowSnapshot row in sheet.BitRows)
            {
                int offset = (row.RowNumber - targetFirst) % blockLength;
                if (offset < 0) offset += blockLength;
                if (!baseByOffset.TryGetValue(offset, out SwpRowSnapshot baseRow)) continue;

                // MISSING_F: the base holds a formula in a column, the target does not.
                uint missing = baseRow.FormulaMask & ~row.FormulaMask;
                for (int column = 1; column <= 21 && missing != 0; column++)
                {
                    uint bit = 1u << (column - 1);
                    if ((missing & bit) == 0) continue;
                    missing &= ~bit;
                    Emit(log, ref logged, ref suppressed,
                        "audit '" + sheet.SheetName + "' row " + row.RowNumber + " col " +
                        ColumnLetter(column) + ": the reference has a formula here, this file has " +
                        "a value — the formula was lost");
                }

                // CONFIG_PLANT_REF_DIFF: external template references pointing elsewhere.
                if (row.ExternalFormulas == null || baseRow.ExternalFormulas == null) continue;
                foreach (var pair in row.ExternalFormulas)
                {
                    if (!baseRow.ExternalFormulas.TryGetValue(pair.Key, out string baseFormula)) continue;
                    string normalized = NormalizeExternal(pair.Value, row.RowNumber);
                    string baseNormalized = NormalizeExternal(baseFormula, baseRow.RowNumber);
                    if (normalized == baseNormalized) continue;
                    Emit(log, ref logged, ref suppressed,
                        "audit '" + sheet.SheetName + "' row " + row.RowNumber + " col " +
                        ColumnLetter(pair.Key) + ": the CONFIG_PLANT reference differs from the " +
                        "reference file (\"" + pair.Value + "\" vs \"" + baseFormula + "\")");
                }
            }

            if (suppressed > 0)
                log.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "audit '" + sheet.SheetName + "': " + suppressed + " further finding(s) not listed"));
        }

        /// <summary>
        /// Strips what legitimately differs between base and target: '$' anchors, the BASE_
        /// prefix convention, and the row's own number in local references.
        /// </summary>
        internal static string NormalizeExternal(string formula, int rowNumber)
        {
            string s = (formula ?? "").Replace("$", "").Replace("BASE_CONFIG_PLANT", "CONFIG_PLANT");
            return s.Replace(rowNumber.ToString(CultureInfo.InvariantCulture), "#R").ToUpperInvariant();
        }

        private static void Emit(List<SwpLogEntry> log, ref int logged, ref int suppressed, string message)
        {
            if (logged >= MaxAuditLogPerSheet) { suppressed++; return; }
            logged++;
            log.Add(SwpLogEntry.For(SwpLogSeverity.Warning, message));
        }

        private static string ColumnLetter(int column)
        {
            string letters = "";
            while (column > 0)
            {
                int remainder = (column - 1) % 26;
                letters = (char)('A' + remainder) + letters;
                column = (column - 1) / 26;
            }
            return letters;
        }
    }
}
