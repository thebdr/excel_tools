using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>One glossary pair: the English pivot phrase and its chosen-language text.</summary>
    public sealed class SwpGlossaryPair
    {
        public string En = "";
        public string Translation = "";
        public bool Enabled = true;
        public string Source = "";    // "standard" / "shadow" / "custom"

        /// <summary>The official troubleshooting card ("TS_001") — the standard glossario's
        /// "Numero scheda" column; "" when the file offers none for this phrase.</summary>
        public string TsRef = "";
    }

    /// <summary>
    /// A user decision about one pair, persisted in the sidecar keyed by the EN text: disabled,
    /// a corrected translation, or a wholly user-added pair (one whose EN the files don't carry).
    /// </summary>
    public sealed class SwpGlossaryTweak
    {
        public string En = "";
        public bool Enabled = true;
        public string Translation = "";   // "" = keep the file's text
    }

    /// <summary>
    /// The glossary machinery: parses the Fives glossario workbooks (any sheet whose header row
    /// carries "EN" plus other two-letter language codes — both the Troubleshooting layout and
    /// the ombra layout qualify), merges them (the second/shadow file wins on a shared EN
    /// phrase), and applies the user's tweaks. Consumption is FULL-PHRASE only: a row uses a
    /// pair when its whole Description permanent part equals the EN phrase — substring
    /// translation was retired (a partial-phrase splice produced wrong sentences).
    /// </summary>
    public static class SwpGlossary
    {
        // ---- parsing ------------------------------------------------------------------------

        /// <summary>
        /// All language codes any sheet of the workbook offers (deduplicated, EN excluded).
        /// </summary>
        public static List<string> AvailableLanguages(byte[] xlsxBytes)
        {
            var codes = new List<string>();
            try
            {
                using (var stream = new MemoryStream(xlsxBytes, writable: false))
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
                {
                    IReadOnlyList<string> sharedStrings = LoadSharedStrings(document.WorkbookPart);
                    foreach (var location in FindGlossarySheets(document, sharedStrings))
                        foreach (string code in location.Languages.Keys)
                            if (code != "EN" && !codes.Contains(code)) codes.Add(code);
                }
            }
            catch { }
            return codes;
        }

        /// <summary>
        /// EN → chosen-language pairs from every glossary-shaped sheet of the workbook. Returns
        /// an empty list (with a log line) when no sheet offers the language.
        /// </summary>
        public static List<SwpGlossaryPair> Parse(byte[] xlsxBytes, string languageCode,
            string sourceLabel, List<SwpLogEntry> log)
        {
            var pairs = new List<SwpGlossaryPair>();
            if (xlsxBytes == null || xlsxBytes.Length == 0) return pairs;
            languageCode = (languageCode ?? "").Trim().ToUpperInvariant();
            if (languageCode.Length == 0) return pairs;

            try
            {
                using (var stream = new MemoryStream(xlsxBytes, writable: false))
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
                {
                    IReadOnlyList<string> sharedStrings = LoadSharedStrings(document.WorkbookPart);
                    bool anySheet = false, anyLanguage = false;
                    foreach (GlossarySheet location in FindGlossarySheets(document, sharedStrings))
                    {
                        anySheet = true;
                        if (!location.Languages.TryGetValue(languageCode, out int languageColumn)) continue;
                        anyLanguage = true;
                        ReadPairs(location, languageColumn, sourceLabel, sharedStrings, pairs);
                    }
                    if (!anySheet)
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "the " + sourceLabel + " glossary has no sheet with an EN + language-code header row"));
                    else if (!anyLanguage)
                        log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                            "the " + sourceLabel + " glossary has no " + languageCode + " column"));
                }
            }
            catch (Exception ex)
            {
                log?.Add(SwpLogEntry.For(SwpLogSeverity.Warning,
                    "the " + sourceLabel + " glossary could not be read: " + ex.Message));
            }
            return pairs;
        }

        private sealed class GlossarySheet
        {
            public WorksheetPart Part; // features only — content comes from the copy below
            public MemoryStream Xml;   // the part's raw XML, copied out of the package
            public int HeaderRow;
            public int EnColumn;
            public int TsColumn;       // "Numero scheda" (TS_xxx) — 0 when the sheet has none
            public Dictionary<string, int> Languages = new Dictionary<string, int>();  // code → column
        }

        /// <summary>
        /// The troubleshooting-card column header. The standard glossario says "Numero scheda";
        /// the OMBRA's bare "scheda" column holds template TEXTS, not cards, and must NOT match.
        /// </summary>
        private static bool IsTsHeader(string upperText) =>
            upperText.Contains("NUMERO SCHEDA") ||
            upperText == "TS" || upperText == "TS REF" || upperText == "TS_REF" || upperText == "T.S. REF";

        /// <summary>A typed reader over copied part XML — the part supplies only its features.</summary>
        private static OpenXmlReader ReaderOver(WorksheetPart part, MemoryStream copy) =>
            new OpenXmlPartReader(new MemoryStream(copy.GetBuffer(), 0, (int)copy.Length, writable: false),
                part.Features, new OpenXmlPartReaderOptions());

        private static IEnumerable<GlossarySheet> FindGlossarySheets(SpreadsheetDocument document,
            IReadOnlyList<string> sharedStrings)
        {
            WorkbookPart workbook = document.WorkbookPart;

            foreach (Sheet sheet in workbook.Workbook.Sheets.Elements<Sheet>())
            {
                if (!(workbook.GetPartById(sheet.Id.Value) is WorksheetPart part)) continue;

                // Copied out in one deterministic pass — partially-consumed package streams
                // leave WindowsBase (net48) zip state that detonates at package dispose.
                MemoryStream xml;
                try
                {
                    using (Stream source = part.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        xml = new MemoryStream();
                        source.CopyTo(xml);
                    }
                }
                catch
                {
                    continue;   // an unreadable sheet is simply not a glossary sheet
                }

                GlossarySheet found = null;
                using (OpenXmlReader reader = ReaderOver(part, xml))
                {
                    int scanned = 0;
                    while (reader.Read() && scanned < 5)   // the header sits in the first rows
                    {
                        if (reader.ElementType != typeof(Row) || !reader.IsStartElement) continue;
                        var row = (Row)reader.LoadCurrentElement();
                        scanned++;

                        var candidate = new GlossarySheet
                        {
                            Part = part,
                            Xml = xml,
                            HeaderRow = row.RowIndex != null ? (int)row.RowIndex.Value : scanned,
                        };
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string text = CellText(cell, sharedStrings).Trim().ToUpperInvariant();
                            int column = ColumnOf(cell.CellReference?.Value);
                            if (IsTsHeader(text))
                            {
                                if (candidate.TsColumn == 0) candidate.TsColumn = column;
                                continue;
                            }
                            if (text.Length != 2 || !char.IsLetter(text[0]) || !char.IsLetter(text[1])) continue;
                            if (text == "EN") candidate.EnColumn = column;
                            else if (!candidate.Languages.ContainsKey(text)) candidate.Languages[text] = column;
                        }
                        if (candidate.EnColumn > 0 && candidate.Languages.Count > 0)
                        {
                            found = candidate;
                            break;
                        }
                    }
                }
                if (found != null) yield return found;
            }
        }

        private static void ReadPairs(GlossarySheet location, int languageColumn,
            string sourceLabel, IReadOnlyList<string> sharedStrings, List<SwpGlossaryPair> pairs)
        {
            using (OpenXmlReader reader = ReaderOver(location.Part, location.Xml))
            {
                while (reader.Read())
                {
                    if (reader.ElementType != typeof(Row) || !reader.IsStartElement) continue;
                    var row = (Row)reader.LoadCurrentElement();
                    if (row.RowIndex == null || (int)row.RowIndex.Value <= location.HeaderRow) continue;

                    string en = null, translation = null, tsRef = null;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        int column = ColumnOf(cell.CellReference?.Value);
                        if (column == location.EnColumn) en = CellText(cell, sharedStrings).Trim();
                        else if (column == languageColumn) translation = CellText(cell, sharedStrings).Trim();
                        else if (location.TsColumn > 0 && column == location.TsColumn)
                            tsRef = CellText(cell, sharedStrings).Trim();
                    }
                    if (!string.IsNullOrEmpty(en) && !string.IsNullOrEmpty(translation))
                        pairs.Add(new SwpGlossaryPair
                        {
                            En = en,
                            Translation = translation,
                            Source = sourceLabel,
                            TsRef = tsRef ?? "",
                        });
                }
            }
        }

        // ---- merging + tweaks ---------------------------------------------------------------

        /// <summary>
        /// Standard pairs first, shadow pairs override by EN phrase, then the user's tweaks:
        /// enable/disable and corrected translations by EN key; a tweak whose EN no file carries
        /// becomes a "custom" pair of its own.
        /// </summary>
        public static List<SwpGlossaryPair> Merge(IReadOnlyList<SwpGlossaryPair> standard,
            IReadOnlyList<SwpGlossaryPair> shadow, IReadOnlyList<SwpGlossaryTweak> tweaks)
        {
            var byEn = new Dictionary<string, SwpGlossaryPair>(StringComparer.OrdinalIgnoreCase);
            var ordered = new List<SwpGlossaryPair>();

            void Take(IReadOnlyList<SwpGlossaryPair> source)
            {
                if (source == null) return;
                foreach (SwpGlossaryPair pair in source)
                {
                    if (byEn.TryGetValue(pair.En, out SwpGlossaryPair existing))
                    {
                        existing.Translation = pair.Translation;   // the later file wins
                        existing.Source = pair.Source;
                        // The shadow file has no card column — never lose the standard's TS ref.
                        if (pair.TsRef.Length > 0) existing.TsRef = pair.TsRef;
                    }
                    else
                    {
                        var copy = new SwpGlossaryPair
                        {
                            En = pair.En,
                            Translation = pair.Translation,
                            Source = pair.Source,
                            TsRef = pair.TsRef,
                        };
                        byEn[copy.En] = copy;
                        ordered.Add(copy);
                    }
                }
            }
            Take(standard);
            Take(shadow);

            foreach (SwpGlossaryTweak tweak in tweaks ?? new List<SwpGlossaryTweak>())
            {
                if (string.IsNullOrEmpty(tweak.En)) continue;
                if (byEn.TryGetValue(tweak.En, out SwpGlossaryPair pair))
                {
                    pair.Enabled = tweak.Enabled;
                    if (tweak.Translation.Length > 0) pair.Translation = tweak.Translation;
                }
                else if (tweak.Translation.Length > 0)
                {
                    var custom = new SwpGlossaryPair
                    {
                        En = tweak.En,
                        Translation = tweak.Translation,
                        Enabled = tweak.Enabled,
                        Source = "custom",
                    };
                    byEn[custom.En] = custom;
                    ordered.Add(custom);
                }
            }
            return ordered;
        }

        // ---- relevance ----------------------------------------------------------------------

        /// <summary>
        /// The pairs whose EN phrase IS one of the given glossary keys (a row's full Description
        /// permanent part) — matching is whole-phrase, trimmed, case-insensitive; substring
        /// matching was retired (a partial-phrase splice produced wrong sentences).
        /// </summary>
        public static List<SwpGlossaryPair> RelevantPairs(IReadOnlyList<SwpGlossaryPair> pairs,
            IEnumerable<string> glossaryKeys, out Dictionary<string, int> occurrences)
        {
            occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var relevant = new List<SwpGlossaryPair>();
            if (pairs == null) return relevant;

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string key in glossaryKeys ?? new string[0])
            {
                string trimmed = (key ?? "").Trim();
                if (trimmed.Length == 0) continue;
                counts.TryGetValue(trimmed, out int n);
                counts[trimmed] = n + 1;
            }

            foreach (SwpGlossaryPair pair in pairs)
            {
                counts.TryGetValue(pair.En.Trim(), out int count);
                if (count > 0 || pair.Source == "custom")
                {
                    occurrences[pair.En] = count;
                    relevant.Add(pair);
                }
            }
            return relevant;
        }

        // ---- shared helpers -----------------------------------------------------------------

        private static IReadOnlyList<string> LoadSharedStrings(WorkbookPart workbook)
        {
            var strings = new List<string>();
            SharedStringTable table = workbook.SharedStringTablePart?.SharedStringTable;
            if (table == null) return strings;
            foreach (OpenXmlElement item in table.Elements())
                strings.Add(item.InnerText ?? "");
            return strings;
        }

        private static string CellText(Cell cell, IReadOnlyList<string> sharedStrings)
        {
            if (cell == null) return "";
            string raw = cell.CellValue?.Text ?? "";
            if (cell.DataType == null) return raw;
            if (cell.DataType.Value == CellValues.SharedString)
            {
                return int.TryParse(raw, out int index) && index >= 0 && index < sharedStrings.Count
                    ? sharedStrings[index]
                    : "";
            }
            if (cell.DataType.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? cell.InnerText ?? "";
            return raw;
        }

        private static int ColumnOf(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            int column = 0;
            foreach (char ch in cellReference)
            {
                if (ch >= 'A' && ch <= 'Z') column = column * 26 + (ch - 'A' + 1);
                else break;
            }
            return column;
        }
    }
}
