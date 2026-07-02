using System.Collections.Generic;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Commands
{
    /// <summary>
    /// Remembers, for each open CSV, the source path and the detected <see cref="CsvFormat"/> fingerprint,
    /// so "Save" writes it back byte-faithfully. Keyed by the workbook's Name (stable while the workbook is
    /// open; sessions live for the Excel session only — reopen the file to re-establish one).
    /// </summary>
    internal sealed class CsvSession
    {
        public string SourcePath { get; set; }
        public CsvFormat Format { get; set; }
        public int HeaderSheetRow { get; set; } = -1;   // 1-based sheet row of the header, or -1
        public int DataWidth { get; set; }              // original widest data record (columns)

        private static readonly Dictionary<string, CsvSession> Sessions = new Dictionary<string, CsvSession>();

        public static void Set(string workbookKey, CsvSession session) => Sessions[workbookKey] = session;

        public static CsvSession Get(string workbookKey) =>
            workbookKey != null && Sessions.TryGetValue(workbookKey, out var s) ? s : null;

        public static void Remove(string workbookKey)
        {
            if (workbookKey != null) Sessions.Remove(workbookKey);
        }
    }
}
