using System.Collections.Generic;
using System.Text;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>
    /// RFC 4180 reader: handles quoted fields, embedded delimiters, embedded newlines, and doubled
    /// quotes. It never coerces values — every field stays a string, and each record remembers the
    /// exact line terminator that followed it so saves stay byte-identical.
    /// A quote only opens a quoted field at the start of a field; elsewhere it is a literal character
    /// (matching Excel), so a stray quote never swallows the real delimiter.
    /// Comment lines are NOT special here — a "#" row parses into fields like any other row.
    /// </summary>
    public static class CsvParser
    {
        public static List<CsvRecord> Parse(string text, CsvFormat format)
        {
            var records = new List<CsvRecord>();
            if (string.IsNullOrEmpty(text)) return records;

            int i = 0, n = text.Length;
            char delim = format.Delimiter;
            char quote = format.Quote;

            while (i < n)
            {
                var fields = new List<string>();
                var sb = new StringBuilder();
                bool inQuotes = false;
                bool fieldStart = true;   // a quote only opens a quoted field while this is true
                bool ended = false;
                string terminator = "";

                while (i < n && !ended)
                {
                    char c = text[i];
                    if (inQuotes)
                    {
                        if (c == quote)
                        {
                            if (i + 1 < n && text[i + 1] == quote) { sb.Append(quote); i += 2; }  // escaped quote
                            else { inQuotes = false; i++; }
                        }
                        else { sb.Append(c); i++; }
                    }
                    else if (c == quote)
                    {
                        if (fieldStart) { inQuotes = true; fieldStart = false; i++; }
                        else { sb.Append(quote); i++; }   // literal quote mid-field (Excel-compatible)
                    }
                    else if (c == delim)
                    {
                        fields.Add(sb.ToString());
                        sb.Length = 0;
                        fieldStart = true;
                        i++;
                    }
                    else if (c == '\r' || c == '\n')
                    {
                        int afterEol = SkipEol(text, i);
                        terminator = text.Substring(i, afterEol - i);
                        i = afterEol;
                        ended = true;
                    }
                    else { sb.Append(c); fieldStart = false; i++; }
                }

                fields.Add(sb.ToString());
                var rec = CsvRecord.Data(fields.ToArray());
                rec.Terminator = terminator;   // "" when the record ended at EOF with no newline
                records.Add(rec);
            }

            return records;
        }

        private static int SkipEol(string s, int i)
        {
            if (i < s.Length && s[i] == '\r') { i++; if (i < s.Length && s[i] == '\n') i++; }
            else if (i < s.Length && s[i] == '\n') i++;
            return i;
        }
    }
}
