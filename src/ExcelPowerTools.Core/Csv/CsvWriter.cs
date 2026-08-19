using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>
    /// RFC 4180 writer. Quotes minimally by default (only fields that need it) and reproduces each
    /// record's own line terminator (falling back to the format default for programmatic records).
    /// A "#" comment row is written exactly like any other row — fields joined by the delimiter.
    /// </summary>
    public static class CsvWriter
    {
        public static string Write(IReadOnlyList<CsvRecord> records, CsvFormat format)
        {
            var sb = new StringBuilder();
            if (records == null) return string.Empty;

            for (int r = 0; r < records.Count; r++)
            {
                CsvRecord rec = records[r];
                string[] fields = rec.Fields ?? Array.Empty<string>();
                for (int f = 0; f < fields.Length; f++)
                {
                    if (f > 0) sb.Append(format.Delimiter);
                    // A final one-field blank record with no trailing newline would otherwise serialize to
                    // zero bytes and disappear when parsed. Quoting that field is the canonical unambiguous
                    // representation (e.g. a\n\"\" instead of a\n).
                    bool preserveFinalBlankRecord =
                        r == records.Count - 1 && !format.HasTrailingNewline && fields.Length == 1 &&
                        string.IsNullOrEmpty(fields[f]);
                    AppendField(sb, fields[f], format, preserveFinalBlankRecord);
                }

                sb.Append(TerminatorFor(rec, format, isLast: r == records.Count - 1));
            }

            return sb.ToString();
        }

        private static string TerminatorFor(CsvRecord rec, CsvFormat format, bool isLast)
        {
            if (rec.Terminator != null) return rec.Terminator;   // exact terminator remembered from load
            if (isLast) return format.HasTrailingNewline ? format.NewlineString : string.Empty;
            return format.NewlineString;
        }

        private static void AppendField(StringBuilder sb, string field, CsvFormat format, bool forceQuote = false)
        {
            field = field ?? string.Empty;

            bool mustQuote =
                forceQuote
                || format.QuotePolicy == CsvQuotePolicy.All
                || field.IndexOf(format.Delimiter) >= 0
                || field.IndexOf(format.Quote) >= 0
                || field.IndexOf('\r') >= 0
                || field.IndexOf('\n') >= 0;

            if (!mustQuote)
            {
                sb.Append(field);
                return;
            }

            string q = format.Quote.ToString();
            sb.Append(format.Quote);
            sb.Append(field.Replace(q, q + q));   // double embedded quotes
            sb.Append(format.Quote);
        }
    }
}
