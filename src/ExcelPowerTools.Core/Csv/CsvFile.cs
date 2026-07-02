using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelPowerTools.Core.Csv
{
    /// <summary>A loaded CSV: its format fingerprint, its records, and the detected header row.</summary>
    public sealed class CsvDocument
    {
        public CsvFormat Format { get; set; }
        public List<CsvRecord> Records { get; set; } = new List<CsvRecord>();

        /// <summary>Index into <see cref="Records"/> of the detected header, or -1.</summary>
        public int HeaderRecordIndex { get; set; } = -1;

        /// <summary>The data rows, in order — comment rows (first field starting with the comment prefix) excluded.</summary>
        public IReadOnlyList<string[]> DataRows()
        {
            string prefix = Format?.CommentPrefix;
            var rows = new List<string[]>();
            foreach (var r in Records)
                if (!CsvDetect.IsCommentFields(r.Fields, prefix)) rows.Add(r.Fields);
            return rows;
        }

        public byte[] Save() => CsvFile.Save(this);
    }

    /// <summary>Byte-level load/save that detects and round-trips the full <see cref="CsvFormat"/> fingerprint.</summary>
    public static class CsvFile
    {
        public static CsvDocument Load(byte[] bytes) => Load(bytes, null);

        public static CsvDocument Load(byte[] bytes, CsvFormat hint)
        {
            var fmt = hint != null ? hint.Clone() : new CsvFormat();

            var (encoding, hasBom) = CsvEncoding.Detect(bytes);
            fmt.Encoding = encoding;
            fmt.HasBom = hasBom;

            string text = CsvEncoding.DecodeText(bytes, encoding, hasBom);

            fmt.LineEnding = CsvDetect.DetectLineEnding(text);
            fmt.HasTrailingNewline = CsvDetect.HasTrailingNewline(text);

            // A leading Excel "sep=<char>" directive is authoritative for the delimiter. It is part of the
            // fingerprint (not a record): stripped here, re-emitted with its own terminator on save.
            // A directive whose delimiter would be unusable (the quote char, or a line break) is ignored.
            string firstLine = ReadFirstLine(text, out int afterFirst);
            if (CsvDetect.TryParseSepDirective(firstLine, out char sepChar)
                && sepChar != fmt.Quote && sepChar != '\r' && sepChar != '\n')
            {
                fmt.HasSepDirective = true;
                fmt.Delimiter = sepChar;
                fmt.SepTerminator = text.Substring(firstLine.Length, afterFirst - firstLine.Length);
                text = text.Substring(afterFirst);
            }
            else if (hint == null)
            {
                string sample = text.Length > 65536 ? text.Substring(0, 65536) : text;
                fmt.Delimiter = CsvDetect.DetectDelimiter(sample, CsvDetect.DefaultDelimiters, fmt.Quote, fmt.Delimiter);
            }

            var records = CsvParser.Parse(text, fmt);

            var doc = new CsvDocument { Format = fmt, Records = records };
            doc.HeaderRecordIndex = FindHeaderRecord(records, fmt.CommentPrefix);
            return doc;
        }

        public static byte[] Save(CsvDocument doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            CsvFormat fmt = doc.Format ?? new CsvFormat();

            string text = CsvWriter.Write(doc.Records, fmt);
            if (fmt.HasSepDirective)
                text = "sep=" + fmt.Delimiter + (fmt.SepTerminator ?? fmt.NewlineString) + text;

            Encoding enc = fmt.Encoding ?? new UTF8Encoding(false);
            byte[] body = enc.GetBytes(text);

            byte[] bom = CsvEncoding.Preamble(enc, fmt.HasBom);
            if (bom.Length == 0) return body;

            var outBytes = new byte[bom.Length + body.Length];
            Buffer.BlockCopy(bom, 0, outBytes, 0, bom.Length);
            Buffer.BlockCopy(body, 0, outBytes, bom.Length, body.Length);
            return outBytes;
        }

        private static string ReadFirstLine(string text, out int afterFirst)
        {
            int i = 0;
            while (i < text.Length && text[i] != '\r' && text[i] != '\n') i++;
            string line = text.Substring(0, i);

            int j = i;
            if (j < text.Length && text[j] == '\r') { j++; if (j < text.Length && text[j] == '\n') j++; }
            else if (j < text.Length && text[j] == '\n') j++;

            afterFirst = j;
            return line;
        }

        private static int FindHeaderRecord(List<CsvRecord> records, string commentPrefix)
        {
            // All rows are candidates (a commented "#Name;IP;..." header included); the detector itself
            // computes the modal width from data rows only. Index maps 1:1 onto Records.
            var rows = new string[records.Count][];
            for (int i = 0; i < records.Count; i++) rows[i] = records[i].Fields;
            return CsvDetect.DetectHeaderRow(rows, commentPrefix);
        }
    }
}
