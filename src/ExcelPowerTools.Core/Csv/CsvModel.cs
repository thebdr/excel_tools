using System;
using System.Text;

namespace ExcelPowerTools.Core.Csv
{
    public enum CsvLineEnding { Crlf, Lf, Cr }

    public enum CsvQuotePolicy { Minimal, All }

    /// <summary>
    /// The full "format fingerprint" of a CSV file: everything needed to round-trip it byte-for-byte.
    /// Detected on load, carried on the document, and reapplied on save so untouched files diff cleanly.
    /// </summary>
    public sealed class CsvFormat
    {
        public char Delimiter { get; set; } = ';';
        public char Quote { get; set; } = '"';
        public CsvQuotePolicy QuotePolicy { get; set; } = CsvQuotePolicy.Minimal;

        /// <summary>Text encoding to read/write with.</summary>
        public Encoding Encoding { get; set; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        /// <summary>Whether the source had a byte-order mark (re-emitted on save when true).</summary>
        public bool HasBom { get; set; }

        /// <summary>The dominant line ending; used as the default terminator for records added programmatically.</summary>
        public CsvLineEnding LineEnding { get; set; } = CsvLineEnding.Crlf;

        /// <summary>Default trailing-newline state for programmatic records (loaded records carry their own terminator).</summary>
        public bool HasTrailingNewline { get; set; } = true;

        /// <summary>
        /// The source began with an Excel "sep=&lt;char&gt;" directive line. The line itself is not a record:
        /// it lives here in the fingerprint and is re-emitted (with <see cref="SepTerminator"/>) on save.
        /// </summary>
        public bool HasSepDirective { get; set; }

        /// <summary>Exact terminator that followed the sep= directive line (null → the format default).</summary>
        public string SepTerminator { get; set; }

        /// <summary>
        /// Rows whose FIRST field starts with this prefix are comment rows ("" disables). They are ordinary
        /// records — split into fields and written back like any other row — but detection (delimiter,
        /// header, ragged) skips them and the Excel layer styles them green.
        /// </summary>
        public string CommentPrefix { get; set; } = "#";

        public string NewlineString =>
            LineEnding == CsvLineEnding.Crlf ? "\r\n" :
            LineEnding == CsvLineEnding.Lf ? "\n" : "\r";

        public CsvFormat Clone() => (CsvFormat)MemberwiseClone();
    }

    /// <summary>One parsed record: its fields, plus the exact line terminator that followed it in the source.</summary>
    public sealed class CsvRecord
    {
        public string[] Fields { get; }

        /// <summary>
        /// The exact line terminator that followed this record in the source ("\r\n", "\n", "\r", or "" at
        /// end-of-file). Null for records built programmatically, in which case the writer uses the format default.
        /// </summary>
        public string Terminator { get; set; }

        private CsvRecord(string[] fields) => Fields = fields;

        public static CsvRecord Data(string[] fields) => new CsvRecord(fields ?? Array.Empty<string>());

        public int FieldCount => Fields.Length;
    }
}
