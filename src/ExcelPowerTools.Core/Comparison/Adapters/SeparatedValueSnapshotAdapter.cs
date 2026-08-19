using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Csv;

namespace ExcelPowerTools.Core.Comparison.Adapters
{
    /// <summary>Result of lossless separated-value parsing plus a sparse positional comparison snapshot.</summary>
    public sealed class SeparatedValueSnapshot
    {
        internal SeparatedValueSnapshot(
            RevisionSnapshot snapshot,
            string byteContentId,
            CsvFormat format,
            int recordCount,
            int maxFieldCount,
            int headerCandidateRecordIndex)
        {
            Snapshot = snapshot;
            ByteContentId = byteContentId;
            Format = format;
            RecordCount = recordCount;
            MaxFieldCount = maxFieldCount;
            HeaderCandidateRecordIndex = headerCandidateRecordIndex;
        }

        public RevisionSnapshot Snapshot { get; }
        public string ByteContentId { get; }
        public CsvFormat Format { get; }
        public int RecordCount { get; }
        public int MaxFieldCount { get; }

        /// <summary>
        /// Raw legacy detector candidate only. It is never accepted as a comparison header until the future
        /// scored/confirmed header mapper classifies it strongly; weak detection therefore remains ordinary data.
        /// </summary>
        public int HeaderCandidateRecordIndex { get; }
    }

    /// <summary>
    /// First-phase adapter for CSV/TSV/separated text. It retains every parsed field as exact text, including
    /// comments, headers, empty fields, and formula-looking strings. Logical row alignment is a later stage;
    /// this adapter intentionally emits explicit positional subject keys.
    /// </summary>
    public static class SeparatedValueSnapshotAdapter
    {
        private const string LogicalSheetKey = "delimited-data";

        public static SeparatedValueSnapshot FromBytes(
            string revisionId,
            byte[] exactSourceBytes,
            string sourcePath = null)
        {
            if (string.IsNullOrEmpty(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (exactSourceBytes == null) throw new ArgumentNullException(nameof(exactSourceBytes));

            EnsureLosslessTextDecode(exactSourceBytes);

            CsvDocument document = sourcePath == null
                ? CsvFile.Load(exactSourceBytes)
                : CsvFile.LoadForPath(exactSourceBytes, sourcePath);

            var cells = new List<CanonicalCellState>();
            int maxFieldCount = 0;
            for (int row = 0; row < document.Records.Count; row++)
            {
                string[] fields = document.Records[row].Fields;
                if (fields.Length > maxFieldCount) maxFieldCount = fields.Length;
                for (int column = 0; column < fields.Length; column++)
                {
                    cells.Add(CanonicalCellState.Constant(
                        PositionalSubjectKey(row + 1, column + 1),
                        CanonicalValue.Text(fields[column] ?? "")));
                }
            }

            return new SeparatedValueSnapshot(
                new RevisionSnapshot(revisionId, cells),
                StableId.ComputeByteContentId(exactSourceBytes),
                document.Format.Clone(),
                document.Records.Count,
                maxFieldCount,
                document.HeaderRecordIndex);
        }

        public static string PositionalSubjectKey(int oneBasedRow, int oneBasedColumn)
        {
            if (oneBasedRow <= 0) throw new ArgumentOutOfRangeException(nameof(oneBasedRow));
            if (oneBasedColumn <= 0) throw new ArgumentOutOfRangeException(nameof(oneBasedColumn));

            return LogicalSheetKey +
                   "/r:" + oneBasedRow.ToString("D10", CultureInfo.InvariantCulture) +
                   "/c:" + oneBasedColumn.ToString("D10", CultureInfo.InvariantCulture);
        }

        private static void EnsureLosslessTextDecode(byte[] bytes)
        {
            (Encoding encoding, bool hasBom) = CsvEncoding.Detect(bytes);
            try
            {
                switch (encoding.CodePage)
                {
                    case 65001:
                        new UTF8Encoding(false, true).GetString(bytes);
                        break;
                    case 1200:
                        new UnicodeEncoding(false, hasBom, true).GetString(bytes);
                        break;
                    case 1201:
                        new UnicodeEncoding(true, hasBom, true).GetString(bytes);
                        break;
                }
            }
            catch (DecoderFallbackException exception)
            {
                throw new InvalidDataException(
                    "The separated-value source declares or strongly indicates Unicode text, but contains " +
                    "an invalid byte sequence. Exact cell comparison stopped rather than replacing data.",
                    exception);
            }
        }
    }
}
