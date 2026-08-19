using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    /// <summary>
    /// Immutable tabular evidence used only for inference and alignment. Rows are defensively copied and retain
    /// exact strings, including the distinction between null, empty, whitespace-only, and physically absent
    /// cells. Alignment results contain indices into this evidence; they never rewrite canonical cell content.
    /// </summary>
    public sealed class AlignmentTable
    {
        private readonly IReadOnlyList<AlignmentRow> _rows;

        public AlignmentTable(IEnumerable<IEnumerable<string>> rows)
        {
            var copy = new List<AlignmentRow>();
            int columnCount = 0;
            if (rows != null)
            {
                foreach (IEnumerable<string> row in rows)
                {
                    if (row == null) throw new ArgumentException("Alignment tables cannot contain null rows.", nameof(rows));
                    var values = new List<string>(row);
                    var alignmentRow = new AlignmentRow(values);
                    copy.Add(alignmentRow);
                    if (alignmentRow.Count > columnCount) columnCount = alignmentRow.Count;
                }
            }

            _rows = new ReadOnlyCollection<AlignmentRow>(copy);
            ColumnCount = columnCount;
        }

        public IReadOnlyList<AlignmentRow> Rows => _rows;
        public int RowCount => _rows.Count;
        public int ColumnCount { get; }

        public bool TryGetCell(int rowIndex, int columnIndex, out string exactValue)
        {
            if (rowIndex < 0 || rowIndex >= RowCount) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            return _rows[rowIndex].TryGetValue(columnIndex, out exactValue);
        }

        public string GetCellOrNull(int rowIndex, int columnIndex)
        {
            TryGetCell(rowIndex, columnIndex, out string value);
            return value;
        }
    }

    public sealed class AlignmentRow
    {
        private readonly IReadOnlyList<string> _values;

        internal AlignmentRow(IList<string> values)
        {
            _values = new ReadOnlyCollection<string>(new List<string>(values));
        }

        public IReadOnlyList<string> Values => _values;
        public int Count => _values.Count;

        public bool TryGetValue(int columnIndex, out string exactValue)
        {
            if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            if (columnIndex >= _values.Count)
            {
                exactValue = null;
                return false;
            }
            exactValue = _values[columnIndex];
            return true;
        }
    }

    public enum AlignmentConfidence
    {
        None,
        Weak,
        Medium,
        Strong,
        Confirmed,
    }

    /// <summary>A stable machine-readable reason plus a human-readable explanation.</summary>
    public sealed class AlignmentReason
    {
        public AlignmentReason(string code, string message, double scoreImpact = 0D)
        {
            if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("A reason code is required.", nameof(code));
            Code = code;
            Message = message ?? "";
            ScoreImpact = scoreImpact;
        }

        public string Code { get; }
        public string Message { get; }
        public double ScoreImpact { get; }
    }

    public enum AxisAlignmentKind
    {
        Matched,
        OnlyBefore,
        OnlyAfter,
    }

    public enum AxisAlignmentStrategy
    {
        None,
        ConfirmedKey,
        InferredKey,
        HeaderLabels,
        UniqueContentAnchors,
        NeighborBoundedSequence,
        Positional,
    }

    /// <summary>One resolved one-to-one match, insertion, or deletion on a zero-based physical axis.</summary>
    public sealed class AxisAlignmentLink
    {
        internal AxisAlignmentLink(
            int? beforeIndex,
            int? afterIndex,
            AxisAlignmentKind kind,
            AlignmentConfidence confidence,
            AxisAlignmentStrategy strategy,
            IEnumerable<AlignmentReason> reasons,
            bool reordered = false)
        {
            if (kind == AxisAlignmentKind.Matched && (!beforeIndex.HasValue || !afterIndex.HasValue))
                throw new ArgumentException("A matched link requires both physical indices.");
            if (kind == AxisAlignmentKind.OnlyBefore && (!beforeIndex.HasValue || afterIndex.HasValue))
                throw new ArgumentException("A Before-only link requires only the Before physical index.");
            if (kind == AxisAlignmentKind.OnlyAfter && (beforeIndex.HasValue || !afterIndex.HasValue))
                throw new ArgumentException("An After-only link requires only the After physical index.");

            BeforeIndex = beforeIndex;
            AfterIndex = afterIndex;
            Kind = kind;
            Confidence = confidence;
            Strategy = strategy;
            Reasons = Freeze(reasons);
            Reordered = reordered;
        }

        public int? BeforeIndex { get; }
        public int? AfterIndex { get; }
        public AxisAlignmentKind Kind { get; }
        public AlignmentConfidence Confidence { get; }
        public AxisAlignmentStrategy Strategy { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }
        public bool Reordered { get; }

        private static IReadOnlyList<AlignmentReason> Freeze(IEnumerable<AlignmentReason> reasons) =>
            new ReadOnlyCollection<AlignmentReason>(new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
    }

    /// <summary>
    /// A many-to-one, one-to-many, or many-to-many candidate set that cannot be resolved safely. Its physical
    /// indices are deliberately absent from resolved links, preventing consumers from fabricating a pairing.
    /// </summary>
    public sealed class AxisAlignmentAmbiguity
    {
        internal AxisAlignmentAmbiguity(
            IEnumerable<int> beforeIndices,
            IEnumerable<int> afterIndices,
            string evidenceToken,
            IEnumerable<AlignmentReason> reasons)
        {
            BeforeIndices = Freeze(beforeIndices);
            AfterIndices = Freeze(afterIndices);
            EvidenceToken = evidenceToken ?? "";
            Reasons = new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
        }

        public IReadOnlyList<int> BeforeIndices { get; }
        public IReadOnlyList<int> AfterIndices { get; }
        public string EvidenceToken { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }

        private static IReadOnlyList<int> Freeze(IEnumerable<int> values)
        {
            var copy = new List<int>(values ?? Array.Empty<int>());
            copy.Sort();
            return new ReadOnlyCollection<int>(copy);
        }
    }

    public sealed class AxisAlignmentResult
    {
        internal AxisAlignmentResult(
            AxisAlignmentStrategy strategy,
            IEnumerable<AxisAlignmentLink> links,
            IEnumerable<AxisAlignmentAmbiguity> ambiguities,
            IEnumerable<AlignmentReason> reasons)
        {
            Strategy = strategy;
            Links = new ReadOnlyCollection<AxisAlignmentLink>(
                new List<AxisAlignmentLink>(links ?? Array.Empty<AxisAlignmentLink>()));
            Ambiguities = new ReadOnlyCollection<AxisAlignmentAmbiguity>(
                new List<AxisAlignmentAmbiguity>(ambiguities ?? Array.Empty<AxisAlignmentAmbiguity>()));
            Reasons = new ReadOnlyCollection<AlignmentReason>(
                new List<AlignmentReason>(reasons ?? Array.Empty<AlignmentReason>()));
        }

        public AxisAlignmentStrategy Strategy { get; }
        public IReadOnlyList<AxisAlignmentLink> Links { get; }
        public IReadOnlyList<AxisAlignmentAmbiguity> Ambiguities { get; }
        public IReadOnlyList<AlignmentReason> Reasons { get; }
        public bool HasAmbiguity => Ambiguities.Count > 0;

        public bool TryMapBefore(int beforeIndex, out int afterIndex)
        {
            foreach (AxisAlignmentLink link in Links)
            {
                if (link.Kind != AxisAlignmentKind.Matched || link.BeforeIndex != beforeIndex) continue;
                afterIndex = link.AfterIndex.Value;
                return true;
            }
            afterIndex = -1;
            return false;
        }
    }

    internal static class AlignmentText
    {
        internal static bool IsBlank(string value) => string.IsNullOrWhiteSpace(value);

        internal static string NormalizeLabel(string value, string commentPrefix = null)
        {
            if (value == null) return "";
            string trimmed = value.Trim();
            if (!string.IsNullOrEmpty(commentPrefix) &&
                trimmed.StartsWith(commentPrefix, StringComparison.Ordinal))
                trimmed = trimmed.Substring(commentPrefix.Length).TrimStart();

            var normalized = new StringBuilder(trimmed.Length);
            bool separator = false;
            foreach (char character in trimmed.Normalize(NormalizationForm.FormKC))
            {
                if (char.IsLetterOrDigit(character))
                {
                    if (separator && normalized.Length > 0) normalized.Append(' ');
                    normalized.Append(char.ToLowerInvariant(character));
                    separator = false;
                }
                else separator = true;
            }
            return normalized.ToString();
        }

        internal static string ExactTuple(AlignmentTable table, int rowIndex, IReadOnlyList<int> columns)
        {
            var text = new StringBuilder();
            foreach (int column in columns)
            {
                bool present = table.TryGetCell(rowIndex, column, out string value);
                if (!present)
                {
                    text.Append('M').Append(';');
                    continue;
                }
                if (value == null)
                {
                    text.Append('N').Append(';');
                    continue;
                }
                text.Append('V')
                    .Append(value.Length.ToString(CultureInfo.InvariantCulture))
                    .Append(':')
                    .Append(value)
                    .Append(';');
            }
            return text.ToString();
        }

        internal static string ExactRowFingerprint(AlignmentTable table, int rowIndex)
        {
            var columns = new int[table.Rows[rowIndex].Count];
            for (int i = 0; i < columns.Length; i++) columns[i] = i;
            return ExactTuple(table, rowIndex, columns);
        }

        internal static string ExactColumnFingerprint(AlignmentTable table, int columnIndex, int dataStartRow)
        {
            var text = new StringBuilder();
            for (int row = dataStartRow; row < table.RowCount; row++)
            {
                bool present = table.TryGetCell(row, columnIndex, out string value);
                if (!present) text.Append('M').Append(';');
                else if (value == null) text.Append('N').Append(';');
                else text.Append('V').Append(value.Length.ToString(CultureInfo.InvariantCulture))
                    .Append(':').Append(value).Append(';');
            }
            return text.ToString();
        }
    }
}
