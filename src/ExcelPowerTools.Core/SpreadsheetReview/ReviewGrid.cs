using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public sealed class ReviewCell
    {
        public ReviewCell(int row, int column, CellObservation observation)
        {
            if (row < 1 || row > 1048576) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > 16384) throw new ArgumentOutOfRangeException(nameof(column));
            Row = row; Column = column; Observation = observation ?? throw new ArgumentNullException(nameof(observation));
        }
        public int Row { get; }
        public int Column { get; }
        public CellObservation Observation { get; }
    }

    /// <summary>A sparse saved observation. Missing stored cells within a row are known blanks, not read failures.</summary>
    public sealed class ReviewSheet
    {
        private static readonly CellObservation UnknownFormattingBlank = CellObservation.Blank();
        private readonly Dictionary<int, Dictionary<int, CellObservation>> rows;
        private readonly Func<int, int, CellObservation> blankReader;

        public ReviewSheet(string artifactId, string name, IEnumerable<ReviewCell> cells,
            Func<int, int, CellObservation> blankReader = null)
        {
            if (string.IsNullOrWhiteSpace(artifactId)) throw new ArgumentException("An artifact identity is required.", nameof(artifactId));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A sheet name is required.", nameof(name));
            ArtifactId = artifactId; Name = name;
            this.blankReader = blankReader ?? ((r, c) => UnknownFormattingBlank);
            rows = new Dictionary<int, Dictionary<int, CellObservation>>();
            foreach (var cell in cells ?? throw new ArgumentNullException(nameof(cells)))
            {
                if (cell == null) throw new ArgumentException("A cell cannot be null.", nameof(cells));
                if (!rows.TryGetValue(cell.Row, out var row)) rows.Add(cell.Row, row = new Dictionary<int, CellObservation>());
                if (row.ContainsKey(cell.Column)) throw new ArgumentException("Duplicate saved cell coordinates.", nameof(cells));
                row.Add(cell.Column, cell.Observation);
            }
            Cells = Array.AsReadOnly(rows.OrderBy(r => r.Key).SelectMany(r => r.Value.OrderBy(c => c.Key)
                .Select(c => new ReviewCell(r.Key, c.Key, c.Value))).ToArray());
        }

        public string ArtifactId { get; }
        public string Name { get; }
        public ReadOnlyCollection<ReviewCell> Cells { get; }
        public CellObservation Cell(int row, int column)
        {
            if (row < 1 || row > 1048576 || column < 1 || column > 16384) throw new ArgumentOutOfRangeException();
            return rows.TryGetValue(row, out var cells) && cells.TryGetValue(column, out var value)
                ? value : blankReader(row, column);
        }
        public int LastContentRow(int columnCap) => Cells.Where(c => c.Column <= columnCap &&
            (c.Observation.HasContent || !c.Observation.IsObserved)).Select(c => c.Row).DefaultIfEmpty(0).Max();
        public int LastContentColumn => Cells.Where(c => c.Observation.HasContent || !c.Observation.IsObserved)
            .Select(c => c.Column).DefaultIfEmpty(0).Max();
    }

    public sealed class ReviewColumnPair
    {
        public ReviewColumnPair(int? beforeColumn, int? afterColumn)
        {
            if (!beforeColumn.HasValue && !afterColumn.HasValue) throw new ArgumentException("A column must exist on at least one side.");
            if (beforeColumn.HasValue && (beforeColumn < 1 || beforeColumn > 16384)) throw new ArgumentOutOfRangeException(nameof(beforeColumn));
            if (afterColumn.HasValue && (afterColumn < 1 || afterColumn > 16384)) throw new ArgumentOutOfRangeException(nameof(afterColumn));
            BeforeColumn = beforeColumn; AfterColumn = afterColumn;
        }
        public int? BeforeColumn { get; }
        public int? AfterColumn { get; }
    }

    /// <summary>Only an explicitly complete official-column mapping can enter the alignment engine.</summary>
    public sealed class ResolvedReviewScope
    {
        public ResolvedReviewScope(int beforeCap, int afterCap, IEnumerable<ReviewColumnPair> columns)
        {
            if (beforeCap < 1 || beforeCap > 16384 || afterCap < 1 || afterCap > 16384) throw new ArgumentOutOfRangeException();
            var copy = (columns ?? throw new ArgumentNullException(nameof(columns))).ToArray();
            if (copy.Any(c => c == null)) throw new ArgumentException("Column pairs cannot be null.", nameof(columns));
            Validate(copy.Where(c => c.BeforeColumn.HasValue).Select(c => c.BeforeColumn.Value), beforeCap);
            Validate(copy.Where(c => c.AfterColumn.HasValue).Select(c => c.AfterColumn.Value), afterCap);
            BeforeCap = beforeCap; AfterCap = afterCap; Columns = Array.AsReadOnly(copy);
        }
        private static void Validate(IEnumerable<int> values, int cap)
        {
            if (!values.OrderBy(c => c).SequenceEqual(Enumerable.Range(1, cap)))
                throw new ArgumentException("Every included physical column must occur exactly once; excluded columns cannot be matching evidence.");
        }
        public int BeforeCap { get; }
        public int AfterCap { get; }
        public ReadOnlyCollection<ReviewColumnPair> Columns { get; }
        public static ResolvedReviewScope SameColumns(int cap) => new ResolvedReviewScope(cap, cap,
            Enumerable.Range(1, cap).Select(c => new ReviewColumnPair(c, c)));

        internal CellObservation At(ReviewSheet sheet, int row, int slot, bool before)
        {
            int? column = before ? Columns[slot].BeforeColumn : Columns[slot].AfterColumn;
            return column.HasValue ? sheet.Cell(row, column.Value) :
                CellObservation.Missing(CellObservationState.AbsentColumn, "The official column does not exist in this revision.");
        }

        internal string AnchorToken(ReviewSheet sheet, int row, bool before)
        {
            var token = new StringBuilder(); int informative = 0;
            for (int i = 0; i < Columns.Count; i++)
            {
                // A new/deleted column can carry a difference, but cannot supply cross-revision identity evidence.
                if (!Columns[i].BeforeColumn.HasValue || !Columns[i].AfterColumn.HasValue) continue;
                var observation = At(sheet, row, i, before);
                if (!observation.IsObserved) return null;
                var content = observation.Content;
                // Formula expressions are a separate channel. Literal row references cannot become identity keys.
                string value = content.IsFormula ? "FORMULA" : content.IdentityToken;
                if (content.HasInformativeConstant) informative++;
                token.Append(value.Length).Append(':').Append(value).Append(';');
            }
            return informative == 0 ? null : token.ToString();
        }
    }
}
