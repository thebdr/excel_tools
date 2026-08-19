using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelPowerTools.Core.Comparison.Canonical
{
    /// <summary>Immutable sparse formula/value snapshot for one revision occurrence.</summary>
    public sealed class RevisionSnapshot
    {
        private readonly Dictionary<string, CanonicalCellState> _cells;
        private readonly IReadOnlyDictionary<string, CanonicalCellState> _readOnlyCells;

        public RevisionSnapshot(string revisionId, IEnumerable<CanonicalCellState> cells)
        {
            if (string.IsNullOrEmpty(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));

            RevisionId = revisionId;
            _cells = new Dictionary<string, CanonicalCellState>(StringComparer.Ordinal);
            _readOnlyCells = new ReadOnlyDictionary<string, CanonicalCellState>(_cells);

            if (cells == null) return;
            foreach (CanonicalCellState cell in cells)
            {
                if (cell == null) throw new ArgumentException("Snapshots cannot contain null cells.", nameof(cells));
                if (_cells.ContainsKey(cell.SubjectKey))
                    throw new ArgumentException("Duplicate logical subject key: " + cell.SubjectKey, nameof(cells));
                _cells.Add(cell.SubjectKey, cell);
            }
        }

        public string RevisionId { get; }
        public IReadOnlyDictionary<string, CanonicalCellState> Cells => _readOnlyCells;

        public bool TryGetCell(string subjectKey, out CanonicalCellState cell) =>
            _cells.TryGetValue(subjectKey, out cell);
    }
}
