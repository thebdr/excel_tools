using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    [Flags]
    public enum ReviewCellChange { None = 0, Content = 1, SavedResult = 2, Interior = 4, Strikethrough = 8, Presence = 16, IncompleteEvidence = 32 }
    public sealed class ReviewCellTransition
    {
        internal ReviewCellTransition(int beforeIndex, int afterIndex, ReviewCellChange changes,
            IEnumerable<ReviewTextSpan> text, IEnumerable<string> reasons)
        { BeforeIndex = beforeIndex; AfterIndex = afterIndex; Changes = changes; Text = Array.AsReadOnly(text.ToArray()); Reasons = Array.AsReadOnly(reasons.ToArray()); }
        public int BeforeIndex { get; }
        public int AfterIndex { get; }
        public ReviewCellChange Changes { get; }
        public ReadOnlyCollection<ReviewTextSpan> Text { get; }
        public ReadOnlyCollection<string> Reasons { get; }
    }
    public sealed class ReviewCellHistory
    {
        internal ReviewCellHistory(bool relevant, bool incomplete, IEnumerable<CellObservation> observations,
            IEnumerable<ReviewCellTransition> transitions, bool intermediateOnly)
        {
            IsRelevant = relevant; HasIncompleteEvidence = incomplete; Observations = Array.AsReadOnly(observations.ToArray());
            Transitions = Array.AsReadOnly(transitions.ToArray()); HasIntermediateOnlyChanges = intermediateOnly;
        }
        public bool IsRelevant { get; }
        public bool HasIncompleteEvidence { get; }
        public bool HasIntermediateOnlyChanges { get; }
        public ReadOnlyCollection<CellObservation> Observations { get; }
        public ReadOnlyCollection<ReviewCellTransition> Transitions { get; }
    }

    public static class CellHistoryComparison
    {
        public static ReviewCellHistory Compare(IEnumerable<CellObservation> history)
        {
            var cells = (history ?? throw new ArgumentNullException(nameof(history))).ToArray();
            if (cells.Length < 2 || cells.Any(c => c == null)) throw new ArgumentException("At least two observations are required.", nameof(history));
            bool relevant = cells.Any(c => c.HasContent);
            bool incomplete = cells.Any(c => !c.IsObserved && !c.IsKnownAbsent);
            var transitions = new List<ReviewCellTransition>();
            // Unknown content is not proof of an always-empty cell. Preserve a diagnostic, without inventing a data change.
            if (relevant || incomplete)
                for (int i = 1; i < cells.Length; i++)
                {
                    var transition = Pair(cells[i - 1], cells[i], i - 1, i);
                    if (transition.Changes != ReviewCellChange.None) transitions.Add(transition);
                }
            incomplete |= transitions.Any(t => (t.Changes & ReviewCellChange.IncompleteEvidence) != 0);
            var endpoints = Pair(cells[0], cells[cells.Length - 1], 0, cells.Length - 1);
            bool middle = relevant && transitions.Any(t => (t.Changes & ~ReviewCellChange.IncompleteEvidence) != 0) &&
                (endpoints.Changes & ~ReviewCellChange.IncompleteEvidence) == ReviewCellChange.None &&
                (endpoints.Changes & ReviewCellChange.IncompleteEvidence) == 0;
            return new ReviewCellHistory(relevant, incomplete, cells, transitions, middle);
        }

        private static ReviewCellTransition Pair(CellObservation a, CellObservation b, int ai, int bi)
        {
            var change = ReviewCellChange.None; var reasons = new List<string>();
            var text = new List<ReviewTextSpan>();
            if ((!a.IsObserved && !a.IsKnownAbsent) || (!b.IsObserved && !b.IsKnownAbsent))
            {
                change |= ReviewCellChange.IncompleteEvidence;
                if (!a.IsObserved) reasons.Add("Before: " + a.State + ". " + a.Reason);
                if (!b.IsObserved) reasons.Add("After: " + b.State + ". " + b.Reason);
            }
            else if (!a.IsObserved || !b.IsObserved)
            {
                if (a.State != b.State) change |= ReviewCellChange.Presence;
                if (a.HasContent || b.HasContent)
                {
                    change |= ReviewCellChange.Content;
                    text.AddRange(ExactTextDifference.Compare(a.IsObserved ? a.Content.ToString() : "", b.IsObserved ? b.Content.ToString() : ""));
                }
            }
            else
            {
                bool formulaMissing = (a.Content.IsFormula && !a.Content.FormulaEvidence.IsAvailable) ||
                    (b.Content.IsFormula && !b.Content.FormulaEvidence.IsAvailable);
                if (formulaMissing)
                {
                    change |= ReviewCellChange.IncompleteEvidence;
                    if (a.Content.IsFormula && !a.Content.FormulaEvidence.IsAvailable) reasons.Add("Formula before: " + a.Content.FormulaEvidence.Reason);
                    if (b.Content.IsFormula && !b.Content.FormulaEvidence.IsAvailable) reasons.Add("Formula after: " + b.Content.FormulaEvidence.Reason);
                    // A formula/constant type transition is known even when the formula text cannot yet be expanded.
                    if (a.Content.IsFormula != b.Content.IsFormula) change |= ReviewCellChange.Content;
                }
                else if (!a.Content.Equals(b.Content))
                { change |= ReviewCellChange.Content; text.AddRange(ExactTextDifference.Compare(a.Content.ToString(), b.Content.ToString())); }
                if (a.Content.IsFormula && b.Content.IsFormula)
                    CompareChannel(a.Content.SavedResult, b.Content.SavedResult, ReviewCellChange.SavedResult, "Saved calculated value", (x, y) => x.Equals(y));
                CompareChannel(a.Interior, b.Interior, ReviewCellChange.Interior, "Interior", (x, y) => x.Equals(y));
                CompareChannel(a.Strike, b.Strike, ReviewCellChange.Strikethrough, "Strikethrough", (x, y) => !ExactTextDifference.StrikeChanged(x, y));
            }
            return new ReviewCellTransition(ai, bi, change, text, reasons);

            void CompareChannel<T>(ChannelEvidence<T> x, ChannelEvidence<T> y, ReviewCellChange flag,
                string label, Func<T, T, bool> equal) where T : class
            {
                if (!x.IsAvailable || !y.IsAvailable)
                {
                    change |= ReviewCellChange.IncompleteEvidence;
                    if (!x.IsAvailable) reasons.Add(label + " before: " + x.Availability + ". " + x.Reason);
                    if (!y.IsAvailable) reasons.Add(label + " after: " + y.Availability + ". " + y.Reason);
                }
                else if (!equal(x.Value, y.Value)) change |= flag;
            }
        }
    }
}
