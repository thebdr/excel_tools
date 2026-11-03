using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using DiffPlex;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public enum ReviewTextSpanKind { Equal, Removed, Added }
    public sealed class ReviewTextSpan
    {
        internal ReviewTextSpan(ReviewTextSpanKind kind, string text, int beforeStart, int afterStart)
        { Kind = kind; Text = text; BeforeStart = beforeStart; AfterStart = afterStart; }
        public ReviewTextSpanKind Kind { get; }
        public string Text { get; }
        public int BeforeStart { get; }
        public int AfterStart { get; }
    }

    public static class ExactTextDifference
    {
        private sealed class TextElementChunker : IChunker
        {
            public IReadOnlyList<string> Chunk(string text)
            {
                int[] starts = StringInfo.ParseCombiningCharacters(text);
                return starts.Select((s, i) => text.Substring(s, (i + 1 < starts.Length ? starts[i + 1] : text.Length) - s)).ToArray();
            }
        }
        public static ReadOnlyCollection<ReviewTextSpan> Compare(string before, string after)
        {
            if (before == null || after == null) throw new ArgumentNullException();
            var diff = Differ.Instance.CreateDiffs(before, after, false, false, new TextElementChunker());
            var spans = new List<ReviewTextSpan>();
            int a = 0, b = 0, ai = 0;
            foreach (var block in diff.DiffBlocks)
            {
                Emit(ReviewTextSpanKind.Equal, diff.PiecesOld, ai, block.DeleteStartA - ai);
                Emit(ReviewTextSpanKind.Removed, diff.PiecesOld, block.DeleteStartA, block.DeleteCountA);
                Emit(ReviewTextSpanKind.Added, diff.PiecesNew, block.InsertStartB, block.InsertCountB);
                ai = block.DeleteStartA + block.DeleteCountA;
            }
            Emit(ReviewTextSpanKind.Equal, diff.PiecesOld, ai, diff.PiecesOld.Count - ai);
            return spans.AsReadOnly();

            void Emit(ReviewTextSpanKind kind, IReadOnlyList<string> pieces, int start, int count)
            {
                if (count == 0) return;
                string text = string.Concat(pieces.Skip(start).Take(count));
                if (text.Length == 0) return;
                spans.Add(new ReviewTextSpan(kind, text, a, b));
                if (kind != ReviewTextSpanKind.Added) a += text.Length;
                if (kind != ReviewTextSpanKind.Removed) b += text.Length;
            }
        }

        internal static bool StrikeChanged(StrikeProfile before, StrikeProfile after)
        {
            if (before.Text.Length == 0 || after.Text.Length == 0) return before.CellStrike != after.CellStrike;
            bool aUniform = before.CharacterStrike.All(s => s == before.CharacterStrike[0]);
            bool bUniform = after.CharacterStrike.All(s => s == after.CharacterStrike[0]);
            if (aUniform && bUniform) return before.CharacterStrike[0] != after.CharacterStrike[0];
            var removed = new List<bool>(); var added = new List<bool>();
            foreach (var span in Compare(before.Text, after.Text))
            {
                if (span.Kind == ReviewTextSpanKind.Equal)
                {
                    if (ChangedRuns()) return true;
                    removed.Clear(); added.Clear();
                    for (int i = 0; i < span.Text.Length; i++)
                        if (before.CharacterStrike[span.BeforeStart + i] != after.CharacterStrike[span.AfterStart + i]) return true;
                }
                else if (span.Kind == ReviewTextSpanKind.Removed)
                    removed.AddRange(before.CharacterStrike.Skip(span.BeforeStart).Take(span.Text.Length));
                else added.AddRange(after.CharacterStrike.Skip(span.AfterStart).Take(span.Text.Length));
            }
            return ChangedRuns();
            bool ChangedRuns() => (removed.Any(s => s) || added.Any(s => s)) && !removed.SequenceEqual(added);
        }
    }
}
