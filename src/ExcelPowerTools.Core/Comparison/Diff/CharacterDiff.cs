using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;

namespace ExcelPowerTools.Core.Comparison.Diff
{
    public enum CharacterDiffKind
    {
        Equal,
        Removed,
        Added,
    }

    public sealed class CharacterDiffRun
    {
        internal CharacterDiffRun(CharacterDiffKind kind, string text)
        {
            Kind = kind;
            Text = text ?? "";
        }

        public CharacterDiffKind Kind { get; }
        public string Text { get; }
    }

    public sealed class CharacterDiffResult
    {
        internal CharacterDiffResult(
            string beforeText,
            string afterText,
            IReadOnlyList<CharacterDiffRun> runs,
            bool isSimplified)
        {
            BeforeText = beforeText;
            AfterText = afterText;
            Runs = new ReadOnlyCollection<CharacterDiffRun>(
                new List<CharacterDiffRun>(runs ?? Array.Empty<CharacterDiffRun>()));
            IsSimplified = isSimplified;
        }

        public string BeforeText { get; }
        public string AfterText { get; }
        public IReadOnlyList<CharacterDiffRun> Runs { get; }

        /// <summary>
        /// True when a very large changed middle was represented as exact removed/added blocks instead of an
        /// expensive minimal alignment. The source strings and every character remain present.
        /// </summary>
        public bool IsSimplified { get; }
    }

    /// <summary>Unicode text-element-aware exact character diff for values and formulas.</summary>
    public static class CharacterDiff
    {
        public const int DefaultMaxMatrixCells = 1_000_000;

        public static CharacterDiffResult Compare(
            string beforeText,
            string afterText,
            int maxMatrixCells = DefaultMaxMatrixCells)
        {
            if (maxMatrixCells < 0) throw new ArgumentOutOfRangeException(nameof(maxMatrixCells));
            beforeText = beforeText ?? "";
            afterText = afterText ?? "";

            string[] before = TextElements(beforeText);
            string[] after = TextElements(afterText);
            int prefix = CommonPrefix(before, after);
            int suffix = CommonSuffix(before, after, prefix);
            int beforeMiddleLength = before.Length - prefix - suffix;
            int afterMiddleLength = after.Length - prefix - suffix;

            var runs = new List<CharacterDiffRun>();
            AddRun(runs, CharacterDiffKind.Equal, Join(before, 0, prefix));

            long matrixCells = (long)(beforeMiddleLength + 1) * (afterMiddleLength + 1);
            bool simplified = matrixCells > maxMatrixCells;
            if (simplified)
            {
                AddRun(runs, CharacterDiffKind.Removed, Join(before, prefix, beforeMiddleLength));
                AddRun(runs, CharacterDiffKind.Added, Join(after, prefix, afterMiddleLength));
            }
            else
            {
                AddMinimalMiddle(
                    before,
                    prefix,
                    beforeMiddleLength,
                    after,
                    prefix,
                    afterMiddleLength,
                    runs);
            }

            AddRun(runs, CharacterDiffKind.Equal, Join(before, before.Length - suffix, suffix));
            return new CharacterDiffResult(beforeText, afterText, runs, simplified);
        }

        public static string MakeWhitespaceVisible(string text)
        {
            if (string.IsNullOrEmpty(text)) return text ?? "";
            var result = new StringBuilder(text.Length);
            foreach (char ch in text)
            {
                switch (ch)
                {
                    case ' ': result.Append('·'); break;
                    case '\t': result.Append('⇥'); break;
                    case '\r': result.Append('␍'); break;
                    case '\n': result.Append('␊'); break;
                    default: result.Append(ch); break;
                }
            }
            return result.ToString();
        }

        private static void AddMinimalMiddle(
            string[] before,
            int beforeStart,
            int beforeLength,
            string[] after,
            int afterStart,
            int afterLength,
            List<CharacterDiffRun> runs)
        {
            var lcs = new int[beforeLength + 1, afterLength + 1];
            for (int i = beforeLength - 1; i >= 0; i--)
            {
                for (int j = afterLength - 1; j >= 0; j--)
                {
                    if (string.Equals(before[beforeStart + i], after[afterStart + j], StringComparison.Ordinal))
                        lcs[i, j] = lcs[i + 1, j + 1] + 1;
                    else
                        lcs[i, j] = Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
                }
            }

            int bi = 0;
            int ai = 0;
            while (bi < beforeLength || ai < afterLength)
            {
                if (bi < beforeLength && ai < afterLength &&
                    string.Equals(before[beforeStart + bi], after[afterStart + ai], StringComparison.Ordinal))
                {
                    AddRun(runs, CharacterDiffKind.Equal, before[beforeStart + bi]);
                    bi++;
                    ai++;
                }
                else if (bi < beforeLength &&
                         (ai >= afterLength || lcs[bi + 1, ai] >= lcs[bi, ai + 1]))
                {
                    AddRun(runs, CharacterDiffKind.Removed, before[beforeStart + bi]);
                    bi++;
                }
                else
                {
                    AddRun(runs, CharacterDiffKind.Added, after[afterStart + ai]);
                    ai++;
                }
            }
        }

        private static string[] TextElements(string text)
        {
            var elements = new List<string>();
            TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(text);
            while (enumerator.MoveNext()) elements.Add(enumerator.GetTextElement());
            return elements.ToArray();
        }

        private static int CommonPrefix(string[] before, string[] after)
        {
            int length = Math.Min(before.Length, after.Length);
            int i = 0;
            while (i < length && string.Equals(before[i], after[i], StringComparison.Ordinal)) i++;
            return i;
        }

        private static int CommonSuffix(string[] before, string[] after, int prefix)
        {
            int max = Math.Min(before.Length, after.Length) - prefix;
            int count = 0;
            while (count < max &&
                   string.Equals(before[before.Length - 1 - count], after[after.Length - 1 - count], StringComparison.Ordinal))
                count++;
            return count;
        }

        private static string Join(string[] elements, int start, int length)
        {
            if (length <= 0) return "";
            var text = new StringBuilder();
            for (int i = 0; i < length; i++) text.Append(elements[start + i]);
            return text.ToString();
        }

        private static void AddRun(List<CharacterDiffRun> runs, CharacterDiffKind kind, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (runs.Count > 0 && runs[runs.Count - 1].Kind == kind)
            {
                CharacterDiffRun previous = runs[runs.Count - 1];
                runs[runs.Count - 1] = new CharacterDiffRun(kind, previous.Text + text);
                return;
            }
            runs.Add(new CharacterDiffRun(kind, text));
        }
    }
}
