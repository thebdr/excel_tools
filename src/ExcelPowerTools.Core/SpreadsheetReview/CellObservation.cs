using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public enum CellObservationState { Observed, AbsentRow, AbsentColumn, Excluded, Unavailable, Failed, Unresolved }
    public enum FormattingOrigin { SavedManual, OrdinaryReaderDisplay, NotAvailable }

    public sealed class CellInterior : IEquatable<CellInterior>
    {
        public static readonly CellInterior NoFill = new CellInterior("none", null, null);

        public CellInterior(string pattern, string foregroundArgb, string backgroundArgb)
        {
            if (string.IsNullOrWhiteSpace(pattern)) throw new ArgumentException("A fill pattern is required.", nameof(pattern));
            Pattern = pattern.ToLowerInvariant();
            ForegroundArgb = Pattern == "none" ? null : Color(foregroundArgb);
            BackgroundArgb = Pattern == "none" ? null : Color(backgroundArgb);
        }

        public string Pattern { get; }
        public string ForegroundArgb { get; }
        public string BackgroundArgb { get; }

        private static string Color(string value)
        {
            if (value == null) return null;
            if (value.Length != 8 || value.Any(c => !Uri.IsHexDigit(c)))
                throw new ArgumentException("A resolved color must contain eight ARGB hexadecimal digits.", nameof(value));
            return value.ToUpperInvariant();
        }

        public bool Equals(CellInterior other) => other != null && Pattern == other.Pattern &&
            ForegroundArgb == other.ForegroundArgb && BackgroundArgb == other.BackgroundArgb;
        public override bool Equals(object obj) => Equals(obj as CellInterior);
        public override int GetHashCode() => (Pattern + ":" + ForegroundArgb + ":" + BackgroundArgb).GetHashCode();
        public override string ToString() => Pattern + ":" + ForegroundArgb + ":" + BackgroundArgb;
    }

    public sealed class StrikeRun
    {
        public StrikeRun(int start, int length, bool struck)
        {
            if (start < 0) throw new ArgumentOutOfRangeException(nameof(start));
            if (length <= 0) throw new ArgumentOutOfRangeException(nameof(length));
            Start = start; Length = length; Struck = struck;
        }
        public int Start { get; }
        public int Length { get; }
        public bool Struck { get; }
    }

    /// <summary>Effective UTF-16 character formatting; storage run boundaries do not define equality.</summary>
    public sealed class StrikeProfile : IEquatable<StrikeProfile>
    {
        private readonly bool[] characters;

        public StrikeProfile(string text, bool cellStrike, IEnumerable<StrikeRun> overrides = null)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            CellStrike = cellStrike;
            characters = Enumerable.Repeat(cellStrike, text.Length).ToArray();
            int lastEnd = 0;
            foreach (var run in (overrides ?? Enumerable.Empty<StrikeRun>()).OrderBy(r => r?.Start ?? -1))
            {
                if (run == null || run.Start < lastEnd || run.Start > text.Length || run.Length > text.Length - run.Start)
                    throw new ArgumentException("Strike runs must be non-overlapping and inside the exact text.", nameof(overrides));
                for (int i = run.Start; i < run.Start + run.Length; i++) characters[i] = run.Struck;
                lastEnd = run.Start + run.Length;
            }
            CharacterStrike = Array.AsReadOnly(characters);
        }

        public string Text { get; }
        public bool CellStrike { get; }
        public ReadOnlyCollection<bool> CharacterStrike { get; }
        public bool Equals(StrikeProfile other) => other != null && Text == other.Text &&
            (Text.Length != 0 || CellStrike == other.CellStrike) && characters.SequenceEqual(other.characters);
        public override bool Equals(object obj) => Equals(obj as StrikeProfile);
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = Text.GetHashCode() * 397 ^ (Text.Length == 0 && CellStrike).GetHashCode();
                foreach (bool struck in characters) hash = hash * 31 + (struck ? 1 : 0);
                return hash;
            }
        }
    }

    public sealed class CellObservation
    {
        private CellObservation(CellObservationState state, CellContent content,
            ChannelEvidence<CellInterior> interior, ChannelEvidence<StrikeProfile> strike,
            FormattingOrigin origin, string reason)
        { State = state; Content = content; Interior = interior; Strike = strike; FormattingOrigin = origin; Reason = reason; }

        public CellObservationState State { get; }
        public CellContent Content { get; }
        public ChannelEvidence<CellInterior> Interior { get; }
        public ChannelEvidence<StrikeProfile> Strike { get; }
        public FormattingOrigin FormattingOrigin { get; }
        public string Reason { get; }
        public bool IsObserved => State == CellObservationState.Observed;
        public bool HasContent => IsObserved && Content.HasContent;
        public bool IsKnownAbsent => State == CellObservationState.AbsentRow || State == CellObservationState.AbsentColumn;

        public static CellObservation Observed(CellContent content, ChannelEvidence<CellInterior> interior = null,
            ChannelEvidence<StrikeProfile> strike = null, FormattingOrigin origin = FormattingOrigin.SavedManual)
        {
            if (content == null) throw new ArgumentNullException(nameof(content));
            if (strike != null && strike.IsAvailable && strike.Value.Text != content.ToString())
                throw new ArgumentException("Strike evidence must describe the exact observed content text.", nameof(strike));
            return new CellObservation(CellObservationState.Observed, content,
                interior ?? ChannelEvidence<CellInterior>.Unavailable("No interior observation was supplied."),
                strike ?? ChannelEvidence<StrikeProfile>.Unavailable("No strikethrough observation was supplied."), origin, null);
        }

        public static CellObservation Blank(ChannelEvidence<CellInterior> interior = null,
            ChannelEvidence<StrikeProfile> strike = null) => Observed(CellContent.Value(ReviewScalar.Blank), interior, strike);

        public static CellObservation Missing(CellObservationState state, string reason)
        {
            if (state == CellObservationState.Observed) throw new ArgumentException("Use Observed for present cells.", nameof(state));
            if (string.IsNullOrWhiteSpace(reason)) throw new ArgumentException("Unobserved content needs an explanation.", nameof(reason));
            return new CellObservation(state, null, ChannelEvidence<CellInterior>.Unavailable(reason),
                ChannelEvidence<StrikeProfile>.Unavailable(reason), FormattingOrigin.NotAvailable, reason);
        }
    }
}
