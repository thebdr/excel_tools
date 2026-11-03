using System;
using System.Globalization;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public enum ReviewScalarKind { Blank, Text, Number, Boolean, Error }

    /// <summary>Typed saved evidence, not a display string and never an executable expression.</summary>
    public sealed class ReviewScalar : IEquatable<ReviewScalar>
    {
        public static readonly ReviewScalar Blank = new ReviewScalar(ReviewScalarKind.Blank, "", 0, false);

        private ReviewScalar(ReviewScalarKind kind, string encoding, double number, bool boolean)
        {
            Kind = kind;
            OriginalEncoding = encoding;
            NumberValue = number;
            BooleanValue = boolean;
        }

        public ReviewScalarKind Kind { get; }
        public string OriginalEncoding { get; }
        public double NumberValue { get; }
        public bool BooleanValue { get; }
        public bool HasContent => Kind != ReviewScalarKind.Blank &&
            (Kind != ReviewScalarKind.Text || OriginalEncoding.Length != 0);
        public bool IsInformative => Kind == ReviewScalarKind.Text
            ? !string.IsNullOrWhiteSpace(OriginalEncoding)
            : Kind == ReviewScalarKind.Number && NumberValue != 0;

        public static ReviewScalar Text(string value) => new ReviewScalar(ReviewScalarKind.Text,
            value ?? throw new ArgumentNullException(nameof(value)), 0, false);

        public static ReviewScalar Number(string savedEncoding)
        {
            if (savedEncoding == null) throw new ArgumentNullException(nameof(savedEncoding));
            if (!double.TryParse(savedEncoding, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ||
                double.IsNaN(value) || double.IsInfinity(value))
                throw new FormatException("A saved spreadsheet number must be a finite invariant-culture number: " + savedEncoding);
            return new ReviewScalar(ReviewScalarKind.Number, savedEncoding, value, false);
        }

        public static ReviewScalar Boolean(bool value) =>
            new ReviewScalar(ReviewScalarKind.Boolean, value ? "TRUE" : "FALSE", 0, value);

        public static ReviewScalar Error(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An error value is required.", nameof(value));
            return new ReviewScalar(ReviewScalarKind.Error, value, 0, false);
        }

        // Numeric identity uses the parsed binary value, never its formatting/serialization spelling.
        internal string IdentityToken => ((int)Kind).ToString(CultureInfo.InvariantCulture) + ":" +
            (Kind == ReviewScalarKind.Number
                ? BitConverter.DoubleToInt64Bits(NumberValue == 0 ? 0 : NumberValue).ToString("X16", CultureInfo.InvariantCulture)
                : OriginalEncoding);

        public bool Equals(ReviewScalar other)
        {
            if (other == null || other.Kind != Kind) return false;
            if (Kind == ReviewScalarKind.Number) return NumberValue.Equals(other.NumberValue);
            return string.Equals(OriginalEncoding, other.OriginalEncoding, StringComparison.Ordinal);
        }

        public override bool Equals(object obj) => Equals(obj as ReviewScalar);
        public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(IdentityToken);
        public override string ToString() => OriginalEncoding;
    }

    public enum EvidenceAvailability { Available, Unavailable, Failed }

    public sealed class ChannelEvidence<T> where T : class
    {
        private ChannelEvidence(EvidenceAvailability availability, T value, string reason)
        { Availability = availability; Value = value; Reason = reason; }

        public EvidenceAvailability Availability { get; }
        public T Value { get; }
        public string Reason { get; }
        public bool IsAvailable => Availability == EvidenceAvailability.Available;

        public static ChannelEvidence<T> Observed(T value) => new ChannelEvidence<T>(
            EvidenceAvailability.Available, value ?? throw new ArgumentNullException(nameof(value)), null);

        public static ChannelEvidence<T> Unavailable(string reason) => Missing(EvidenceAvailability.Unavailable, reason);
        public static ChannelEvidence<T> Failed(string reason) => Missing(EvidenceAvailability.Failed, reason);

        private static ChannelEvidence<T> Missing(EvidenceAvailability availability, string reason)
        {
            if (string.IsNullOrWhiteSpace(reason)) throw new ArgumentException("Missing evidence needs an explanation.", nameof(reason));
            return new ChannelEvidence<T>(availability, null, reason);
        }
    }

    public sealed class CellContent : IEquatable<CellContent>
    {
        private CellContent(ReviewScalar constant, ChannelEvidence<string> formula, ChannelEvidence<ReviewScalar> result,
            string savedFormulaDescriptor = null)
        { Constant = constant; FormulaEvidence = formula; SavedResult = result; SavedFormulaDescriptor = savedFormulaDescriptor; }

        public ReviewScalar Constant { get; }
        public ChannelEvidence<string> FormulaEvidence { get; }
        public string Formula => FormulaEvidence?.Value;
        public string SavedFormulaDescriptor { get; }
        public ChannelEvidence<ReviewScalar> SavedResult { get; }
        public bool IsFormula => FormulaEvidence != null;
        public bool HasContent => IsFormula || Constant.HasContent;
        public bool HasInformativeConstant => !IsFormula && Constant.IsInformative;

        public static CellContent Value(ReviewScalar value) => new CellContent(
            value ?? throw new ArgumentNullException(nameof(value)), null,
            ChannelEvidence<ReviewScalar>.Unavailable("A constant has no saved formula result."));

        public static CellContent Expression(string formula, ChannelEvidence<ReviewScalar> savedResult = null)
        {
            if (string.IsNullOrWhiteSpace(formula) || formula == "=")
                throw new ArgumentException("A resolved formula expression is required.", nameof(formula));
            return new CellContent(null, ChannelEvidence<string>.Observed(formula[0] == '=' ? formula : "=" + formula),
                savedResult ?? ChannelEvidence<ReviewScalar>.Unavailable("No saved formula result was supplied."));
        }

        public static CellContent UnexpandedExpression(string savedDescriptor, string reason, ChannelEvidence<ReviewScalar> savedResult)
        {
            if (string.IsNullOrWhiteSpace(savedDescriptor)) throw new ArgumentException("Retain the saved formula descriptor.", nameof(savedDescriptor));
            return new CellContent(null, ChannelEvidence<string>.Unavailable(reason),
                savedResult ?? throw new ArgumentNullException(nameof(savedResult)), savedDescriptor);
        }

        internal string IdentityToken => IsFormula ? "F:" + (Formula ?? "?" + SavedFormulaDescriptor) : "V:" + Constant.IdentityToken;
        public bool Equals(CellContent other) => other != null && IsFormula == other.IsFormula &&
            (IsFormula ? FormulaEvidence.Availability == other.FormulaEvidence.Availability &&
                string.Equals(Formula, other.Formula, StringComparison.Ordinal) && SavedFormulaDescriptor == other.SavedFormulaDescriptor : Constant.Equals(other.Constant));
        public override bool Equals(object obj) => Equals(obj as CellContent);
        public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(IdentityToken);
        public override string ToString() => IsFormula ? Formula ?? "[Unexpanded formula: " + SavedFormulaDescriptor + "]" : Constant.ToString();
    }
}
