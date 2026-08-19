using System;
using System.Globalization;

namespace ExcelPowerTools.Core.Comparison.Canonical
{
    public enum CanonicalValueKind
    {
        Blank,
        Text,
        Number,
        Boolean,
        Error,
        DateSerial,
        IsoDateTime,
    }

    public enum ExcelDateSystem
    {
        NotApplicable,
        Date1900,
        Date1904,
    }

    /// <summary>
    /// Immutable typed scalar used by the comparison engine. Exact equality includes the value kind and, for
    /// dates, the workbook date system; display formatting is deliberately not part of primary equality.
    /// </summary>
    public sealed class CanonicalValue : IEquatable<CanonicalValue>
    {
        private static readonly CanonicalValue BlankInstance =
            new CanonicalValue(CanonicalValueKind.Blank, "", ExcelDateSystem.NotApplicable);

        private CanonicalValue(CanonicalValueKind kind, string exactText, ExcelDateSystem dateSystem)
        {
            Kind = kind;
            ExactText = exactText ?? "";
            DateSystem = dateSystem;
        }

        public CanonicalValueKind Kind { get; }

        /// <summary>Deterministic invariant representation of the typed value.</summary>
        public string ExactText { get; }

        public ExcelDateSystem DateSystem { get; }

        public static CanonicalValue Blank => BlankInstance;

        public static CanonicalValue Text(string value) =>
            new CanonicalValue(CanonicalValueKind.Text, value ?? "", ExcelDateSystem.NotApplicable);

        public static CanonicalValue Number(double value)
        {
            EnsureFinite(value, nameof(value));
            return new CanonicalValue(
                CanonicalValueKind.Number,
                value.ToString("R", CultureInfo.InvariantCulture),
                ExcelDateSystem.NotApplicable);
        }

        public static CanonicalValue Boolean(bool value) =>
            new CanonicalValue(
                CanonicalValueKind.Boolean,
                value ? "true" : "false",
                ExcelDateSystem.NotApplicable);

        public static CanonicalValue Error(string errorText)
        {
            if (string.IsNullOrEmpty(errorText))
                throw new ArgumentException("An Excel error value must have exact text.", nameof(errorText));

            return new CanonicalValue(CanonicalValueKind.Error, errorText, ExcelDateSystem.NotApplicable);
        }

        public static CanonicalValue DateSerial(double serial, ExcelDateSystem dateSystem)
        {
            EnsureFinite(serial, nameof(serial));
            if (dateSystem != ExcelDateSystem.Date1900 && dateSystem != ExcelDateSystem.Date1904)
                throw new ArgumentOutOfRangeException(nameof(dateSystem), "A date serial requires the 1900 or 1904 date system.");

            return new CanonicalValue(
                CanonicalValueKind.DateSerial,
                serial.ToString("R", CultureInfo.InvariantCulture),
                dateSystem);
        }

        /// <summary>
        /// An ISO 8601 date/date-time stored directly in an OOXML cell with t="d". Unlike a numeric Excel
        /// serial, this value is independent of the workbook's 1900/1904 date system. The validated stored
        /// lexical form is retained so exact comparison does not discard precision or offset evidence.
        /// </summary>
        public static CanonicalValue IsoDateTime(string exactIsoText)
        {
            if (string.IsNullOrEmpty(exactIsoText))
                throw new ArgumentException("An ISO date/time value must have exact stored text.", nameof(exactIsoText));

            return new CanonicalValue(
                CanonicalValueKind.IsoDateTime,
                exactIsoText,
                ExcelDateSystem.NotApplicable);
        }

        public bool Equals(CanonicalValue other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Kind == other.Kind &&
                   DateSystem == other.DateSystem &&
                   string.Equals(ExactText, other.ExactText, StringComparison.Ordinal);
        }

        public override bool Equals(object obj) => Equals(obj as CanonicalValue);

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = (int)Kind;
                hash = (hash * 397) ^ (int)DateSystem;
                hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(ExactText);
                return hash;
            }
        }

        public override string ToString()
        {
            if (Kind == CanonicalValueKind.Blank) return "";
            if (Kind == CanonicalValueKind.DateSerial) return ExactText + "@" + DateSystem;
            return ExactText;
        }

        private static void EnsureFinite(double value, string parameterName)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(parameterName, "Excel numeric values must be finite.");
        }
    }
}
