using System;
using System.Globalization;
using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>
    /// Immutable evidence observed for one logical subject in one revision. The public constructor is also the
    /// persistence constructor: every identity and content invariant is checked before an instance can exist.
    /// </summary>
    public sealed class Observation
    {
        public Observation(
            string observationId,
            string subjectId,
            string revisionId,
            ObservationExactType exactType,
            string contentHash,
            string formula = null,
            CanonicalScalar value = null,
            CanonicalScalar cachedResult = null,
            CachedResultTrust cachedResultTrust = CachedResultTrust.NotPresent,
            string displayText = "",
            ObservationProvenance provenance = null)
        {
            ObservationId = Require(observationId, nameof(observationId));
            SubjectId = Require(subjectId, nameof(subjectId));
            RevisionId = Require(revisionId, nameof(revisionId));
            ContentHash = Require(contentHash, nameof(contentHash));
            if (!Enum.IsDefined(typeof(ObservationExactType), exactType))
                throw new ArgumentOutOfRangeException(nameof(exactType));
            if (!Enum.IsDefined(typeof(CachedResultTrust), cachedResultTrust))
                throw new ArgumentOutOfRangeException(nameof(cachedResultTrust));

            ExactType = exactType;
            if (exactType == ObservationExactType.Formula)
            {
                if (string.IsNullOrEmpty(formula))
                    throw new ArgumentException("A formula observation requires exact formula text.", nameof(formula));
                if (value != null)
                    throw new ArgumentException("A formula observation cannot also contain a constant value.", nameof(value));
            }
            else
            {
                if (!string.IsNullOrEmpty(formula))
                    throw new ArgumentException("Only a formula observation may contain formula text.", nameof(formula));

                value = value ?? DefaultScalar(exactType);
                if (value == null || value.Type != exactType)
                    throw new ArgumentException("The canonical value type must match the observation exact type.", nameof(value));
                if (cachedResult != null || cachedResultTrust != CachedResultTrust.NotPresent)
                    throw new ArgumentException("Only a formula observation may contain a cached result.", nameof(cachedResult));
            }

            if (cachedResult == null && cachedResultTrust != CachedResultTrust.NotPresent)
                throw new ArgumentException("Cached-result trust requires a cached result.", nameof(cachedResultTrust));
            if (cachedResult != null && cachedResultTrust == CachedResultTrust.NotPresent)
                throw new ArgumentException("A cached result requires an explicit trust classification.", nameof(cachedResultTrust));

            Formula = formula ?? "";
            Value = value;
            CachedResult = cachedResult;
            CachedResultTrust = cachedResultTrust;
            DisplayText = displayText ?? "";
            Provenance = provenance ?? ObservationProvenance.Empty;
        }

        public string ObservationId { get; }
        public string SubjectId { get; }
        public string RevisionId { get; }
        public ObservationExactType ExactType { get; }
        public string ContentHash { get; }
        public string Formula { get; }
        public CanonicalScalar Value { get; }
        public CanonicalScalar CachedResult { get; }
        public CachedResultTrust CachedResultTrust { get; }
        public string DisplayText { get; }
        public ObservationProvenance Provenance { get; }

        private static CanonicalScalar DefaultScalar(ObservationExactType type)
        {
            switch (type)
            {
                case ObservationExactType.Missing: return CanonicalScalar.Missing;
                case ObservationExactType.Blank: return CanonicalScalar.Blank;
                default: return null;
            }
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable identity component is required.", name);
            return value;
        }
    }

    /// <summary>
    /// Culture-independent scalar representation. CanonicalText is produced by the canonicalizer and is never
    /// presentation text (for example, numbers use invariant round-trip formatting).
    /// </summary>
    public sealed class CanonicalScalar
    {
        private static readonly CanonicalScalar MissingInstance =
            new CanonicalScalar(ObservationExactType.Missing, "", ExcelDateSystem.NotApplicable);
        private static readonly CanonicalScalar BlankInstance =
            new CanonicalScalar(ObservationExactType.Blank, "", ExcelDateSystem.NotApplicable);

        public CanonicalScalar(
            ObservationExactType type,
            string canonicalText,
            ExcelDateSystem dateSystem = ExcelDateSystem.NotApplicable)
        {
            if (!Enum.IsDefined(typeof(ObservationExactType), type))
                throw new ArgumentOutOfRangeException(nameof(type));
            if (!Enum.IsDefined(typeof(ExcelDateSystem), dateSystem))
                throw new ArgumentOutOfRangeException(nameof(dateSystem));
            if (type == ObservationExactType.Formula)
                throw new ArgumentException("Formula is not a scalar value type.", nameof(type));
            if (canonicalText == null) throw new ArgumentNullException(nameof(canonicalText));

            if (type == ObservationExactType.DateSerial)
            {
                if (dateSystem != ExcelDateSystem.Date1900 && dateSystem != ExcelDateSystem.Date1904)
                    throw new ArgumentOutOfRangeException(nameof(dateSystem), "A date serial requires the 1900 or 1904 date system.");
                EnsureFiniteNumber(canonicalText, nameof(canonicalText));
            }
            else
            {
                if (dateSystem != ExcelDateSystem.NotApplicable)
                    throw new ArgumentOutOfRangeException(nameof(dateSystem), "Only a date serial may specify an Excel date system.");

                switch (type)
                {
                    case ObservationExactType.Missing:
                    case ObservationExactType.Blank:
                        if (canonicalText.Length != 0)
                            throw new ArgumentException("Missing and blank scalars cannot contain text.", nameof(canonicalText));
                        break;
                    case ObservationExactType.Number:
                        EnsureFiniteNumber(canonicalText, nameof(canonicalText));
                        break;
                    case ObservationExactType.Boolean:
                        if (canonicalText != "true" && canonicalText != "false")
                            throw new ArgumentException("A canonical Boolean must be 'true' or 'false'.", nameof(canonicalText));
                        break;
                    case ObservationExactType.Error:
                    case ObservationExactType.Unsupported:
                    case ObservationExactType.IsoDateTime:
                        if (canonicalText.Length == 0)
                            throw new ArgumentException("This scalar type requires exact content.", nameof(canonicalText));
                        break;
                }
            }

            Type = type;
            CanonicalText = canonicalText;
            DateSystem = dateSystem;
        }

        public static CanonicalScalar Missing => MissingInstance;
        public static CanonicalScalar Blank => BlankInstance;

        public ObservationExactType Type { get; }
        public string CanonicalText { get; }
        public ExcelDateSystem DateSystem { get; }

        private static void EnsureFiniteNumber(string text, string parameterName)
        {
            double value;
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new ArgumentException("A canonical numeric value must use finite invariant text.", parameterName);
            }
        }
    }

    public sealed class ObservationProvenance
    {
        private static readonly ObservationProvenance EmptyInstance = new ObservationProvenance();

        public ObservationProvenance(
            string sourceLabel = "",
            string sheetName = "",
            string address = "",
            int physicalRow = 0,
            int physicalColumn = 0,
            string logicalSheetId = "")
        {
            if (physicalRow < 0) throw new ArgumentOutOfRangeException(nameof(physicalRow));
            if (physicalColumn < 0) throw new ArgumentOutOfRangeException(nameof(physicalColumn));
            SourceLabel = sourceLabel ?? "";
            SheetName = sheetName ?? "";
            Address = address ?? "";
            PhysicalRow = physicalRow;
            PhysicalColumn = physicalColumn;
            LogicalSheetId = logicalSheetId ?? "";
        }

        public static ObservationProvenance Empty => EmptyInstance;

        public string SourceLabel { get; }
        public string SheetName { get; }
        public string Address { get; }
        public int PhysicalRow { get; }
        public int PhysicalColumn { get; }
        /// <summary>
        /// Stable aligned sheet identity. Empty preserves legacy observations that recorded physical provenance
        /// only; new aligned review evidence should populate it whenever a physical occurrence exists.
        /// </summary>
        public string LogicalSheetId { get; }
    }
}
