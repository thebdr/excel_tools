using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    public enum WorkbookSheetKind
    {
        Unknown,
        Worksheet,
        ChartSheet,
        DialogSheet,
        MacroSheet,
    }

    public enum SheetMappingKind
    {
        ExactName,
        InferredRename,
        UserConfirmed,
    }

    /// <summary>
    /// Stable reason codes exposed to diagnostics and a future mapping UI. Evidence codes describe why a pair
    /// is plausible; unresolved codes describe why the engine deliberately declined to choose it.
    /// </summary>
    public static class SheetMappingReasonCodes
    {
        public const string ExactName = "SHEET-NAME-EXACT";
        public const string AddInIdentity = "SHEET-ADDIN-IDENTITY-MATCH";
        public const string CodeName = "SHEET-CODENAME-MATCH";
        public const string StructureFingerprint = "SHEET-STRUCTURE-FINGERPRINT-MATCH";
        public const string ContentFingerprint = "SHEET-CONTENT-FINGERPRINT-MATCH";
        public const string HeaderFingerprint = "SHEET-HEADER-FINGERPRINT-MATCH";
        public const string FormulaFingerprint = "SHEET-FORMULA-FINGERPRINT-MATCH";
        public const string Dimensions = "SHEET-DIMENSIONS-MATCH";
        public const string StoredCellCount = "SHEET-STORED-CELL-COUNT-SIMILAR";
        public const string UserConfirmed = "SHEET-PAIR-USER-CONFIRMED";
        public const string ConfirmedTargetClaimed = "SHEET-TARGET-CLAIMED-BY-USER-CONFIRMATION";

        public const string ExactNameNotUnique = "SHEET-NAME-MATCH-NOT-UNIQUE";
        public const string NoSupportedRenameEvidence = "SHEET-RENAME-EVIDENCE-INSUFFICIENT";
        public const string BelowRenameThreshold = "SHEET-RENAME-CONFIDENCE-BELOW-THRESHOLD";
        public const string CompetingTargets = "SHEET-RENAME-COMPETING-TARGETS";
        public const string CompetingSources = "SHEET-RENAME-COMPETING-SOURCES";
        public const string NotAutoResolved = "SHEET-RENAME-REQUIRES-CONFIRMATION";
    }

    /// <summary>
    /// Immutable, workbook-local evidence for one sheet. Fingerprints are opaque canonical tokens produced by
    /// adapters; this type does not read a workbook or infer meaning from mutable COM objects.
    /// </summary>
    public sealed class WorkbookSheetFingerprint
    {
        public WorkbookSheetFingerprint(
            string sheetKey,
            string name,
            WorkbookSheetKind kind = WorkbookSheetKind.Worksheet,
            string codeName = null,
            string addInIdentity = null,
            string structureFingerprint = null,
            string contentFingerprint = null,
            string headerFingerprint = null,
            string formulaFingerprint = null,
            int usedRowCount = 0,
            int usedColumnCount = 0,
            long storedCellCount = 0)
        {
            if (string.IsNullOrWhiteSpace(sheetKey))
                throw new ArgumentException("A workbook-local sheet key is required.", nameof(sheetKey));
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            if (name.Length > 0 && string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("A sheet name must be empty or non-blank.", nameof(name));
            if (!Enum.IsDefined(typeof(WorkbookSheetKind), kind))
                throw new ArgumentOutOfRangeException(nameof(kind));
            if (usedRowCount < 0) throw new ArgumentOutOfRangeException(nameof(usedRowCount));
            if (usedColumnCount < 0) throw new ArgumentOutOfRangeException(nameof(usedColumnCount));
            if (storedCellCount < 0) throw new ArgumentOutOfRangeException(nameof(storedCellCount));

            SheetKey = sheetKey;
            Name = name;
            Kind = kind;
            CodeName = NormalizeOptional(codeName, nameof(codeName));
            AddInIdentity = NormalizeOptional(addInIdentity, nameof(addInIdentity));
            StructureFingerprint = NormalizeOptional(structureFingerprint, nameof(structureFingerprint));
            ContentFingerprint = NormalizeOptional(contentFingerprint, nameof(contentFingerprint));
            HeaderFingerprint = NormalizeOptional(headerFingerprint, nameof(headerFingerprint));
            FormulaFingerprint = NormalizeOptional(formulaFingerprint, nameof(formulaFingerprint));
            UsedRowCount = usedRowCount;
            UsedColumnCount = usedColumnCount;
            StoredCellCount = storedCellCount;
        }

        public string SheetKey { get; }
        public string Name { get; }
        public WorkbookSheetKind Kind { get; }
        public string CodeName { get; }
        public string AddInIdentity { get; }
        public string StructureFingerprint { get; }
        public string ContentFingerprint { get; }
        public string HeaderFingerprint { get; }
        public string FormulaFingerprint { get; }
        public int UsedRowCount { get; }
        public int UsedColumnCount { get; }
        public long StoredCellCount { get; }

        private static string NormalizeOptional(string value, string parameterName)
        {
            if (value == null || value.Length == 0) return null;
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Optional sheet evidence must be null, empty, or non-blank.", parameterName);
            return value;
        }
    }

    public sealed class WorkbookSheetMappingOptions
    {
        public static readonly WorkbookSheetMappingOptions Default = new WorkbookSheetMappingOptions();

        public WorkbookSheetMappingOptions(
            double minimumRenameConfidence = 0.65,
            double minimumWinnerMargin = 0.15)
        {
            if (double.IsNaN(minimumRenameConfidence) || double.IsInfinity(minimumRenameConfidence) ||
                minimumRenameConfidence <= 0 || minimumRenameConfidence > 1)
                throw new ArgumentOutOfRangeException(nameof(minimumRenameConfidence));
            if (double.IsNaN(minimumWinnerMargin) || double.IsInfinity(minimumWinnerMargin) ||
                minimumWinnerMargin < 0 || minimumWinnerMargin >= 1)
                throw new ArgumentOutOfRangeException(nameof(minimumWinnerMargin));

            MinimumRenameConfidence = minimumRenameConfidence;
            MinimumWinnerMargin = minimumWinnerMargin;
        }

        public double MinimumRenameConfidence { get; }
        public double MinimumWinnerMargin { get; }
    }

    public sealed class SheetMappingCandidate
    {
        internal SheetMappingCandidate(
            WorkbookSheetFingerprint before,
            WorkbookSheetFingerprint after,
            double confidence,
            IReadOnlyList<string> reasonCodes)
        {
            Before = before ?? throw new ArgumentNullException(nameof(before));
            After = after ?? throw new ArgumentNullException(nameof(after));
            Confidence = confidence;
            ReasonCodes = Copy(reasonCodes);
        }

        public WorkbookSheetFingerprint Before { get; }
        public WorkbookSheetFingerprint After { get; }
        public double Confidence { get; }
        public IReadOnlyList<string> ReasonCodes { get; }

        private static IReadOnlyList<string> Copy(IReadOnlyList<string> values) =>
            new ReadOnlyCollection<string>(new List<string>(values ?? Array.Empty<string>()));
    }

    public sealed class ResolvedSheetMapping
    {
        internal ResolvedSheetMapping(
            WorkbookSheetFingerprint before,
            WorkbookSheetFingerprint after,
            SheetMappingKind kind,
            double confidence,
            IReadOnlyList<string> reasonCodes)
        {
            Before = before ?? throw new ArgumentNullException(nameof(before));
            After = after ?? throw new ArgumentNullException(nameof(after));
            Kind = kind;
            Confidence = confidence;
            ReasonCodes = new ReadOnlyCollection<string>(
                new List<string>(reasonCodes ?? Array.Empty<string>()));
        }

        public WorkbookSheetFingerprint Before { get; }
        public WorkbookSheetFingerprint After { get; }
        public SheetMappingKind Kind { get; }
        public double Confidence { get; }
        public IReadOnlyList<string> ReasonCodes { get; }
    }

    public sealed class UnresolvedSheetMapping
    {
        internal UnresolvedSheetMapping(
            WorkbookSheetFingerprint before,
            IReadOnlyList<SheetMappingCandidate> candidates,
            IReadOnlyList<string> reasonCodes)
        {
            Before = before ?? throw new ArgumentNullException(nameof(before));
            Candidates = new ReadOnlyCollection<SheetMappingCandidate>(
                new List<SheetMappingCandidate>(candidates ?? Array.Empty<SheetMappingCandidate>()));
            ReasonCodes = new ReadOnlyCollection<string>(
                new List<string>(reasonCodes ?? Array.Empty<string>()));
        }

        public WorkbookSheetFingerprint Before { get; }
        public IReadOnlyList<SheetMappingCandidate> Candidates { get; }
        public IReadOnlyList<string> ReasonCodes { get; }
    }

    public sealed class WorkbookSheetMappingResult
    {
        internal WorkbookSheetMappingResult(
            IReadOnlyList<ResolvedSheetMapping> mappings,
            IReadOnlyList<UnresolvedSheetMapping> unresolvedBeforeSheets,
            IReadOnlyList<WorkbookSheetFingerprint> unmatchedBeforeSheets,
            IReadOnlyList<WorkbookSheetFingerprint> unmatchedAfterSheets)
        {
            Mappings = new ReadOnlyCollection<ResolvedSheetMapping>(
                new List<ResolvedSheetMapping>(mappings ?? Array.Empty<ResolvedSheetMapping>()));
            UnresolvedBeforeSheets = new ReadOnlyCollection<UnresolvedSheetMapping>(
                new List<UnresolvedSheetMapping>(unresolvedBeforeSheets ?? Array.Empty<UnresolvedSheetMapping>()));
            UnmatchedBeforeSheets = new ReadOnlyCollection<WorkbookSheetFingerprint>(
                new List<WorkbookSheetFingerprint>(unmatchedBeforeSheets ?? Array.Empty<WorkbookSheetFingerprint>()));
            UnmatchedAfterSheets = new ReadOnlyCollection<WorkbookSheetFingerprint>(
                new List<WorkbookSheetFingerprint>(unmatchedAfterSheets ?? Array.Empty<WorkbookSheetFingerprint>()));
        }

        public IReadOnlyList<ResolvedSheetMapping> Mappings { get; }
        public IReadOnlyList<UnresolvedSheetMapping> UnresolvedBeforeSheets { get; }
        /// <summary>Exact source occurrences explicitly confirmed as deleted on this edge.</summary>
        public IReadOnlyList<WorkbookSheetFingerprint> UnmatchedBeforeSheets { get; }
        public IReadOnlyList<WorkbookSheetFingerprint> UnmatchedAfterSheets { get; }
        public bool RequiresUserMapping => UnresolvedBeforeSheets.Count > 0;
    }
}
