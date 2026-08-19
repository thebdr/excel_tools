using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;

namespace ExcelPowerTools.Core.Comparison.Alignment
{
    public static class OpenXmlSheetFingerprintCoverageCodes
    {
        public const string InvalidCanonicalCellReference =
            "CMP-OX-SHEET-FINGERPRINT-CELL-REFERENCE-INVALID";
    }

    /// <summary>
    /// One generic fingerprint together with the immutable OOXML evidence and coverage caveats from which it was
    /// derived. Coverage codes are retained even when they do not block formula/value evidence.
    /// </summary>
    public sealed class OpenXmlSheetFingerprintEvidence
    {
        internal OpenXmlSheetFingerprintEvidence(
            OpenXmlWorksheetSnapshot source,
            WorkbookSheetFingerprint fingerprint,
            bool canUseCanonicalCellEvidence,
            IReadOnlyList<string> coverageCodes)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Fingerprint = fingerprint ?? throw new ArgumentNullException(nameof(fingerprint));
            CanUseCanonicalCellEvidence = canUseCanonicalCellEvidence;
            CoverageCodes = new ReadOnlyCollection<string>(
                new List<string>(coverageCodes ?? Array.Empty<string>()));
        }

        public OpenXmlWorksheetSnapshot Source { get; }
        public WorkbookSheetFingerprint Fingerprint { get; }
        public bool CanUseCanonicalCellEvidence { get; }
        public IReadOnlyList<string> CoverageCodes { get; }
    }

    public sealed class OpenXmlWorkbookSheetFingerprintSet
    {
        internal OpenXmlWorkbookSheetFingerprintSet(
            OpenXmlWorkbookSnapshot source,
            IReadOnlyList<OpenXmlSheetFingerprintEvidence> sheets,
            IReadOnlyList<string> workbookCoverageCodes)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Sheets = new ReadOnlyCollection<OpenXmlSheetFingerprintEvidence>(
                new List<OpenXmlSheetFingerprintEvidence>(sheets ?? Array.Empty<OpenXmlSheetFingerprintEvidence>()));
            Fingerprints = new ReadOnlyCollection<WorkbookSheetFingerprint>(
                Sheets.Select(item => item.Fingerprint).ToList());
            WorkbookCoverageCodes = new ReadOnlyCollection<string>(
                new List<string>(workbookCoverageCodes ?? Array.Empty<string>()));
        }

        public OpenXmlWorkbookSnapshot Source { get; }
        public string RevisionId => Source.RevisionId;
        public IReadOnlyList<OpenXmlSheetFingerprintEvidence> Sheets { get; }
        public IReadOnlyList<WorkbookSheetFingerprint> Fingerprints { get; }
        public IReadOnlyList<string> WorkbookCoverageCodes { get; }
    }

    public sealed class OpenXmlWorkbookSheetMappingResult
    {
        internal OpenXmlWorkbookSheetMappingResult(
            OpenXmlWorkbookSheetFingerprintSet before,
            OpenXmlWorkbookSheetFingerprintSet after,
            WorkbookSheetMappingResult mapping)
        {
            Before = before ?? throw new ArgumentNullException(nameof(before));
            After = after ?? throw new ArgumentNullException(nameof(after));
            Mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));
        }

        public OpenXmlWorkbookSheetFingerprintSet Before { get; }
        public OpenXmlWorkbookSheetFingerprintSet After { get; }
        public WorkbookSheetMappingResult Mapping { get; }
    }
}
