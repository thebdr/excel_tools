using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using ExcelPowerTools.Core.Comparison.Canonical;

namespace ExcelPowerTools.Core.Comparison.Adapters.OpenXml
{
    /// <summary>Stable diagnostic identifiers emitted by the read-only OOXML adapter.</summary>
    public static class OpenXmlCoverageCodes
    {
        public const string VbaProjectNotInspected = "CMP-OX-VBA-NOT-INSPECTED";
        public const string ExternalRelationshipsIgnored = "CMP-OX-EXTERNAL-RELATIONSHIP-IGNORED";
        public const string ExternalSheetRelationshipIgnored = "CMP-OX-EXTERNAL-SHEET-IGNORED";
        public const string MissingSheetsCollection = "CMP-OX-SHEETS-MISSING";
        public const string MissingSheetName = "CMP-OX-SHEET-NAME-MISSING";
        public const string MissingSheetId = "CMP-OX-SHEET-ID-MISSING";
        public const string MissingSheetRelationship = "CMP-OX-SHEET-RELATIONSHIP-MISSING";
        public const string UnresolvedSheetRelationship = "CMP-OX-SHEET-RELATIONSHIP-UNRESOLVED";
        public const string UnknownSheetVisibility = "CMP-OX-SHEET-VISIBILITY-UNKNOWN";
        public const string UnsupportedPhysicalSheet = "CMP-OX-SHEET-KIND-UNSUPPORTED";
        public const string MissingSheetData = "CMP-OX-SHEET-DATA-MISSING";
        public const string MultipleSheetData = "CMP-OX-SHEET-DATA-MULTIPLE";
        public const string InvalidCellReference = "CMP-OX-CELL-REFERENCE-INVALID";
        public const string DuplicateCellReference = "CMP-OX-CELL-REFERENCE-DUPLICATE";
        public const string MalformedCellPayload = "CMP-OX-CELL-PAYLOAD-MALFORMED";
        public const string InvalidSharedStringIndex = "CMP-OX-SHARED-STRING-INDEX-INVALID";
        public const string InvalidCellStyleIndex = "CMP-OX-CELL-STYLE-INDEX-INVALID";
        public const string InvalidBooleanValue = "CMP-OX-BOOLEAN-INVALID";
        public const string InvalidErrorValue = "CMP-OX-ERROR-INVALID";
        public const string InvalidNumericValue = "CMP-OX-NUMBER-INVALID";
        public const string InvalidIsoDateValue = "CMP-OX-ISO-DATE-INVALID";
        public const string UnsupportedCellDataType = "CMP-OX-CELL-TYPE-UNSUPPORTED";
        public const string ValueMetadataNotInterpreted = "CMP-OX-VALUE-METADATA-NOT-INTERPRETED";
        public const string SharedFormulaNotExpanded = "CMP-OX-SHARED-FORMULA-NOT-EXPANDED";
        public const string ArrayFormulaNotExpanded = "CMP-OX-ARRAY-FORMULA-NOT-EXPANDED";
        public const string DynamicArrayNotExpanded = "CMP-OX-DYNAMIC-ARRAY-NOT-EXPANDED";
        public const string FormulaMetadataUnresolved = "CMP-OX-FORMULA-METADATA-UNRESOLVED";
        public const string ExternalFormulaDependencyNotResolved =
            "CMP-OX-EXTERNAL-FORMULA-DEPENDENCY-NOT-RESOLVED";
        public const string TableFormulaNotProjected = "CMP-OX-TABLE-FORMULA-NOT-PROJECTED";
        public const string TableFormulaPayloadMalformed = "CMP-OX-TABLE-FORMULA-PAYLOAD-MALFORMED";
        public const string DataTableFormulaUnsupported = "CMP-OX-DATA-TABLE-FORMULA-UNSUPPORTED";
        public const string UnknownFormulaType = "CMP-OX-FORMULA-TYPE-UNKNOWN";
        public const string EmptyFormulaText = "CMP-OX-FORMULA-TEXT-MISSING";
        public const string UnsupportedPackagePart = "CMP-OX-PACKAGE-PART-UNSUPPORTED";
        public const string PackageExternalRelationshipIgnored = "CMP-OX-PACKAGE-EXTERNAL-IGNORED";
        public const string UnmodeledWorkbookMetadata = "CMP-OX-WORKBOOK-METADATA-UNMODELED";
    }

    public enum OpenXmlSheetVisibility
    {
        Visible,
        Hidden,
        VeryHidden,
        Unknown,
    }

    public enum OpenXmlPhysicalSheetKind
    {
        Worksheet,
        Unsupported,
    }

    public enum OpenXmlFormulaForm
    {
        Normal,
        SharedMaster,
        SharedFollower,
        Array,
        DynamicArray,
        DataTable,
        Unknown,
    }

    /// <summary>
    /// Stored and safely reconstructed evidence for one physical formula cell. EffectiveFormulaText is null when
    /// the adapter could not prove a reconstruction; callers must then honor the matching coverage message.
    /// </summary>
    public sealed class OpenXmlFormulaObservation
    {
        internal OpenXmlFormulaObservation(
            string cellReference,
            OpenXmlFormulaForm form,
            string storedFormulaText,
            string effectiveFormulaText,
            string formulaRange,
            uint? sharedIndex,
            bool wasReconstructed,
            string exactMetadata,
            CanonicalValue cachedResult)
        {
            CellReference = cellReference ?? "";
            Form = form;
            StoredFormulaText = storedFormulaText ?? "";
            EffectiveFormulaText = effectiveFormulaText;
            FormulaRange = formulaRange;
            SharedIndex = sharedIndex;
            WasReconstructed = wasReconstructed;
            ExactMetadata = exactMetadata;
            CachedResult = cachedResult;
        }

        public string CellReference { get; }
        public OpenXmlFormulaForm Form { get; }
        /// <summary>Exact text physically stored in this cell's f element; shared followers normally store none.</summary>
        public string StoredFormulaText { get; }
        /// <summary>Stored text or a proven shared-formula reconstruction; null means reconstruction is unsupported.</summary>
        public string EffectiveFormulaText { get; }
        public string FormulaRange { get; }
        public uint? SharedIndex { get; }
        public bool WasReconstructed { get; }
        public string ExactMetadata { get; }
        /// <summary>Last stored cached result, never a newly calculated value.</summary>
        public CanonicalValue CachedResult { get; }
    }

    public enum OpenXmlTableFormulaKind
    {
        CalculatedColumn,
        TotalsRow,
    }

    /// <summary>
    /// Exact table-definition formula evidence. It is intentionally not projected onto worksheet cells because
    /// doing so requires table layout and calculation semantics beyond the stored formula payload.
    /// </summary>
    public sealed class OpenXmlTableFormulaObservation
    {
        internal OpenXmlTableFormulaObservation(
            string tableName,
            string tableRange,
            uint? columnId,
            string columnName,
            OpenXmlTableFormulaKind kind,
            string exactFormulaText,
            string exactTableMetadata,
            string exactColumnMetadata,
            string exactMetadata,
            string partUri)
        {
            TableName = tableName ?? "";
            TableRange = tableRange ?? "";
            ColumnId = columnId;
            ColumnName = columnName ?? "";
            Kind = kind;
            ExactFormulaText = exactFormulaText ?? "";
            ExactTableMetadata = exactTableMetadata;
            ExactColumnMetadata = exactColumnMetadata;
            ExactMetadata = exactMetadata;
            PartUri = partUri ?? "";
        }

        public string TableName { get; }
        public string TableRange { get; }
        public uint? ColumnId { get; }
        public string ColumnName { get; }
        public OpenXmlTableFormulaKind Kind { get; }
        public string ExactFormulaText { get; }
        public string ExactTableMetadata { get; }
        public string ExactColumnMetadata { get; }
        public string ExactMetadata { get; }
        public string PartUri { get; }
    }

    /// <summary>
    /// A stable, user-reportable warning that the adapter could not make an exact formula/value observation.
    /// Callers must not report equality for a scope whose coverage contains an equality-affecting message.
    /// </summary>
    public sealed class OpenXmlCoverageMessage
    {
        internal OpenXmlCoverageMessage(
            string code,
            string message,
            bool affectsExactComparison,
            bool affectsExactCellComparison,
            string sheetName,
            string cellReference,
            string partUri)
        {
            Code = code ?? "";
            Message = message ?? "";
            AffectsExactComparison = affectsExactComparison;
            AffectsExactCellComparison = affectsExactCellComparison;
            SheetName = sheetName;
            CellReference = cellReference;
            PartUri = partUri;
        }

        public string Code { get; }
        public string Message { get; }
        public bool AffectsExactComparison { get; }
        /// <summary>Whether the finding prevents an exact formula/value-cell equality conclusion.</summary>
        public bool AffectsExactCellComparison { get; }
        public string SheetName { get; }
        public string CellReference { get; }
        public string PartUri { get; }
    }

    /// <summary>Read-only metadata and canonical cell contents for one physical sheet occurrence.</summary>
    public sealed class OpenXmlWorksheetSnapshot
    {
        internal OpenXmlWorksheetSnapshot(
            string name,
            uint? sheetId,
            string relationshipId,
            string partUri,
            string codeName,
            OpenXmlSheetVisibility visibility,
            OpenXmlPhysicalSheetKind kind,
            RevisionSnapshot snapshot,
            IList<OpenXmlFormulaObservation> formulaObservations,
            IList<OpenXmlTableFormulaObservation> tableFormulaObservations,
            IList<OpenXmlCoverageMessage> coverageMessages)
        {
            Name = name ?? "";
            SheetId = sheetId;
            RelationshipId = relationshipId ?? "";
            PartUri = partUri ?? "";
            CodeName = codeName;
            Visibility = visibility;
            Kind = kind;
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
            FormulaObservations = new ReadOnlyCollection<OpenXmlFormulaObservation>(
                new List<OpenXmlFormulaObservation>(
                    formulaObservations ?? Array.Empty<OpenXmlFormulaObservation>()));
            TableFormulaObservations = new ReadOnlyCollection<OpenXmlTableFormulaObservation>(
                new List<OpenXmlTableFormulaObservation>(
                    tableFormulaObservations ?? Array.Empty<OpenXmlTableFormulaObservation>()));
            CoverageMessages = new ReadOnlyCollection<OpenXmlCoverageMessage>(
                new List<OpenXmlCoverageMessage>(coverageMessages ?? Array.Empty<OpenXmlCoverageMessage>()));
        }

        public string Name { get; }
        public uint? SheetId { get; }
        public string RelationshipId { get; }
        public string PartUri { get; }
        public string CodeName { get; }
        public OpenXmlSheetVisibility Visibility { get; }
        public OpenXmlPhysicalSheetKind Kind { get; }
        public RevisionSnapshot Snapshot { get; }
        public IReadOnlyList<OpenXmlFormulaObservation> FormulaObservations { get; }
        public IReadOnlyList<OpenXmlTableFormulaObservation> TableFormulaObservations { get; }
        public IReadOnlyList<OpenXmlCoverageMessage> CoverageMessages { get; }

        /// <summary>
        /// False means the visible canonical cells remain useful, but an equality result for this sheet would be
        /// unsafe without resolving its coverage messages.
        /// </summary>
        public bool CanAssertExactCellEquality
        {
            get
            {
                for (int i = 0; i < CoverageMessages.Count; i++)
                    if (CoverageMessages[i].AffectsExactCellComparison)
                        return false;
                return true;
            }
        }

        /// <summary>
        /// Whether the exact stored formula/value payloads are fully covered. This does not claim semantic
        /// equality of dependencies such as defined names, tables, calculation settings, or external links.
        /// </summary>
        public bool CanAssertExactFormulaValueEquality => CanAssertExactCellEquality;

        /// <summary>
        /// Whether the canonical cells contain exact stored formula text, exact formula metadata, and stored cached
        /// values suitable for a diagnostic review. This is deliberately weaker than
        /// <see cref="CanAssertExactCellEquality"/>: a single-cell dynamic-array formula can be reviewed exactly as
        /// stored even though the adapter cannot assert recalculated spill semantics or authorize a merge.
        /// </summary>
        public bool CanProjectExactStoredFormulaValueEvidence
        {
            get
            {
                bool hasDynamicCoverage = false;
                for (int i = 0; i < CoverageMessages.Count; i++)
                {
                    OpenXmlCoverageMessage message = CoverageMessages[i];
                    if (!message.AffectsExactCellComparison) continue;
                    if (!string.Equals(
                            message.Code,
                            OpenXmlCoverageCodes.DynamicArrayNotExpanded,
                            StringComparison.Ordinal))
                        return false;
                    hasDynamicCoverage = true;
                }

                if (!hasDynamicCoverage) return true;

                bool observedDynamicFormula = false;
                for (int i = 0; i < FormulaObservations.Count; i++)
                {
                    OpenXmlFormulaObservation observation = FormulaObservations[i];
                    if (observation.Form != OpenXmlFormulaForm.DynamicArray) continue;
                    observedDynamicFormula = true;
                    if (string.IsNullOrEmpty(observation.EffectiveFormulaText) ||
                        string.IsNullOrEmpty(observation.ExactMetadata) ||
                        !IsSingleCellFormulaRange(observation.CellReference, observation.FormulaRange))
                        return false;
                }
                return observedDynamicFormula;
            }
        }

        private static bool IsSingleCellFormulaRange(string cellReference, string formulaRange)
        {
            if (!OpenXmlSharedFormulaTranslator.TryParseCellReference(
                    cellReference,
                    out int cellRow,
                    out int cellColumn) ||
                !OpenXmlSharedFormulaTranslator.TryParseRange(
                    formulaRange,
                    out int firstRow,
                    out int firstColumn,
                    out int lastRow,
                    out int lastColumn))
                return false;
            return firstRow == cellRow && firstColumn == cellColumn &&
                   lastRow == cellRow && lastColumn == cellColumn;
        }

        internal bool CanAssertExactArtifactEquality
        {
            get
            {
                for (int i = 0; i < CoverageMessages.Count; i++)
                    if (CoverageMessages[i].AffectsExactComparison)
                        return false;
                return true;
            }
        }
    }

    /// <summary>
    /// Immutable, Excel-free observation of an OOXML workbook. The byte content ID covers the exact input
    /// package, including parts the adapter deliberately does not inspect (for example VBA).
    /// </summary>
    public sealed class OpenXmlWorkbookSnapshot
    {
        internal OpenXmlWorkbookSnapshot(
            string revisionId,
            string byteContentId,
            ExcelDateSystem dateSystem,
            bool hasVbaProject,
            IList<OpenXmlWorksheetSnapshot> sheets,
            IList<OpenXmlCoverageMessage> coverageMessages)
        {
            RevisionId = revisionId ?? "";
            ByteContentId = byteContentId ?? "";
            DateSystem = dateSystem;
            HasVbaProject = hasVbaProject;
            Sheets = new ReadOnlyCollection<OpenXmlWorksheetSnapshot>(
                new List<OpenXmlWorksheetSnapshot>(sheets ?? Array.Empty<OpenXmlWorksheetSnapshot>()));
            CoverageMessages = new ReadOnlyCollection<OpenXmlCoverageMessage>(
                new List<OpenXmlCoverageMessage>(coverageMessages ?? Array.Empty<OpenXmlCoverageMessage>()));
        }

        public string RevisionId { get; }
        public string ByteContentId { get; }
        public ExcelDateSystem DateSystem { get; }
        public bool HasVbaProject { get; }
        public IReadOnlyList<OpenXmlWorksheetSnapshot> Sheets { get; }
        public IReadOnlyList<OpenXmlCoverageMessage> CoverageMessages { get; }

        public bool CanAssertExactEquality
        {
            get
            {
                for (int i = 0; i < CoverageMessages.Count; i++)
                    if (CoverageMessages[i].AffectsExactComparison)
                        return false;

                for (int i = 0; i < Sheets.Count; i++)
                    if (!Sheets[i].CanAssertExactArtifactEquality)
                        return false;

                return true;
            }
        }

        public bool CanAssertExactCellEquality
        {
            get
            {
                for (int i = 0; i < CoverageMessages.Count; i++)
                    if (CoverageMessages[i].AffectsExactCellComparison)
                        return false;

                for (int i = 0; i < Sheets.Count; i++)
                    if (!Sheets[i].CanAssertExactCellEquality)
                        return false;

                return true;
            }
        }
    }
}
