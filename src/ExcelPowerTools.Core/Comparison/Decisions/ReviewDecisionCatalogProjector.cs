using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.History;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>
    /// Portable byte identity and display label for one revision occurrence supplied to the decision projector.
    /// SourceLabel must already be a strict basename; paths and URI-like labels are rejected so a decision catalog
    /// cannot retain private location or credential material.
    /// </summary>
    public sealed class ReviewDecisionRevisionEvidence
    {
        public ReviewDecisionRevisionEvidence(
            string revisionId,
            string byteContentId,
            string sourceLabel)
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (!StableId.IsSha256(byteContentId))
                throw new ArgumentException(
                    "A revision byte-content identity must be a SHA-256 value.",
                    nameof(byteContentId));

            RevisionId = revisionId;
            ByteContentId = byteContentId.ToUpperInvariant();
            SourceLabel = PortableSourceLabel(sourceLabel);
        }

        public string RevisionId { get; }
        public string ByteContentId { get; }
        public string SourceLabel { get; }

        private static string PortableSourceLabel(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException(
                    "A non-blank portable source label is required.",
                    nameof(value));

            string result = value.Trim();
            bool driveRelative = result.Length >= 2 &&
                                 char.IsLetter(result[0]) &&
                                 result[1] == Path.VolumeSeparatorChar;
            bool pathLike = Path.IsPathRooted(result) ||
                            driveRelative ||
                            result.IndexOf('/') >= 0 ||
                            result.IndexOf('\\') >= 0 ||
                            string.Equals(result, ".", StringComparison.Ordinal) ||
                            string.Equals(result, "..", StringComparison.Ordinal);
            bool uriLike = Uri.TryCreate(result, UriKind.Absolute, out Uri _) ||
                           result.IndexOf('?') >= 0 ||
                           result.IndexOf('#') >= 0;
            bool invalidPortableCharacter = result.Any(character =>
                char.IsControl(character) ||
                character == '<' || character == '>' || character == '"' ||
                character == ':' || character == '|' || character == '*');
            if (result.Length == 0 || pathLike || uriLike || invalidPortableCharacter)
                throw new ArgumentException(
                    "A source label must be a privacy-safe portable basename, not a path, URI, or ambiguous label.",
                    nameof(value));
            return result;
        }
    }

    /// <summary>
    /// Projects the complete aligned project/history authority into an immutable decision catalog. Every history
    /// subject has exactly one observation per revision occurrence. Physical provenance is recorded only where an
    /// exact aligned cell exists; missing states retain their proven logical sheet without inventing an address.
    /// </summary>
    public static class ReviewDecisionCatalogProjector
    {
        public static ReviewDecisionCatalog Project(
            OpenXmlComparisonProjectResult result,
            IEnumerable<ReviewDecisionRevisionEvidence> revisionEvidence,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (result == null) throw new ArgumentNullException(nameof(result));
            cancellationToken.ThrowIfCancellationRequested();
            if (!result.IsComplete || result.History == null)
                throw new InvalidOperationException(
                    "A decision catalog requires a complete comparison project and logical history.");
            Require(result.ProjectId, "The comparison project ID is missing.");
            Require(result.EffectiveMappingGenerationId,
                "The effective mapping generation ID is missing.");

            LogicalHistoryResult history = result.History;
            if (!string.Equals(result.ProjectId, history.ProjectId, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The logical history belongs to a different comparison project.");
            if (!string.Equals(
                    result.EffectiveMappingGenerationId,
                    history.MappingGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The logical history does not use the project's effective mapping generation.");

            List<OpenXmlAlignedProjectRevision> revisions = ValidateAndOrderRevisions(result, history);
            IReadOnlyDictionary<string, ReviewDecisionRevisionEvidence> evidenceByRevision =
                ValidateRevisionEvidence(revisions, revisionEvidence);
            var evidenceGeneration = new ReviewEvidenceGeneration(
                result.ProjectId,
                result.EffectiveMappingGenerationId,
                revisions.Select(item => new ReviewEvidenceRevisionIdentity(
                    item.Source.RevisionId,
                    item.Source.Order,
                    evidenceByRevision[item.Source.RevisionId].ByteContentId)));

            List<LogicalSubjectHistory> historySubjects = ValidateHistorySubjects(
                history,
                revisions,
                cancellationToken);
            var expectedKeys = new HashSet<string>(
                historySubjects.Select(item => item.LogicalSubjectKey),
                StringComparer.Ordinal);
            IReadOnlyDictionary<string, IReadOnlyDictionary<string, PhysicalOccurrence>> occurrencesByRevision =
                ResolveAllOccurrences(revisions, expectedKeys, cancellationToken);

            var subjects = new List<ReviewDecisionSubject>(historySubjects.Count);
            var observations = new List<Observation>(checked(historySubjects.Count * revisions.Count));
            foreach (LogicalSubjectHistory subject in historySubjects)
            {
                cancellationToken.ThrowIfCancellationRequested();
                SubjectCoordinates coordinates = ResolveSubjectCoordinates(
                    subject,
                    revisions,
                    occurrencesByRevision);
                bool hasMasterTarget = occurrencesByRevision[history.MasterRevisionId]
                    .ContainsKey(subject.LogicalSubjectKey);
                subjects.Add(new ReviewDecisionSubject(
                    subject.SubjectId,
                    subject.LogicalSubjectKey,
                    ToComparisonAspect(subject.Aspect),
                    SubjectPolicy(subject.Aspect, hasMasterTarget)));

                IReadOnlyDictionary<string, HistoryObservation> historyByRevision =
                    subject.Observations.ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
                foreach (OpenXmlAlignedProjectRevision revision in revisions)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string revisionId = revision.Source.RevisionId;
                    HistoryObservation historyObservation = historyByRevision[revisionId];
                    occurrencesByRevision[revisionId].TryGetValue(
                        subject.LogicalSubjectKey,
                        out PhysicalOccurrence occurrence);
                    ValidateOccurrenceState(subject, historyObservation, occurrence);

                    ProjectedObservation exact = ProjectExactObservation(subject.Aspect, historyObservation);
                    ReviewDecisionRevisionEvidence revisionMetadata = evidenceByRevision[revisionId];
                    var provenance = new ObservationProvenance(
                        sourceLabel: revisionMetadata.SourceLabel,
                        sheetName: occurrence == null ? "" : occurrence.PhysicalSheetName,
                        address: occurrence == null ? "" : occurrence.Address,
                        physicalRow: occurrence == null ? 0 : occurrence.PhysicalRow,
                        physicalColumn: occurrence == null ? 0 : occurrence.PhysicalColumn,
                        logicalSheetId: coordinates.LogicalSheetId);
                    string observationId = ComparisonIds.Observation(
                        subject.SubjectId,
                        revisionId,
                        exact.ExactType,
                        historyObservation.StateId);
                    observations.Add(new Observation(
                        observationId,
                        subject.SubjectId,
                        revisionId,
                        exact.ExactType,
                        historyObservation.StateId,
                        formula: exact.Formula,
                        value: exact.Value,
                        displayText: exact.DisplayText,
                        provenance: provenance));
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            List<ReviewDecisionGroup> groups = FreezeHistoryGroups(
                history,
                subjects,
                cancellationToken);
            return new ReviewDecisionCatalog(
                evidenceGeneration,
                history.MasterRevisionId,
                subjects,
                observations,
                groups);
        }

        private static List<OpenXmlAlignedProjectRevision> ValidateAndOrderRevisions(
            OpenXmlComparisonProjectResult result,
            LogicalHistoryResult history)
        {
            if (result.Revisions == null || result.Revisions.Count < 2)
                throw new InvalidOperationException(
                    "A decision catalog requires at least two aligned revision occurrences.");

            var revisions = new List<OpenXmlAlignedProjectRevision>(result.Revisions.Count);
            var revisionIds = new HashSet<string>(StringComparer.Ordinal);
            var orders = new HashSet<int>();
            foreach (OpenXmlAlignedProjectRevision revision in result.Revisions)
            {
                if (revision == null || revision.Source == null)
                    throw new InvalidOperationException(
                        "The comparison project contains a null aligned revision.");
                if (!revision.IsComplete || revision.AlignedSnapshot == null)
                    throw new InvalidOperationException(
                        "Every decision revision must have a complete aligned snapshot: " +
                        revision.Source.RevisionId + ".");
                Require(revision.Source.RevisionId, "An aligned revision ID is missing.");
                if (!revisionIds.Add(revision.Source.RevisionId))
                    throw new InvalidOperationException(
                        "The comparison project repeats revision ID " + revision.Source.RevisionId + ".");
                if (!orders.Add(revision.Source.Order))
                    throw new InvalidOperationException(
                        "The comparison project repeats revision order " +
                        revision.Source.Order.ToString(CultureInfo.InvariantCulture) + ".");
                if (!string.Equals(
                        revision.Source.RevisionId,
                        revision.AlignedSnapshot.RevisionId,
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "An aligned snapshot has the wrong revision occurrence ID.");
                if (!StableId.IsSha256(revision.Source.Snapshot.ByteContentId))
                    throw new InvalidOperationException(
                        "Revision " + revision.Source.RevisionId +
                        " lacks a valid byte-content SHA-256 identity.");
                revisions.Add(revision);
            }

            revisions.Sort((left, right) =>
            {
                int order = left.Source.Order.CompareTo(right.Source.Order);
                return order != 0
                    ? order
                    : StringComparer.Ordinal.Compare(left.Source.RevisionId, right.Source.RevisionId);
            });
            List<OpenXmlAlignedProjectRevision> masters = revisions
                .Where(item => item.Source.IsMaster)
                .ToList();
            if (masters.Count != 1 || !string.Equals(
                    masters[0].Source.RevisionId,
                    history.MasterRevisionId,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The complete project and logical history do not identify the same single master revision.");
            return revisions;
        }

        private static IReadOnlyDictionary<string, ReviewDecisionRevisionEvidence> ValidateRevisionEvidence(
            IReadOnlyList<OpenXmlAlignedProjectRevision> revisions,
            IEnumerable<ReviewDecisionRevisionEvidence> revisionEvidence)
        {
            if (revisionEvidence == null) throw new ArgumentNullException(nameof(revisionEvidence));
            var byId = new Dictionary<string, ReviewDecisionRevisionEvidence>(StringComparer.Ordinal);
            foreach (ReviewDecisionRevisionEvidence evidence in revisionEvidence)
            {
                if (evidence == null)
                    throw new ArgumentException(
                        "Revision evidence cannot contain null entries.",
                        nameof(revisionEvidence));
                if (byId.ContainsKey(evidence.RevisionId))
                    throw new ArgumentException(
                        "Revision evidence repeats occurrence " + evidence.RevisionId + ".",
                        nameof(revisionEvidence));
                byId.Add(evidence.RevisionId, evidence);
            }
            if (byId.Count != revisions.Count)
                throw new ArgumentException(
                    "Revision evidence must contain exactly the project revision occurrences.",
                    nameof(revisionEvidence));

            foreach (OpenXmlAlignedProjectRevision revision in revisions)
            {
                string revisionId = revision.Source.RevisionId;
                if (!byId.TryGetValue(revisionId, out ReviewDecisionRevisionEvidence evidence))
                    throw new ArgumentException(
                        "Revision evidence is missing occurrence " + revisionId + ".",
                        nameof(revisionEvidence));
                string snapshotHash = revision.Source.Snapshot.ByteContentId.ToUpperInvariant();
                if (!string.Equals(evidence.ByteContentId, snapshotHash, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Revision evidence byte identity disagrees with the analyzed snapshot for " +
                        revisionId + ".");
            }
            return new ReadOnlyDictionary<string, ReviewDecisionRevisionEvidence>(byId);
        }

        private static List<LogicalSubjectHistory> ValidateHistorySubjects(
            LogicalHistoryResult history,
            IReadOnlyList<OpenXmlAlignedProjectRevision> revisions,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (history.Subjects == null)
                throw new InvalidOperationException("The logical history subject collection is missing.");
            var subjects = new List<LogicalSubjectHistory>(history.Subjects.Count);
            var subjectIds = new HashSet<string>(StringComparer.Ordinal);
            var semanticIdentities = new HashSet<string>(StringComparer.Ordinal);
            var expectedRevisionIds = new HashSet<string>(
                revisions.Select(item => item.Source.RevisionId),
                StringComparer.Ordinal);
            IReadOnlyDictionary<string, OpenXmlAlignedProjectRevision> revisionsById =
                revisions.ToDictionary(item => item.Source.RevisionId, StringComparer.Ordinal);

            foreach (LogicalSubjectHistory subject in history.Subjects)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (subject == null)
                    throw new InvalidOperationException("The logical history contains a null subject.");
                Require(subject.SubjectId, "A logical history subject ID is missing.");
                Require(subject.LogicalSubjectKey, "A logical history subject key is missing.");
                if (!subjectIds.Add(subject.SubjectId))
                    throw new InvalidOperationException(
                        "The logical history repeats subject ID " + subject.SubjectId + ".");
                string semanticIdentity = StableId.Compute(
                    "review-decision-history-subject-semantics-v1",
                    subject.LogicalSubjectKey,
                    ((int)subject.Aspect).ToString(CultureInfo.InvariantCulture));
                if (!semanticIdentities.Add(semanticIdentity))
                    throw new InvalidOperationException(
                        "The logical history repeats a logical subject/aspect identity: " +
                        subject.LogicalSubjectKey + ".");
                if (!Enum.IsDefined(typeof(CellChangeAspect), subject.Aspect))
                    throw new InvalidOperationException("The logical history contains an unknown subject aspect.");
                if (subject.Observations == null || subject.Observations.Count != revisions.Count)
                    throw new InvalidOperationException(
                        "Every logical subject must retain one observation for every revision occurrence: " +
                        subject.SubjectId + ".");

                var seenRevisions = new HashSet<string>(StringComparer.Ordinal);
                int observationIndex = 0;
                foreach (HistoryObservation observation in subject.Observations)
                {
                    if ((observationIndex++ & 0xFF) == 0)
                        cancellationToken.ThrowIfCancellationRequested();
                    if (observation == null ||
                        !expectedRevisionIds.Contains(observation.RevisionId) ||
                        !seenRevisions.Add(observation.RevisionId))
                        throw new InvalidOperationException(
                            "Logical subject " + subject.SubjectId +
                            " has missing, duplicate, or unknown revision observations.");
                    OpenXmlAlignedProjectRevision revision = revisionsById[observation.RevisionId];
                    if (observation.RevisionOrder != revision.Source.Order)
                        throw new InvalidOperationException(
                            "A logical history observation has an inconsistent revision order.");
                    Require(observation.StateId, "A logical history observation state ID is missing.");
                }
                subjects.Add(subject);
            }
            subjects.Sort((left, right) => StringComparer.Ordinal.Compare(left.SubjectId, right.SubjectId));
            return subjects;
        }

        private static IReadOnlyDictionary<string, IReadOnlyDictionary<string, PhysicalOccurrence>>
            ResolveAllOccurrences(
                IReadOnlyList<OpenXmlAlignedProjectRevision> revisions,
                ISet<string> expectedLogicalKeys,
                CancellationToken cancellationToken)
        {
            var result = new Dictionary<string, IReadOnlyDictionary<string, PhysicalOccurrence>>(
                StringComparer.Ordinal);
            foreach (OpenXmlAlignedProjectRevision revision in revisions)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var occurrences = new Dictionary<string, PhysicalOccurrence>(StringComparer.Ordinal);
                foreach (OpenXmlAlignedSheetRevision sheet in revision.Sheets
                             .OrderBy(item => item.LogicalSheetId, StringComparer.Ordinal)
                             .ThenBy(item => item.PhysicalSheetKey, StringComparer.Ordinal))
                {
                    Require(sheet.LogicalSheetId, "An aligned sheet logical ID is missing.");
                    Require(sheet.PhysicalSheetName, "An aligned physical sheet name is missing.");
                    string expectedSheetRevisionPrefix = revision.Source.RevisionId + "/sheet:";
                    if (sheet.Grid == null || sheet.Grid.Source == null ||
                        !sheet.Grid.Source.RevisionId.StartsWith(
                            expectedSheetRevisionPrefix,
                            StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            "An aligned sheet grid belongs to a different revision occurrence.");
                    IReadOnlyDictionary<int, string> rows =
                        ToPhysicalIdentityMap(sheet.Grid, sheet.Rows, true);
                    IReadOnlyDictionary<int, string> columns =
                        ToPhysicalIdentityMap(sheet.Grid, sheet.Columns, false);
                    foreach (AlignedCanonicalCellProjection cell in sheet.Grid
                                 .CreateAlignedCells(sheet.LogicalSheetId, rows, columns, cancellationToken))
                    {
                        if (!expectedLogicalKeys.Contains(cell.Cell.SubjectKey)) continue;
                        int rowIndex = checked(cell.PhysicalRow - sheet.Grid.FirstPhysicalRow);
                        int columnIndex = checked(cell.PhysicalColumn - sheet.Grid.FirstPhysicalColumn);
                        if (!sheet.Rows.AssignmentsByPhysicalIndex.TryGetValue(
                                rowIndex,
                                out LogicalAxisIdentityAssignment row) ||
                            !sheet.Columns.AssignmentsByPhysicalIndex.TryGetValue(
                                columnIndex,
                                out LogicalAxisIdentityAssignment column))
                            throw new InvalidOperationException(
                                "Aligned evidence lacks an exact row or column assignment in revision " +
                                revision.Source.RevisionId + ".");
                        if (!string.Equals(row.LogicalIdentityId, cell.RowIdentity, StringComparison.Ordinal) ||
                            !string.Equals(column.LogicalIdentityId, cell.ColumnIdentity, StringComparison.Ordinal))
                            throw new InvalidOperationException(
                                "Aligned cell evidence disagrees with its axis identities in revision " +
                                revision.Source.RevisionId + ".");
                        string expectedKey = StableId.Compute(
                            "aligned-cell-subject-v1",
                            sheet.LogicalSheetId,
                            cell.RowIdentity,
                            cell.ColumnIdentity);
                        if (!string.Equals(expectedKey, cell.Cell.SubjectKey, StringComparison.Ordinal))
                            throw new InvalidOperationException(
                                "Aligned cell evidence has an inconsistent logical subject key.");

                        var occurrence = new PhysicalOccurrence(
                            cell.Cell,
                            sheet.LogicalSheetId,
                            cell.RowIdentity,
                            cell.ColumnIdentity,
                            sheet.PhysicalSheetName,
                            ToA1(cell.PhysicalRow, cell.PhysicalColumn),
                            cell.PhysicalRow,
                            cell.PhysicalColumn);
                        if (occurrences.ContainsKey(cell.Cell.SubjectKey))
                            throw new InvalidOperationException(
                                "Revision " + revision.Source.RevisionId +
                                " contains more than one physical occurrence for logical subject " +
                                cell.Cell.SubjectKey + ".");
                        occurrences.Add(cell.Cell.SubjectKey, occurrence);
                    }
                }
                result.Add(
                    revision.Source.RevisionId,
                    new ReadOnlyDictionary<string, PhysicalOccurrence>(occurrences));
            }
            return new ReadOnlyDictionary<string, IReadOnlyDictionary<string, PhysicalOccurrence>>(result);
        }

        private static SubjectCoordinates ResolveSubjectCoordinates(
            LogicalSubjectHistory subject,
            IReadOnlyList<OpenXmlAlignedProjectRevision> revisions,
            IReadOnlyDictionary<string, IReadOnlyDictionary<string, PhysicalOccurrence>> occurrencesByRevision)
        {
            PhysicalOccurrence first = null;
            foreach (OpenXmlAlignedProjectRevision revision in revisions)
            {
                if (!occurrencesByRevision[revision.Source.RevisionId].TryGetValue(
                        subject.LogicalSubjectKey,
                        out PhysicalOccurrence occurrence))
                    continue;
                if (first == null)
                {
                    first = occurrence;
                    continue;
                }
                if (!string.Equals(first.LogicalSheetId, occurrence.LogicalSheetId, StringComparison.Ordinal) ||
                    !string.Equals(first.RowIdentity, occurrence.RowIdentity, StringComparison.Ordinal) ||
                    !string.Equals(first.ColumnIdentity, occurrence.ColumnIdentity, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Revision occurrences disagree on the aligned identity of subject " +
                        subject.SubjectId + ".");
            }
            if (first == null)
                throw new InvalidOperationException(
                    "No exact aligned occurrence exists in any revision for logical subject " +
                    subject.SubjectId + ".");
            return new SubjectCoordinates(
                first.LogicalSheetId,
                first.RowIdentity,
                first.ColumnIdentity);
        }

        private static void ValidateOccurrenceState(
            LogicalSubjectHistory subject,
            HistoryObservation observation,
            PhysicalOccurrence occurrence)
        {
            if (observation.IsPhysicallyPresent != (occurrence != null))
                throw new InvalidOperationException(
                    "Logical history presence evidence disagrees with the aligned occurrence for subject " +
                    subject.SubjectId + " in revision " + observation.RevisionId + ".");

            CanonicalCellState cell = occurrence == null
                ? CanonicalCellState.Blank(subject.LogicalSubjectKey)
                : occurrence.Cell;
            ExtractAspectState(
                subject.Aspect,
                observation.IsPhysicallyPresent,
                cell,
                out string expectedStateId,
                out string expectedText,
                out CanonicalValue expectedValue);
            if (!string.Equals(observation.StateId, expectedStateId, StringComparison.Ordinal) ||
                !string.Equals(observation.ExactText, expectedText, StringComparison.Ordinal) ||
                !CanonicalValuesEqual(observation.ExactValue, expectedValue))
                throw new InvalidOperationException(
                    "Logical history exact evidence disagrees with the aligned cell for subject " +
                    subject.SubjectId + " in revision " + observation.RevisionId + ".");
        }

        private static void ExtractAspectState(
            CellChangeAspect aspect,
            bool present,
            CanonicalCellState cell,
            out string stateId,
            out string text,
            out CanonicalValue value)
        {
            text = null;
            value = null;
            string payload;
            switch (aspect)
            {
                case CellChangeAspect.Presence:
                    text = present ? "present" : "missing";
                    payload = text;
                    break;
                case CellChangeAspect.Formula:
                    text = cell.FormulaText;
                    payload = text == null ? "<no-formula>" : "formula:" + text;
                    break;
                case CellChangeAspect.FormulaStructure:
                    text = cell.FormulaMetadata;
                    payload = text == null
                        ? "<no-formula-metadata>"
                        : "formula-metadata:" + text;
                    break;
                case CellChangeAspect.Value:
                    value = cell.ConstantValue;
                    payload = ValuePayload(value);
                    break;
                case CellChangeAspect.CachedResult:
                    value = cell.CachedResult;
                    payload = value == null ? "<no-cache>" : ValuePayload(value);
                    break;
                default:
                    throw new InvalidOperationException("The logical history contains an unsupported aspect.");
            }
            stateId = StableId.Compute(
                "history-aspect-state-v1",
                ((int)aspect).ToString(CultureInfo.InvariantCulture),
                payload);
        }

        private static ProjectedObservation ProjectExactObservation(
            CellChangeAspect aspect,
            HistoryObservation observation)
        {
            if (!observation.IsPhysicallyPresent)
                return new ProjectedObservation(
                    ObservationExactType.Missing,
                    null,
                    CanonicalScalar.Missing,
                    "");

            switch (aspect)
            {
                case CellChangeAspect.Formula:
                    return observation.ExactText == null
                        ? new ProjectedObservation(
                            ObservationExactType.Blank,
                            null,
                            CanonicalScalar.Blank,
                            "")
                        : new ProjectedObservation(
                            ObservationExactType.Formula,
                            observation.ExactText,
                            null,
                            observation.ExactText);
                case CellChangeAspect.Value:
                    return FromCanonicalValue(observation.ExactValue);
                case CellChangeAspect.CachedResult:
                    return observation.ExactValue == null
                        ? new ProjectedObservation(
                            ObservationExactType.Missing,
                            null,
                            CanonicalScalar.Missing,
                            "")
                        : FromCanonicalValue(observation.ExactValue);
                case CellChangeAspect.Presence:
                case CellChangeAspect.FormulaStructure:
                    if (observation.ExactText == null)
                        return new ProjectedObservation(
                            ObservationExactType.Blank,
                            null,
                            CanonicalScalar.Blank,
                            "");
                    return new ProjectedObservation(
                        ObservationExactType.String,
                        null,
                        new CanonicalScalar(ObservationExactType.String, observation.ExactText),
                        observation.ExactText);
                default:
                    throw new InvalidOperationException("The logical history contains an unsupported aspect.");
            }
        }

        private static ProjectedObservation FromCanonicalValue(CanonicalValue value)
        {
            if (value == null)
                throw new InvalidOperationException("A value history observation lacks exact canonical evidence.");
            ObservationExactType type;
            switch (value.Kind)
            {
                case CanonicalValueKind.Blank: type = ObservationExactType.Blank; break;
                case CanonicalValueKind.Text: type = ObservationExactType.String; break;
                case CanonicalValueKind.Number: type = ObservationExactType.Number; break;
                case CanonicalValueKind.Boolean: type = ObservationExactType.Boolean; break;
                case CanonicalValueKind.Error: type = ObservationExactType.Error; break;
                case CanonicalValueKind.DateSerial: type = ObservationExactType.DateSerial; break;
                case CanonicalValueKind.IsoDateTime: type = ObservationExactType.IsoDateTime; break;
                default:
                    throw new InvalidOperationException("The canonical value kind is not supported for review.");
            }
            CanonicalScalar scalar = type == ObservationExactType.Blank
                ? CanonicalScalar.Blank
                : new CanonicalScalar(type, value.ExactText, value.DateSystem);
            return new ProjectedObservation(type, null, scalar, value.ExactText);
        }

        private static ComparisonAspect ToComparisonAspect(CellChangeAspect aspect)
        {
            switch (aspect)
            {
                case CellChangeAspect.Presence: return ComparisonAspect.Coverage;
                case CellChangeAspect.Formula: return ComparisonAspect.Formula;
                case CellChangeAspect.FormulaStructure: return ComparisonAspect.Structure;
                case CellChangeAspect.Value: return ComparisonAspect.Value;
                case CellChangeAspect.CachedResult: return ComparisonAspect.CachedResult;
                default:
                    throw new InvalidOperationException("The logical history contains an unsupported aspect.");
            }
        }

        private static ReviewDecisionSubjectPolicy SubjectPolicy(
            CellChangeAspect aspect,
            bool hasMasterTarget)
        {
            if (aspect != CellChangeAspect.Formula && aspect != CellChangeAspect.Value)
                return ReviewDecisionSubjectPolicy.ReviewOnly;
            return hasMasterTarget
                ? ReviewDecisionSubjectPolicy.SelectableAndMergeable
                : ReviewDecisionSubjectPolicy.Selectable;
        }

        private static List<ReviewDecisionGroup> FreezeHistoryGroups(
            LogicalHistoryResult history,
            IReadOnlyList<ReviewDecisionSubject> subjects,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var subjectIds = new HashSet<string>(
                subjects.Select(item => item.SubjectId),
                StringComparer.Ordinal);
            var historiesById = new Dictionary<string, LogicalSubjectHistory>(StringComparer.Ordinal);
            foreach (LogicalSubjectHistory subject in history.Subjects)
            {
                cancellationToken.ThrowIfCancellationRequested();
                historiesById.Add(subject.SubjectId, subject);
            }
            var groups = new List<ReviewDecisionGroup>();
            var groupIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (HistoryChangeGroup group in history.Groups ?? Array.Empty<HistoryChangeGroup>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (group == null || string.IsNullOrWhiteSpace(group.GroupId) ||
                    group.SubjectIds == null || group.SubjectIds.Count == 0)
                    throw new InvalidOperationException("The logical history contains an invalid change group.");
                if (group.SubjectIds.Any(item => !subjectIds.Contains(item)))
                    throw new InvalidOperationException(
                        "A logical history group references an unknown decision subject.");
                if (!groupIds.Add(group.GroupId))
                    throw new InvalidOperationException("The logical history repeats a change-group ID.");
                string[] orderedSubjectIds = group.SubjectIds
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(item => item, StringComparer.Ordinal)
                    .ToArray();
                if (orderedSubjectIds.Length != group.SubjectIds.Count)
                    throw new InvalidOperationException(
                        "A logical history group repeats a subject occurrence.");
                string expectedGroupId = ComparisonIds.Group(
                    history.MappingGenerationId,
                    "identical-exact-transition-v1",
                    orderedSubjectIds);
                if (!string.Equals(group.GroupId, expectedGroupId, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "A logical history group ID disagrees with its exact frozen membership.");
                foreach (string subjectId in orderedSubjectIds)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    LogicalSubjectHistory subject = historiesById[subjectId];
                    if (subject.Aspect != group.Aspect || !subject.Transitions.Any(transition =>
                            string.Equals(transition.GroupId, group.GroupId, StringComparison.Ordinal) &&
                            string.Equals(transition.EdgeId, group.EdgeId, StringComparison.Ordinal) &&
                            string.Equals(
                                transition.BeforeStateId,
                                group.BeforeStateId,
                                StringComparison.Ordinal) &&
                            string.Equals(
                                transition.AfterStateId,
                                group.AfterStateId,
                                StringComparison.Ordinal)))
                        throw new InvalidOperationException(
                            "A logical history group disagrees with its member transition evidence.");
                }
                groups.Add(new ReviewDecisionGroup(group.GroupId, orderedSubjectIds));
            }

            foreach (LogicalSubjectHistory subject in history.Subjects)
            {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (HistoryTransition transition in subject.Transitions)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (transition.GroupId.Length != 0 && !groupIds.Contains(transition.GroupId))
                        throw new InvalidOperationException(
                            "A logical history transition references an unknown change group.");
                }
            }
            return groups;
        }

        private static IReadOnlyDictionary<int, string> ToPhysicalIdentityMap(
            CanonicalGridProjection grid,
            LogicalAxisIdentityProjection identities,
            bool rows)
        {
            if (grid == null || identities == null)
                throw new InvalidOperationException("Aligned sheet grid or axis identities are missing.");
            var result = new Dictionary<int, string>();
            foreach (KeyValuePair<int, string> pair in identities.IdentityByPhysicalIndex
                         .OrderBy(item => item.Key))
            {
                int physicalIndex = rows
                    ? grid.PhysicalRowForTableIndex(pair.Key)
                    : grid.PhysicalColumnForTableIndex(pair.Key);
                result.Add(physicalIndex, pair.Value);
            }
            return new ReadOnlyDictionary<int, string>(result);
        }

        private static bool CanonicalValuesEqual(CanonicalValue left, CanonicalValue right) =>
            left == null ? right == null : left.Equals(right);

        private static string ValuePayload(CanonicalValue value)
        {
            if (value == null) return "<null>";
            return string.Join(":", new[]
            {
                ((int)value.Kind).ToString(CultureInfo.InvariantCulture),
                ((int)value.DateSystem).ToString(CultureInfo.InvariantCulture),
                value.ExactText,
            });
        }

        private static string ToA1(int row, int column)
        {
            if (row <= 0 || column <= 0)
                throw new InvalidOperationException("An aligned physical cell coordinate is invalid.");
            var letters = new StringBuilder();
            int remaining = column;
            while (remaining > 0)
            {
                remaining--;
                letters.Insert(0, (char)('A' + remaining % 26));
                remaining /= 26;
            }
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }

        private static void Require(string value, string message)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new InvalidOperationException(message);
        }

        private sealed class PhysicalOccurrence
        {
            internal PhysicalOccurrence(
                CanonicalCellState cell,
                string logicalSheetId,
                string rowIdentity,
                string columnIdentity,
                string physicalSheetName,
                string address,
                int physicalRow,
                int physicalColumn)
            {
                Cell = cell;
                LogicalSheetId = logicalSheetId;
                RowIdentity = rowIdentity;
                ColumnIdentity = columnIdentity;
                PhysicalSheetName = physicalSheetName;
                Address = address;
                PhysicalRow = physicalRow;
                PhysicalColumn = physicalColumn;
            }

            internal CanonicalCellState Cell { get; }
            internal string LogicalSheetId { get; }
            internal string RowIdentity { get; }
            internal string ColumnIdentity { get; }
            internal string PhysicalSheetName { get; }
            internal string Address { get; }
            internal int PhysicalRow { get; }
            internal int PhysicalColumn { get; }
        }

        private sealed class SubjectCoordinates
        {
            internal SubjectCoordinates(string logicalSheetId, string rowIdentity, string columnIdentity)
            {
                LogicalSheetId = logicalSheetId;
                RowIdentity = rowIdentity;
                ColumnIdentity = columnIdentity;
            }

            internal string LogicalSheetId { get; }
            internal string RowIdentity { get; }
            internal string ColumnIdentity { get; }
        }

        private sealed class ProjectedObservation
        {
            internal ProjectedObservation(
                ObservationExactType exactType,
                string formula,
                CanonicalScalar value,
                string displayText)
            {
                ExactType = exactType;
                Formula = formula;
                Value = value;
                DisplayText = displayText;
            }

            internal ObservationExactType ExactType { get; }
            internal string Formula { get; }
            internal CanonicalScalar Value { get; }
            internal string DisplayText { get; }
        }
    }
}
