using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.History;

namespace ExcelPowerTools.Core.Comparison.Persistence
{
    /// <summary>
    /// Display metadata for one revision occurrence. Both the portable filename and the original user-selected
    /// filesystem path are retained so the generated review can identify similarly named copies unambiguously.
    /// </summary>
    public sealed class ReviewRevisionDisplayMetadata
    {
        public ReviewRevisionDisplayMetadata(
            string revisionId,
            string label,
            int order,
            string parentRevisionId,
            string sourceFileName = "",
            string byteContentId = "",
            DateTimeOffset? modifiedUtc = null,
            string inference = "",
            string confidence = "")
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (order < 0) throw new ArgumentOutOfRangeException(nameof(order));

            RevisionId = revisionId;
            Label = label ?? "";
            Order = order;
            ParentRevisionId = parentRevisionId ?? "";
            SourcePath = sourceFileName ?? "";
            SourceBaseName = PortableBaseName(sourceFileName);
            ByteContentId = byteContentId ?? "";
            ModifiedUtc = modifiedUtc;
            Inference = inference ?? "";
            Confidence = confidence ?? "";
        }

        public string RevisionId { get; }
        public string Label { get; }
        public int Order { get; }
        public string ParentRevisionId { get; }
        public string SourceBaseName { get; }
        public string SourcePath { get; }
        public string ByteContentId { get; }
        public DateTimeOffset? ModifiedUtc { get; }
        public string Inference { get; }
        public string Confidence { get; }

        private static string PortableBaseName(string value)
        {
            value = value ?? "";
            if (value.Length == 0) return "";

            // Split both separators explicitly. This remains portable even when a Windows path is projected by
            // a non-Windows host, and likewise prevents a POSIX path from surviving on Windows.
            int slash = value.LastIndexOf('/');
            int backslash = value.LastIndexOf('\\');
            int separator = Math.Max(slash, backslash);
            string result = separator < 0 ? value : value.Substring(separator + 1);
            if (string.IsNullOrWhiteSpace(result))
                throw new ArgumentException(
                    "Source display metadata must identify a filename, not a directory.",
                    nameof(value));
            return result;
        }
    }

    /// <summary>Portable visible location and alignment evidence for one logical cell identity.</summary>
    public sealed class ReviewSubjectDisplayMetadata
    {
        public ReviewSubjectDisplayMetadata(
            string logicalSubjectKey,
            string sheetName,
            string address,
            AlignmentConfidence alignmentConfidence)
        {
            if (string.IsNullOrWhiteSpace(logicalSubjectKey))
                throw new ArgumentException("A logical subject key is required.", nameof(logicalSubjectKey));
            if (!Enum.IsDefined(typeof(AlignmentConfidence), alignmentConfidence))
                throw new ArgumentOutOfRangeException(nameof(alignmentConfidence));

            LogicalSubjectKey = logicalSubjectKey;
            SheetName = sheetName ?? "";
            Address = address ?? "";
            AlignmentConfidence = alignmentConfidence;
        }

        public string LogicalSubjectKey { get; }
        public string SheetName { get; }
        public string Address { get; }
        public AlignmentConfidence AlignmentConfidence { get; }
    }

    /// <summary>
    /// Pure projection from immutable logical history into the visible, non-authoritative Review.xlsx state.
    /// Each history subject becomes one baseline-or-first to master row. Exact history remains in stateJson.
    /// </summary>
    public static class ReviewHistoryProjector
    {
        public const string EquivalentFlag = "Equivalent";
        public const string RevertedFlag = "Reverted";
        public const string ConflictFlag = "Conflict";
        public const string BranchOnlyFlag = "Branch-only";

        public static ReviewWorkbookState Project(
            LogicalHistoryResult history,
            string stateJson,
            IEnumerable<ReviewRevisionDisplayMetadata> revisions,
            IEnumerable<ReviewSubjectDisplayMetadata> subjects)
        {
            if (history == null) throw new ArgumentNullException(nameof(history));
            if (stateJson == null) throw new ArgumentNullException(nameof(stateJson));

            List<ReviewRevisionDisplayMetadata> revisionList = MaterializeRevisions(revisions);
            Dictionary<string, ReviewRevisionDisplayMetadata> revisionById = ValidateRevisions(
                history,
                revisionList);
            Dictionary<string, ReviewSubjectDisplayMetadata> displayByKey = ValidateSubjectDisplays(
                history,
                subjects);
            HashSet<string> masterLineage = BuildMasterLineage(history.MasterRevisionId, revisionById);

            var state = new ReviewWorkbookState(history.ProjectId, stateJson);
            foreach (ReviewRevisionDisplayMetadata revision in revisionList
                         .OrderBy(item => item.Order)
                         .ThenBy(item => item.RevisionId, StringComparer.Ordinal))
            {
                state.RevisionRows.Add(new ReviewRevisionProjectionRow(
                    revision.RevisionId,
                    revision.Label,
                    revision.Order,
                    revision.ParentRevisionId,
                    RevisionRole(history, revision.RevisionId, masterLineage),
                    revision.SourceBaseName,
                    revision.SourcePath,
                    revision.ByteContentId,
                    revision.ModifiedUtc.HasValue
                        ? revision.ModifiedUtc.Value.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture)
                        : "",
                    revision.Inference,
                    revision.Confidence));
            }

            var formulaTextSubjects = new HashSet<string>(
                history.Subjects
                    .Where(item => item.Aspect == CellChangeAspect.Formula)
                    .Select(item => item.LogicalSubjectKey),
                StringComparer.Ordinal);
            var details = new List<DetailCandidate>();
            foreach (LogicalSubjectHistory subject in history.Subjects
                         .OrderBy(item => item.LogicalSubjectKey, StringComparer.Ordinal)
                         .ThenBy(item => AspectRank(item.Aspect))
                         .ThenBy(item => item.Aspect)
                         .ThenBy(item => item.SubjectId, StringComparer.Ordinal))
            {
                // Formula attributes and shared/array metadata are part of the Formula channel. When exact
                // formula text also changed, retain the metadata in authoritative history without creating a
                // fourth, duplicate visible content row. A metadata-only change is still shown as Formula.
                if (subject.Aspect == CellChangeAspect.FormulaStructure &&
                    formulaTextSubjects.Contains(subject.LogicalSubjectKey))
                    continue;

                ReviewSubjectDisplayMetadata display = displayByKey[subject.LogicalSubjectKey];
                HistoryObservation before = EndpointBefore(history, subject);
                HistoryObservation after = EndpointMaster(history, subject);
                string beforeText = ObservationText(subject.Aspect, before);
                string afterText = ObservationText(subject.Aspect, after);

                state.ChangeRows.Add(new ReviewChangeProjectionRow(
                    subject.SubjectId,
                    display.SheetName,
                    display.Address,
                    AspectName(subject.Aspect),
                    beforeText,
                    afterText,
                    RepresentativeGroupId(subject),
                    Flags(subject),
                    subject.Observations.Count,
                    subject.CandidateStates.Count,
                    display.AlignmentConfidence.ToString()));

                details.Add(new DetailCandidate(
                    subject,
                    display,
                    beforeText,
                    afterText));
            }

            DetailCandidate representative = SelectRepresentativeDetail(details);
            if (representative != null)
            {
                state.DetailBefore = representative.BeforeText;
                state.DetailAfter = representative.AfterText;
            }
            return state;
        }

        private static List<ReviewRevisionDisplayMetadata> MaterializeRevisions(
            IEnumerable<ReviewRevisionDisplayMetadata> revisions)
        {
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));
            List<ReviewRevisionDisplayMetadata> result = revisions.ToList();
            if (result.Count < 2)
                throw new ArgumentException("A history projection requires at least two revisions.", nameof(revisions));
            if (result.Any(item => item == null))
                throw new ArgumentException("Revision display metadata cannot contain null entries.", nameof(revisions));
            return result;
        }

        private static Dictionary<string, ReviewRevisionDisplayMetadata> ValidateRevisions(
            LogicalHistoryResult history,
            IList<ReviewRevisionDisplayMetadata> revisions)
        {
            var byId = new Dictionary<string, ReviewRevisionDisplayMetadata>(StringComparer.Ordinal);
            var orders = new HashSet<int>();
            foreach (ReviewRevisionDisplayMetadata revision in revisions)
            {
                if (byId.ContainsKey(revision.RevisionId))
                    throw new ArgumentException(
                        "Duplicate revision display metadata: " + revision.RevisionId,
                        nameof(revisions));
                byId.Add(revision.RevisionId, revision);
                if (!orders.Add(revision.Order))
                    throw new ArgumentException("Revision display orders must be unique.", nameof(revisions));
            }

            if (!byId.ContainsKey(history.MasterRevisionId))
                throw new ArgumentException("Revision metadata does not contain the history master.", nameof(revisions));
            if (history.TrustedBaselineRevisionId.Length != 0 &&
                !byId.ContainsKey(history.TrustedBaselineRevisionId))
                throw new ArgumentException(
                    "Revision metadata does not contain the trusted baseline.",
                    nameof(revisions));

            int roots = 0;
            foreach (ReviewRevisionDisplayMetadata revision in revisions)
            {
                if (revision.ParentRevisionId.Length == 0)
                {
                    roots++;
                    continue;
                }
                if (!byId.TryGetValue(revision.ParentRevisionId, out ReviewRevisionDisplayMetadata parent))
                    throw new ArgumentException(
                        "Revision metadata refers to an unknown parent: " + revision.ParentRevisionId,
                        nameof(revisions));
                if (parent.Order >= revision.Order)
                    throw new ArgumentException(
                        "Revision metadata parents must precede their children.",
                        nameof(revisions));
            }
            if (roots != 1)
                throw new ArgumentException("Revision metadata must form one rooted tree.", nameof(revisions));

            HashSet<string> observedIds = null;
            foreach (LogicalSubjectHistory subject in history.Subjects)
            {
                var subjectIds = new HashSet<string>(
                    subject.Observations.Select(item => item.RevisionId),
                    StringComparer.Ordinal);
                if (observedIds == null)
                    observedIds = subjectIds;
                else if (!observedIds.SetEquals(subjectIds))
                    throw new InvalidOperationException(
                        "Logical history subjects do not share one revision occurrence set.");

                foreach (HistoryObservation observation in subject.Observations)
                {
                    if (!byId.TryGetValue(observation.RevisionId, out ReviewRevisionDisplayMetadata metadata) ||
                        metadata.Order != observation.RevisionOrder)
                        throw new ArgumentException(
                            "Revision display metadata does not match the logical history order.",
                            nameof(revisions));
                }
                foreach (HistoryTransition transition in subject.Transitions)
                {
                    if (!byId.TryGetValue(
                            transition.AfterRevisionId,
                            out ReviewRevisionDisplayMetadata after) ||
                        !string.Equals(
                            after.ParentRevisionId,
                            transition.BeforeRevisionId,
                            StringComparison.Ordinal))
                        throw new ArgumentException(
                            "Revision display metadata does not match a logical history edge.",
                            nameof(revisions));
                }
            }

            if (observedIds != null && !observedIds.SetEquals(byId.Keys))
                throw new ArgumentException(
                    "Revision display metadata must contain exactly the history revision occurrences.",
                    nameof(revisions));
            return byId;
        }

        private static Dictionary<string, ReviewSubjectDisplayMetadata> ValidateSubjectDisplays(
            LogicalHistoryResult history,
            IEnumerable<ReviewSubjectDisplayMetadata> subjects)
        {
            if (subjects == null) throw new ArgumentNullException(nameof(subjects));
            var byKey = new Dictionary<string, ReviewSubjectDisplayMetadata>(StringComparer.Ordinal);
            foreach (ReviewSubjectDisplayMetadata display in subjects)
            {
                if (display == null)
                    throw new ArgumentException("Subject display metadata cannot contain null entries.", nameof(subjects));
                if (byKey.ContainsKey(display.LogicalSubjectKey))
                    throw new ArgumentException(
                        "Duplicate subject display metadata: " + display.LogicalSubjectKey,
                        nameof(subjects));
                byKey.Add(display.LogicalSubjectKey, display);
            }

            var expected = new HashSet<string>(
                history.Subjects.Select(item => item.LogicalSubjectKey),
                StringComparer.Ordinal);
            if (!expected.SetEquals(byKey.Keys))
                throw new ArgumentException(
                    "Subject display metadata must contain exactly the logical history subjects.",
                    nameof(subjects));
            return byKey;
        }

        private static HashSet<string> BuildMasterLineage(
            string masterRevisionId,
            IReadOnlyDictionary<string, ReviewRevisionDisplayMetadata> revisions)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            ReviewRevisionDisplayMetadata cursor = revisions[masterRevisionId];
            while (cursor != null)
            {
                if (!result.Add(cursor.RevisionId))
                    throw new ArgumentException("Revision display metadata contains a parent cycle.", nameof(revisions));
                cursor = cursor.ParentRevisionId.Length == 0 ? null : revisions[cursor.ParentRevisionId];
            }
            return result;
        }

        private static string RevisionRole(
            LogicalHistoryResult history,
            string revisionId,
            ISet<string> masterLineage)
        {
            var roles = new List<string>();
            if (string.Equals(revisionId, history.MasterRevisionId, StringComparison.Ordinal))
                roles.Add("Master");
            if (string.Equals(revisionId, history.TrustedBaselineRevisionId, StringComparison.Ordinal))
                roles.Add("Trusted baseline");
            if (!masterLineage.Contains(revisionId)) roles.Add("Branch");
            return roles.Count == 0 ? "Revision" : string.Join("; ", roles);
        }

        private static HistoryObservation EndpointBefore(
            LogicalHistoryResult history,
            LogicalSubjectHistory subject)
        {
            if (history.TrustedBaselineRevisionId.Length != 0)
                return subject.Observations.Single(item => string.Equals(
                    item.RevisionId,
                    history.TrustedBaselineRevisionId,
                    StringComparison.Ordinal));
            return subject.Observations
                .OrderBy(item => item.RevisionOrder)
                .ThenBy(item => item.RevisionId, StringComparer.Ordinal)
                .First();
        }

        private static HistoryObservation EndpointMaster(
            LogicalHistoryResult history,
            LogicalSubjectHistory subject) =>
            subject.Observations.Single(item => string.Equals(
                item.RevisionId,
                history.MasterRevisionId,
                StringComparison.Ordinal));

        private static string ObservationText(CellChangeAspect aspect, HistoryObservation observation)
        {
            switch (aspect)
            {
                case CellChangeAspect.Value:
                case CellChangeAspect.CachedResult:
                    return observation.ExactValue == null ? "" : observation.ExactValue.ToString();
                default:
                    return observation.ExactText ?? "";
            }
        }

        private static string AspectName(CellChangeAspect aspect)
        {
            switch (aspect)
            {
                case CellChangeAspect.Formula: return "Formula";
                case CellChangeAspect.FormulaStructure: return "Formula";
                case CellChangeAspect.Value: return "Value";
                case CellChangeAspect.CachedResult: return "Calculated value";
                default: return "Cell presence";
            }
        }

        private static int AspectRank(CellChangeAspect aspect)
        {
            switch (aspect)
            {
                case CellChangeAspect.Value: return 0;
                case CellChangeAspect.CachedResult: return 1;
                case CellChangeAspect.Formula: return 2;
                case CellChangeAspect.FormulaStructure: return 3;
                default: return 4;
            }
        }

        private static string RepresentativeGroupId(LogicalSubjectHistory subject)
        {
            string[] groupIds = subject.Transitions
                .Select(item => item.GroupId)
                .Where(item => !string.IsNullOrEmpty(item))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(item => item, StringComparer.Ordinal)
                .Take(2)
                .ToArray();
            return groupIds.Length == 1 ? groupIds[0] : "";
        }

        private static string Flags(LogicalSubjectHistory subject)
        {
            var flags = new List<string>();
            if (subject.Transitions.Any(item => item.EquivalenceTags != EquivalenceTag.None))
                flags.Add(EquivalentFlag);
            if (subject.IsReverted) flags.Add(RevertedFlag);
            if (subject.HasConflict) flags.Add(ConflictFlag);
            if (subject.HasBranchOnlyChange) flags.Add(BranchOnlyFlag);
            return string.Join("; ", flags);
        }

        private static DetailCandidate SelectRepresentativeDetail(IEnumerable<DetailCandidate> candidates)
        {
            List<DetailCandidate> ordered = candidates
                .Where(item => item.BeforeText.Length != 0 || item.AfterText.Length != 0)
                .OrderBy(item => string.Equals(item.BeforeText, item.AfterText, StringComparison.Ordinal) ? 1 : 0)
                .ThenBy(item => DetailRank(item.Subject.Aspect))
                .ThenBy(item => item.Display.Address.Length == 0 ? 1 : 0)
                .ThenBy(item => item.Subject.LogicalSubjectKey, StringComparer.Ordinal)
                .ThenBy(item => item.Subject.SubjectId, StringComparer.Ordinal)
                .ToList();
            return ordered.FirstOrDefault();
        }

        private static int DetailRank(CellChangeAspect aspect)
        {
            switch (aspect)
            {
                case CellChangeAspect.Value: return 0;
                case CellChangeAspect.CachedResult: return 1;
                case CellChangeAspect.Formula: return 2;
                case CellChangeAspect.FormulaStructure: return 3;
                default: return 4;
            }
        }

        private sealed class DetailCandidate
        {
            internal DetailCandidate(
                LogicalSubjectHistory subject,
                ReviewSubjectDisplayMetadata display,
                string beforeText,
                string afterText)
            {
                Subject = subject;
                Display = display;
                BeforeText = beforeText;
                AfterText = afterText;
            }

            internal LogicalSubjectHistory Subject { get; }
            internal ReviewSubjectDisplayMetadata Display { get; }
            internal string BeforeText { get; }
            internal string AfterText { get; }
        }
    }
}
