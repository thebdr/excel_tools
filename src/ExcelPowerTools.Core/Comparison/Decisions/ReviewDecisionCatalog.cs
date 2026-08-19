using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>
    /// Immutable membership of one review group. Subject identities are stored in ordinal order so the same
    /// logical group has the same durable representation regardless of presentation order.
    /// </summary>
    public sealed class ReviewDecisionGroup
    {
        public ReviewDecisionGroup(string groupId, IEnumerable<string> subjectIds)
        {
            GroupId = Require(groupId, nameof(groupId));
            SubjectIds = new ReadOnlyCollection<string>(
                FreezeIdentities(subjectIds, nameof(subjectIds), requireAtLeastOne: true));
        }

        public string GroupId { get; }
        public IReadOnlyList<string> SubjectIds { get; }

        private static List<string> FreezeIdentities(
            IEnumerable<string> values,
            string parameterName,
            bool requireAtLeastOne)
        {
            if (values == null) throw new ArgumentNullException(parameterName);
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string value in values)
            {
                string required = Require(value, parameterName);
                if (!seen.Add(required))
                    throw new ArgumentException("A review group cannot repeat a subject ID.", parameterName);
                result.Add(required);
            }

            if (requireAtLeastOne && result.Count == 0)
                throw new ArgumentException("A review group must contain at least one subject.", parameterName);
            result.Sort(StringComparer.Ordinal);
            return result;
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable identity is required.", name);
            return value;
        }
    }

    /// <summary>
    /// Exact immutable evidence universe against which reviewer decisions are validated. It intentionally contains
    /// identities and observations only; presentation rows, filters, and workbook state are not review authority.
    /// </summary>
    public sealed class ReviewDecisionCatalog
    {
        private readonly HashSet<string> _subjectIds;
        private readonly IReadOnlyDictionary<string, ReviewDecisionSubject> _subjectsById;
        private readonly IReadOnlyDictionary<string, Observation> _observationsById;
        private readonly IReadOnlyDictionary<string, IReadOnlyDictionary<string, Observation>>
            _observationsBySubjectAndRevision;
        private readonly IReadOnlyDictionary<string, ReviewDecisionGroup> _groupsById;

        /// <summary>
        /// Backward-compatible subject-ID-only constructor. Its subjects hydrate as Unspecified/Unknown and are
        /// safe for non-selection review statuses only; durable new projects should use explicit subject metadata.
        /// </summary>
        public ReviewDecisionCatalog(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string masterRevisionId,
            IEnumerable<string> subjectIds,
            IEnumerable<Observation> observations,
            IEnumerable<ReviewDecisionGroup> groups = null)
            : this(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                masterRevisionId,
                LegacySubjects(subjectIds),
                observations,
                groups,
                null,
                true)
        {
        }

        public ReviewDecisionCatalog(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string masterRevisionId,
            IEnumerable<ReviewDecisionSubject> subjects,
            IEnumerable<Observation> observations,
            IEnumerable<ReviewDecisionGroup> groups = null)
            : this(
                projectId,
                baseGenerationId,
                mappingGenerationId,
                masterRevisionId,
                subjects,
                observations,
                groups,
                null,
                true)
        {
        }

        public ReviewDecisionCatalog(
            ReviewEvidenceGeneration evidence,
            string masterRevisionId,
            IEnumerable<ReviewDecisionSubject> subjects,
            IEnumerable<Observation> observations,
            IEnumerable<ReviewDecisionGroup> groups = null)
            : this(
                RequireEvidence(evidence).ProjectId,
                evidence.GenerationId,
                evidence.EffectiveMappingGenerationId,
                RequireMasterRevision(evidence, masterRevisionId),
                subjects,
                observations,
                groups,
                evidence,
                true)
        {
        }

        private ReviewDecisionCatalog(
            string projectId,
            string baseGenerationId,
            string mappingGenerationId,
            string masterRevisionId,
            IEnumerable<ReviewDecisionSubject> subjects,
            IEnumerable<Observation> observations,
            IEnumerable<ReviewDecisionGroup> groups,
            ReviewEvidenceGeneration evidenceGeneration,
            bool validatedOverload)
        {
            ProjectId = Require(projectId, nameof(projectId));
            BaseGenerationId = Require(baseGenerationId, nameof(baseGenerationId));
            MappingGenerationId = Require(mappingGenerationId, nameof(mappingGenerationId));
            MasterRevisionId = Require(masterRevisionId, nameof(masterRevisionId));
            EvidenceGeneration = evidenceGeneration;
            if (EvidenceGeneration != null)
            {
                if (!string.Equals(EvidenceGeneration.ProjectId, ProjectId, StringComparison.Ordinal) ||
                    !string.Equals(EvidenceGeneration.GenerationId, BaseGenerationId, StringComparison.Ordinal) ||
                    !string.Equals(
                        EvidenceGeneration.EffectiveMappingGenerationId,
                        MappingGenerationId,
                        StringComparison.Ordinal) ||
                    !EvidenceGeneration.Revisions.Any(item =>
                        string.Equals(item.RevisionId, MasterRevisionId, StringComparison.Ordinal)))
                    throw new InvalidOperationException(
                        "The retained evidence generation does not match the catalog bindings.");
            }

            List<ReviewDecisionSubject> frozenSubjects = FreezeSubjects(subjects);
            Subjects = new ReadOnlyCollection<ReviewDecisionSubject>(frozenSubjects);
            _subjectsById = new ReadOnlyDictionary<string, ReviewDecisionSubject>(
                frozenSubjects.ToDictionary(item => item.SubjectId, StringComparer.Ordinal));
            List<string> subjectIds = frozenSubjects.Select(item => item.SubjectId).ToList();
            SubjectIds = new ReadOnlyCollection<string>(subjectIds);
            _subjectIds = new HashSet<string>(subjectIds, StringComparer.Ordinal);

            ISet<string> evidenceRevisionIds = EvidenceGeneration == null
                ? null
                : new HashSet<string>(
                    EvidenceGeneration.Revisions.Select(item => item.RevisionId),
                    StringComparer.Ordinal);
            List<Observation> frozenObservations = FreezeObservations(
                observations,
                _subjectIds,
                evidenceRevisionIds);
            Observations = new ReadOnlyCollection<Observation>(frozenObservations);
            _observationsById = new ReadOnlyDictionary<string, Observation>(
                frozenObservations.ToDictionary(item => item.ObservationId, StringComparer.Ordinal));
            var observationsBySubjectAndRevision =
                new Dictionary<string, IReadOnlyDictionary<string, Observation>>(StringComparer.Ordinal);
            foreach (IGrouping<string, Observation> subjectObservations in frozenObservations.GroupBy(
                         item => item.SubjectId,
                         StringComparer.Ordinal))
            {
                observationsBySubjectAndRevision.Add(
                    subjectObservations.Key,
                    new ReadOnlyDictionary<string, Observation>(
                        subjectObservations.ToDictionary(
                            item => item.RevisionId,
                            StringComparer.Ordinal)));
            }
            _observationsBySubjectAndRevision =
                new ReadOnlyDictionary<string, IReadOnlyDictionary<string, Observation>>(
                    observationsBySubjectAndRevision);

            List<ReviewDecisionGroup> frozenGroups = FreezeGroups(groups, _subjectIds);
            Groups = new ReadOnlyCollection<ReviewDecisionGroup>(frozenGroups);
            _groupsById = new ReadOnlyDictionary<string, ReviewDecisionGroup>(
                frozenGroups.ToDictionary(item => item.GroupId, StringComparer.Ordinal));

            CatalogId = ComputeCatalogId();
        }

        public string CatalogId { get; }
        public string ProjectId { get; }
        /// <summary>The immutable evidence/history generation reviewed by this catalog.</summary>
        public string BaseGenerationId { get; }
        public string EvidenceGenerationId => BaseGenerationId;
        /// <summary>
        /// Exact ordered revision-byte evidence for new durable catalogs. Null only for legacy/manual catalogs
        /// whose callers supplied an opaque base generation ID instead of a verifiable evidence generation.
        /// </summary>
        public ReviewEvidenceGeneration EvidenceGeneration { get; }
        public string MappingGenerationId { get; }
        public string MasterRevisionId { get; }
        public IReadOnlyList<ReviewDecisionSubject> Subjects { get; }
        public IReadOnlyList<string> SubjectIds { get; }
        public IReadOnlyList<Observation> Observations { get; }
        public IReadOnlyList<ReviewDecisionGroup> Groups { get; }

        public bool ContainsSubject(string subjectId) =>
            subjectId != null && _subjectIds.Contains(subjectId);

        public bool TryGetSubject(string subjectId, out ReviewDecisionSubject subject)
        {
            if (subjectId == null)
            {
                subject = null;
                return false;
            }
            return _subjectsById.TryGetValue(subjectId, out subject);
        }

        public bool TryGetObservation(string observationId, out Observation observation)
        {
            if (observationId == null)
            {
                observation = null;
                return false;
            }
            return _observationsById.TryGetValue(observationId, out observation);
        }

        /// <summary>O(1) lookup of the unique observation for one exact subject/revision occurrence.</summary>
        public bool TryGetObservation(
            string subjectId,
            string revisionId,
            out Observation observation)
        {
            observation = null;
            return subjectId != null &&
                   revisionId != null &&
                   _observationsBySubjectAndRevision.TryGetValue(
                       subjectId,
                       out IReadOnlyDictionary<string, Observation> byRevision) &&
                   byRevision.TryGetValue(revisionId, out observation);
        }

        public bool TryGetGroup(string groupId, out ReviewDecisionGroup group)
        {
            if (groupId == null)
            {
                group = null;
                return false;
            }
            return _groupsById.TryGetValue(groupId, out group);
        }

        private string ComputeCatalogId()
        {
            var components = new List<string>
            {
                ProjectId,
                BaseGenerationId,
                MappingGenerationId,
                MasterRevisionId,
                Subjects.Count.ToString(CultureInfo.InvariantCulture),
            };
            foreach (ReviewDecisionSubject subject in Subjects)
            {
                components.Add("subject-id:" + subject.SubjectId);
                components.Add("subject-key:" + subject.LogicalSubjectKey);
                components.Add("subject-aspect:" +
                               ((int)subject.Aspect).ToString(CultureInfo.InvariantCulture));
                components.Add("subject-policy:" +
                               ((int)subject.Policy).ToString(CultureInfo.InvariantCulture));
            }
            components.Add(Observations.Count.ToString(CultureInfo.InvariantCulture));
            foreach (Observation observation in Observations)
            {
                components.Add("observation-id:" + observation.ObservationId);
                components.Add("observation-subject:" + observation.SubjectId);
                components.Add("observation-revision:" + observation.RevisionId);
                components.Add("observation-type:" +
                               ((int)observation.ExactType).ToString(CultureInfo.InvariantCulture));
                components.Add("observation-content:" + observation.ContentHash);
                components.Add("observation-formula-present:" +
                               (observation.ExactType == ObservationExactType.Formula ? "1" : "0"));
                components.Add("observation-formula:" + observation.Formula);
                AddScalarComponents(components, "observation-value", observation.Value);
                AddScalarComponents(
                    components,
                    "observation-cached-result",
                    observation.CachedResult);
                components.Add("observation-cache-trust:" +
                               ((int)observation.CachedResultTrust).ToString(CultureInfo.InvariantCulture));
                components.Add("observation-display-text:" + observation.DisplayText);
                components.Add("observation-logical-sheet:" + observation.Provenance.LogicalSheetId);
                components.Add("observation-source:" + observation.Provenance.SourceLabel);
                components.Add("observation-sheet:" + observation.Provenance.SheetName);
                components.Add("observation-address:" + observation.Provenance.Address);
                components.Add("observation-row:" +
                               observation.Provenance.PhysicalRow.ToString(CultureInfo.InvariantCulture));
                components.Add("observation-column:" +
                               observation.Provenance.PhysicalColumn.ToString(CultureInfo.InvariantCulture));
            }
            components.Add(Groups.Count.ToString(CultureInfo.InvariantCulture));
            foreach (ReviewDecisionGroup group in Groups)
            {
                components.Add("group-id:" + group.GroupId);
                components.Add("group-size:" + group.SubjectIds.Count.ToString(CultureInfo.InvariantCulture));
                components.AddRange(group.SubjectIds.Select(item => "group-subject:" + item));
            }
            return StableId.Compute("review-decision-catalog-v2", components.ToArray());
        }

        private static IEnumerable<ReviewDecisionSubject> LegacySubjects(IEnumerable<string> subjectIds)
        {
            if (subjectIds == null) throw new ArgumentNullException(nameof(subjectIds));
            return subjectIds.Select(subjectId => new ReviewDecisionSubject(subjectId));
        }

        private static List<ReviewDecisionSubject> FreezeSubjects(
            IEnumerable<ReviewDecisionSubject> subjects)
        {
            if (subjects == null) throw new ArgumentNullException(nameof(subjects));
            var result = new List<ReviewDecisionSubject>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var semanticIdentities = new HashSet<string>(StringComparer.Ordinal);
            foreach (ReviewDecisionSubject subject in subjects)
            {
                if (subject == null)
                    throw new ArgumentException(
                        "The decision catalog cannot contain a null subject.",
                        nameof(subjects));
                if (!seen.Add(subject.SubjectId))
                    throw new ArgumentException(
                        "The decision catalog cannot repeat a subject ID.",
                        nameof(subjects));
                string semanticIdentity = StableId.Compute(
                    "review-decision-subject-semantics-v1",
                    subject.LogicalSubjectKey,
                    ((int)subject.Aspect).ToString(CultureInfo.InvariantCulture));
                if (!semanticIdentities.Add(semanticIdentity))
                    throw new ArgumentException(
                        "The decision catalog cannot assign two IDs to the same logical subject aspect.",
                        nameof(subjects));
                result.Add(subject);
            }
            result.Sort((left, right) =>
                StringComparer.Ordinal.Compare(left.SubjectId, right.SubjectId));
            return result;
        }

        private static List<Observation> FreezeObservations(
            IEnumerable<Observation> observations,
            ISet<string> knownSubjectIds,
            ISet<string> evidenceRevisionIds)
        {
            if (observations == null) throw new ArgumentNullException(nameof(observations));
            var result = new List<Observation>();
            var observationIds = new HashSet<string>(StringComparer.Ordinal);
            var subjectRevisionPairs = new HashSet<string>(StringComparer.Ordinal);
            foreach (Observation observation in observations)
            {
                if (observation == null)
                    throw new ArgumentException("A decision catalog cannot contain a null observation.", nameof(observations));
                if (!knownSubjectIds.Contains(observation.SubjectId))
                    throw new ArgumentException(
                        "Every observation must reference a subject in the decision catalog.",
                        nameof(observations));
                if (evidenceRevisionIds != null &&
                    !evidenceRevisionIds.Contains(observation.RevisionId))
                    throw new ArgumentException(
                        "Every evidence-bound observation must reference a revision in the exact evidence generation.",
                        nameof(observations));
                if (evidenceRevisionIds != null)
                {
                    if (!StableId.IsSha256(observation.ContentHash))
                        throw new ArgumentException(
                            "Every evidence-bound observation content identity must be a SHA-256 value.",
                            nameof(observations));
                    string expectedObservationId = ComparisonIds.Observation(
                        observation.SubjectId,
                        observation.RevisionId,
                        observation.ExactType,
                        observation.ContentHash);
                    if (!string.Equals(
                            observation.ObservationId,
                            expectedObservationId,
                            StringComparison.Ordinal))
                        throw new ArgumentException(
                            "An evidence-bound observation ID disagrees with its exact subject, revision, type, and content identity.",
                            nameof(observations));
                }
                if (!observationIds.Add(observation.ObservationId))
                    throw new ArgumentException("The decision catalog cannot repeat an observation ID.", nameof(observations));

                // One logical subject has at most one exact observation in a revision. Accepting two would make a
                // revision-level selection intrinsically ambiguous even when their observation IDs differ.
                string subjectRevision = StableId.Compute(
                    "review-decision-subject-revision-v1",
                    observation.SubjectId,
                    observation.RevisionId);
                if (!subjectRevisionPairs.Add(subjectRevision))
                    throw new ArgumentException(
                        "A subject cannot have multiple observations in the same revision.",
                        nameof(observations));
                result.Add(observation);
            }
            if (evidenceRevisionIds != null)
            {
                long expectedCount = (long)knownSubjectIds.Count * evidenceRevisionIds.Count;
                if (result.Count != expectedCount)
                    throw new ArgumentException(
                        "An evidence-bound catalog requires exactly one observation for every subject and evidence revision.",
                        nameof(observations));
            }
            result.Sort((left, right) => StringComparer.Ordinal.Compare(left.ObservationId, right.ObservationId));
            return result;
        }

        private static List<ReviewDecisionGroup> FreezeGroups(
            IEnumerable<ReviewDecisionGroup> groups,
            ISet<string> knownSubjectIds)
        {
            var result = new List<ReviewDecisionGroup>();
            var groupIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (ReviewDecisionGroup group in groups ?? Array.Empty<ReviewDecisionGroup>())
            {
                if (group == null)
                    throw new ArgumentException("A decision catalog cannot contain a null group.", nameof(groups));
                if (!groupIds.Add(group.GroupId))
                    throw new ArgumentException("The decision catalog cannot repeat a group ID.", nameof(groups));
                if (group.SubjectIds.Any(subjectId => !knownSubjectIds.Contains(subjectId)))
                    throw new ArgumentException(
                        "Every review-group member must reference a subject in the decision catalog.",
                        nameof(groups));
                result.Add(group);
            }
            result.Sort((left, right) => StringComparer.Ordinal.Compare(left.GroupId, right.GroupId));
            return result;
        }

        private static void AddScalarComponents(
            ICollection<string> components,
            string prefix,
            CanonicalScalar scalar)
        {
            components.Add(prefix + "-present:" + (scalar == null ? "0" : "1"));
            if (scalar == null) return;
            components.Add(prefix + "-type:" +
                           ((int)scalar.Type).ToString(CultureInfo.InvariantCulture));
            components.Add(prefix + "-text:" + scalar.CanonicalText);
            components.Add(prefix + "-date-system:" +
                           ((int)scalar.DateSystem).ToString(CultureInfo.InvariantCulture));
        }

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable identity is required.", name);
            return value;
        }

        private static ReviewEvidenceGeneration RequireEvidence(ReviewEvidenceGeneration evidence) =>
            evidence ?? throw new ArgumentNullException(nameof(evidence));

        private static string RequireMasterRevision(
            ReviewEvidenceGeneration evidence,
            string masterRevisionId)
        {
            RequireEvidence(evidence);
            string required = Require(masterRevisionId, nameof(masterRevisionId));
            if (!evidence.Revisions.Any(item =>
                    string.Equals(item.RevisionId, required, StringComparison.Ordinal)))
                throw new ArgumentException(
                    "The master revision must belong to the exact evidence generation.",
                    nameof(masterRevisionId));
            return required;
        }
    }
}
