using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Diff;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.History
{
    /// <summary>
    /// One immutable revision occurrence supplied to the history engine. Snapshots must already use logical,
    /// aligned subject keys; physical addresses are evidence only and are deliberately outside this contract.
    /// </summary>
    public sealed class HistoryRevision
    {
        public HistoryRevision(
            string revisionId,
            string parentRevisionId,
            int order,
            RevisionSnapshot snapshot,
            bool isMaster = false,
            bool isTrustedBaseline = false)
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (order < 0) throw new ArgumentOutOfRangeException(nameof(order));
            if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
            if (!string.Equals(revisionId, snapshot.RevisionId, StringComparison.Ordinal))
                throw new ArgumentException("The snapshot revision ID must match its history revision.", nameof(snapshot));

            RevisionId = revisionId;
            ParentRevisionId = parentRevisionId ?? "";
            Order = order;
            Snapshot = snapshot;
            IsMaster = isMaster;
            IsTrustedBaseline = isTrustedBaseline;
        }

        public string RevisionId { get; }
        public string ParentRevisionId { get; }
        public int Order { get; }
        public RevisionSnapshot Snapshot { get; }
        public bool IsMaster { get; }
        public bool IsTrustedBaseline { get; }
    }

    /// <summary>Exact state of one independently reviewable cell aspect in one revision.</summary>
    public sealed class HistoryObservation
    {
        internal HistoryObservation(
            string revisionId,
            int revisionOrder,
            string stateId,
            bool isPhysicallyPresent,
            string exactText,
            CanonicalValue exactValue)
        {
            RevisionId = revisionId;
            RevisionOrder = revisionOrder;
            StateId = stateId;
            IsPhysicallyPresent = isPhysicallyPresent;
            ExactText = exactText;
            ExactValue = exactValue;
        }

        public string RevisionId { get; }
        public int RevisionOrder { get; }
        public string StateId { get; }
        public bool IsPhysicallyPresent { get; }
        public string ExactText { get; }
        public CanonicalValue ExactValue { get; }
    }

    /// <summary>One distinct exact state and every revision occurrence that supports it.</summary>
    public sealed class HistoryCandidateState
    {
        internal HistoryCandidateState(string stateId, IReadOnlyList<string> revisionIds)
        {
            StateId = stateId;
            RevisionIds = revisionIds;
        }

        public string StateId { get; }
        public IReadOnlyList<string> RevisionIds { get; }
    }

    /// <summary>Exact parent-to-child evidence for one logical subject/aspect.</summary>
    public sealed class HistoryTransition
    {
        internal HistoryTransition(
            string transitionId,
            string edgeId,
            string beforeRevisionId,
            string afterRevisionId,
            string beforeStateId,
            string afterStateId,
            CellChangeKind exactKind,
            EquivalenceTag equivalenceTags,
            bool isOnMasterLineage)
        {
            TransitionId = transitionId;
            EdgeId = edgeId;
            BeforeRevisionId = beforeRevisionId;
            AfterRevisionId = afterRevisionId;
            BeforeStateId = beforeStateId;
            AfterStateId = afterStateId;
            ExactKind = exactKind;
            EquivalenceTags = equivalenceTags;
            IsOnMasterLineage = isOnMasterLineage;
        }

        public string TransitionId { get; }
        public string EdgeId { get; }
        public string BeforeRevisionId { get; }
        public string AfterRevisionId { get; }
        public string BeforeStateId { get; }
        public string AfterStateId { get; }
        public CellChangeKind ExactKind { get; }
        public EquivalenceTag EquivalenceTags { get; }
        public bool IsOnMasterLineage { get; }
        public bool IsRevert { get; internal set; }
        public bool IsConflict { get; internal set; }
        public bool IsBranchOnly { get; internal set; }
        public string GroupId { get; internal set; } = "";
    }

    public sealed class LogicalSubjectHistory
    {
        internal LogicalSubjectHistory(
            string subjectId,
            string logicalSubjectKey,
            CellChangeAspect aspect,
            IReadOnlyList<HistoryObservation> observations,
            IReadOnlyList<HistoryCandidateState> candidateStates,
            IReadOnlyList<HistoryTransition> transitions)
        {
            SubjectId = subjectId;
            LogicalSubjectKey = logicalSubjectKey;
            Aspect = aspect;
            Observations = observations;
            CandidateStates = candidateStates;
            Transitions = transitions;
        }

        public string SubjectId { get; }
        public string LogicalSubjectKey { get; }
        public CellChangeAspect Aspect { get; }
        public IReadOnlyList<HistoryObservation> Observations { get; }
        public IReadOnlyList<HistoryCandidateState> CandidateStates { get; }
        public IReadOnlyList<HistoryTransition> Transitions { get; }
        public bool IsReverted { get; internal set; }
        public bool HasConflict { get; internal set; }
        public bool HasBranchOnlyChange { get; internal set; }
    }

    public sealed class HistoryChangeGroup
    {
        internal HistoryChangeGroup(
            string groupId,
            string edgeId,
            CellChangeAspect aspect,
            string beforeStateId,
            string afterStateId,
            IReadOnlyList<string> subjectIds)
        {
            GroupId = groupId;
            EdgeId = edgeId;
            Aspect = aspect;
            BeforeStateId = beforeStateId;
            AfterStateId = afterStateId;
            SubjectIds = subjectIds;
        }

        public string GroupId { get; }
        public string EdgeId { get; }
        public CellChangeAspect Aspect { get; }
        public string BeforeStateId { get; }
        public string AfterStateId { get; }
        public IReadOnlyList<string> SubjectIds { get; }
    }

    public sealed class LogicalHistoryResult
    {
        internal LogicalHistoryResult(
            string projectId,
            string mappingGenerationId,
            string masterRevisionId,
            string trustedBaselineRevisionId,
            IReadOnlyList<LogicalSubjectHistory> subjects,
            IReadOnlyList<HistoryChangeGroup> groups)
        {
            ProjectId = projectId;
            MappingGenerationId = mappingGenerationId;
            MasterRevisionId = masterRevisionId;
            TrustedBaselineRevisionId = trustedBaselineRevisionId;
            Subjects = subjects;
            Groups = groups;
        }

        public string ProjectId { get; }
        public string MappingGenerationId { get; }
        public string MasterRevisionId { get; }
        public string TrustedBaselineRevisionId { get; }
        public IReadOnlyList<LogicalSubjectHistory> Subjects { get; }
        public IReadOnlyList<HistoryChangeGroup> Groups { get; }
        public int AtomicTransitionCount => Subjects.Sum(subject => subject.Transitions.Count);
    }

    /// <summary>
    /// Builds adjacent-edge histories and repeated-transition groups without altering atomic evidence. The
    /// engine deliberately assumes that sheet/row/column identity has already been confirmed or conservatively
    /// aligned by an earlier stage.
    /// </summary>
    public static class LogicalHistoryEngine
    {
        public static LogicalHistoryResult Build(
            string projectId,
            string mappingGenerationId,
            IEnumerable<HistoryRevision> revisions,
            ExactComparisonOptions options = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            Require(projectId, nameof(projectId));
            Require(mappingGenerationId, nameof(mappingGenerationId));
            List<HistoryRevision> ordered = ValidateAndOrder(revisions);
            options = options ?? new ExactComparisonOptions();

            HistoryRevision master = ordered.Single(revision => revision.IsMaster);
            HistoryRevision baseline = ordered.SingleOrDefault(revision => revision.IsTrustedBaseline);
            var byId = ordered.ToDictionary(revision => revision.RevisionId, StringComparer.Ordinal);
            HashSet<string> masterLineage = BuildLineage(master, byId);
            var changes = new List<EdgeChange>();

            foreach (HistoryRevision after in ordered)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (after.ParentRevisionId.Length == 0) continue;
                HistoryRevision before = byId[after.ParentRevisionId];
                string edgeId = ComparisonIds.Edge(projectId, before.RevisionId, after.RevisionId);
                IReadOnlyList<AtomicCellChange> edgeChanges = ExactComparisonEngine.Compare(
                    before.Snapshot,
                    after.Snapshot,
                    options,
                    cancellationToken);

                foreach (AtomicCellChange change in edgeChanges)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    AspectState beforeState = Extract(before.Snapshot, change.SubjectKey, change.Aspect);
                    AspectState afterState = Extract(after.Snapshot, change.SubjectKey, change.Aspect);
                    changes.Add(new EdgeChange(
                        edgeId,
                        before,
                        after,
                        change,
                        beforeState,
                        afterState,
                        masterLineage.Contains(after.RevisionId)));
                }
            }

            List<LogicalSubjectHistory> subjects = BuildSubjects(
                projectId,
                ordered,
                changes,
                byId,
                masterLineage,
                baseline,
                master,
                cancellationToken);
            List<HistoryChangeGroup> groups = BuildGroups(
                mappingGenerationId,
                subjects,
                cancellationToken);

            return new LogicalHistoryResult(
                projectId,
                mappingGenerationId,
                master.RevisionId,
                baseline == null ? "" : baseline.RevisionId,
                new ReadOnlyCollection<LogicalSubjectHistory>(subjects),
                new ReadOnlyCollection<HistoryChangeGroup>(groups));
        }

        private static List<LogicalSubjectHistory> BuildSubjects(
            string projectId,
            IReadOnlyList<HistoryRevision> ordered,
            IReadOnlyList<EdgeChange> changes,
            IReadOnlyDictionary<string, HistoryRevision> byId,
            ISet<string> masterLineage,
            HistoryRevision baseline,
            HistoryRevision master,
            CancellationToken cancellationToken)
        {
            var changesBySubject = new Dictionary<SubjectAspect, List<EdgeChange>>();
            for (int index = 0; index < changes.Count; index++)
            {
                if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                EdgeChange change = changes[index];
                var key = new SubjectAspect(change.Change.SubjectKey, change.Change.Aspect);
                if (!changesBySubject.TryGetValue(key, out List<EdgeChange> subjectChanges))
                {
                    subjectChanges = new List<EdgeChange>();
                    changesBySubject.Add(key, subjectChanges);
                }
                subjectChanges.Add(change);
            }

            var keys = changesBySubject.Keys
                .OrderBy(key => key.SubjectKey, StringComparer.Ordinal)
                .ThenBy(key => key.Aspect)
                .ToList();
            var results = new List<LogicalSubjectHistory>(keys.Count);

            foreach (SubjectAspect key in keys)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string subjectId = StableId.Compute(
                    "history-subject-v1",
                    projectId,
                    key.SubjectKey,
                    ((int)key.Aspect).ToString(CultureInfo.InvariantCulture));
                var observations = new List<HistoryObservation>(ordered.Count);
                var observationByRevision = new Dictionary<string, HistoryObservation>(StringComparer.Ordinal);

                foreach (HistoryRevision revision in ordered)
                {
                    AspectState state = Extract(revision.Snapshot, key.SubjectKey, key.Aspect);
                    var observation = new HistoryObservation(
                        revision.RevisionId,
                        revision.Order,
                        state.StateId,
                        state.IsPresent,
                        state.Text,
                        state.Value);
                    observations.Add(observation);
                    observationByRevision.Add(revision.RevisionId, observation);
                }

                var transitions = new List<HistoryTransition>();
                foreach (EdgeChange edgeChange in changesBySubject[key])
                {
                    string transitionId = StableId.Compute(
                        "history-transition-v1",
                        edgeChange.EdgeId,
                        subjectId,
                        edgeChange.BeforeState.StateId,
                        edgeChange.AfterState.StateId);
                    transitions.Add(new HistoryTransition(
                        transitionId,
                        edgeChange.EdgeId,
                        edgeChange.Before.RevisionId,
                        edgeChange.After.RevisionId,
                        edgeChange.BeforeState.StateId,
                        edgeChange.AfterState.StateId,
                        edgeChange.Change.Kind,
                        EquivalenceClassifier.Classify(edgeChange.Change),
                        edgeChange.IsOnMasterLineage));
                }

                MarkReverts(transitions, observationByRevision, byId);
                bool hasConflict = MarkConflicts(transitions, observationByRevision, ordered);
                bool hasBranchOnly = MarkBranchOnly(transitions, masterLineage);
                bool isReverted = IsSubjectReverted(
                    observations,
                    masterLineage,
                    baseline,
                    master);

                var candidateStates = observations
                    .GroupBy(observation => observation.StateId, StringComparer.Ordinal)
                    .OrderBy(group => group.Min(observation => observation.RevisionOrder))
                    .ThenBy(group => group.Key, StringComparer.Ordinal)
                    .Select(group => new HistoryCandidateState(
                        group.Key,
                        new ReadOnlyCollection<string>(group
                            .OrderBy(observation => observation.RevisionOrder)
                            .Select(observation => observation.RevisionId)
                            .ToList())))
                    .ToList();

                var history = new LogicalSubjectHistory(
                    subjectId,
                    key.SubjectKey,
                    key.Aspect,
                    new ReadOnlyCollection<HistoryObservation>(observations),
                    new ReadOnlyCollection<HistoryCandidateState>(candidateStates),
                    new ReadOnlyCollection<HistoryTransition>(transitions))
                {
                    IsReverted = isReverted,
                    HasConflict = hasConflict,
                    HasBranchOnlyChange = hasBranchOnly,
                };
                results.Add(history);
            }

            return results;
        }

        private static List<HistoryChangeGroup> BuildGroups(
            string generationId,
            IReadOnlyList<LogicalSubjectHistory> subjects,
            CancellationToken cancellationToken)
        {
            var memberships = new Dictionary<string, List<Tuple<LogicalSubjectHistory, HistoryTransition>>>(
                StringComparer.Ordinal);
            foreach (LogicalSubjectHistory subject in subjects)
            {
                foreach (HistoryTransition transition in subject.Transitions)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string key = string.Join("|", new[]
                    {
                        transition.EdgeId,
                        ((int)subject.Aspect).ToString(CultureInfo.InvariantCulture),
                        transition.BeforeStateId,
                        transition.AfterStateId,
                    });
                    if (!memberships.TryGetValue(key, out List<Tuple<LogicalSubjectHistory, HistoryTransition>> list))
                    {
                        list = new List<Tuple<LogicalSubjectHistory, HistoryTransition>>();
                        memberships.Add(key, list);
                    }
                    list.Add(Tuple.Create(subject, transition));
                }
            }

            var groups = new List<HistoryChangeGroup>();
            foreach (List<Tuple<LogicalSubjectHistory, HistoryTransition>> members in memberships.Values
                         .Where(list => list.Count > 1)
                         .OrderBy(list => list[0].Item2.EdgeId, StringComparer.Ordinal)
                         .ThenBy(list => list[0].Item1.Aspect)
                         .ThenBy(list => list[0].Item2.BeforeStateId, StringComparer.Ordinal)
                         .ThenBy(list => list[0].Item2.AfterStateId, StringComparer.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string[] subjectIds = members
                    .Select(member => member.Item1.SubjectId)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(id => id, StringComparer.Ordinal)
                    .ToArray();
                if (subjectIds.Length < 2) continue;
                HistoryTransition firstTransition = members[0].Item2;
                LogicalSubjectHistory firstSubject = members[0].Item1;
                string groupId = ComparisonIds.Group(
                    generationId,
                    "identical-exact-transition-v1",
                    subjectIds);
                foreach (Tuple<LogicalSubjectHistory, HistoryTransition> member in members)
                    member.Item2.GroupId = groupId;
                groups.Add(new HistoryChangeGroup(
                    groupId,
                    firstTransition.EdgeId,
                    firstSubject.Aspect,
                    firstTransition.BeforeStateId,
                    firstTransition.AfterStateId,
                    new ReadOnlyCollection<string>(subjectIds)));
            }

            return groups;
        }

        private static void MarkReverts(
            IList<HistoryTransition> transitions,
            IReadOnlyDictionary<string, HistoryObservation> observations,
            IReadOnlyDictionary<string, HistoryRevision> revisions)
        {
            foreach (HistoryTransition transition in transitions)
            {
                HistoryRevision cursor = revisions[transition.BeforeRevisionId];
                while (cursor.ParentRevisionId.Length != 0)
                {
                    cursor = revisions[cursor.ParentRevisionId];
                    if (string.Equals(
                        observations[cursor.RevisionId].StateId,
                        transition.AfterStateId,
                        StringComparison.Ordinal))
                    {
                        transition.IsRevert = true;
                        break;
                    }
                }
            }
        }

        private static bool MarkConflicts(
            IList<HistoryTransition> transitions,
            IReadOnlyDictionary<string, HistoryObservation> observations,
            IReadOnlyList<HistoryRevision> revisions)
        {
            bool any = false;
            foreach (IGrouping<string, HistoryRevision> siblings in revisions
                         .Where(revision => revision.ParentRevisionId.Length != 0)
                         .GroupBy(revision => revision.ParentRevisionId, StringComparer.Ordinal))
            {
                List<HistoryRevision> changedChildren = siblings
                    .Where(child => !string.Equals(
                        observations[child.RevisionId].StateId,
                        observations[child.ParentRevisionId].StateId,
                        StringComparison.Ordinal))
                    .ToList();
                if (changedChildren.Select(child => observations[child.RevisionId].StateId)
                    .Distinct(StringComparer.Ordinal).Count() < 2) continue;

                any = true;
                var ids = new HashSet<string>(
                    changedChildren.Select(child => child.RevisionId),
                    StringComparer.Ordinal);
                foreach (HistoryTransition transition in transitions)
                    if (ids.Contains(transition.AfterRevisionId)) transition.IsConflict = true;
            }
            return any;
        }

        private static bool MarkBranchOnly(
            IList<HistoryTransition> transitions,
            ISet<string> masterLineage)
        {
            bool any = false;
            foreach (HistoryTransition transition in transitions)
            {
                if (masterLineage.Contains(transition.AfterRevisionId)) continue;
                transition.IsBranchOnly = true;
                any = true;
            }
            return any;
        }

        private static bool IsSubjectReverted(
            IReadOnlyList<HistoryObservation> observations,
            ISet<string> masterLineage,
            HistoryRevision baseline,
            HistoryRevision master)
        {
            HistoryObservation endpoint = observations.Single(item => item.RevisionId == master.RevisionId);
            HistoryObservation start;
            if (baseline != null && masterLineage.Contains(baseline.RevisionId))
                start = observations.Single(item => item.RevisionId == baseline.RevisionId);
            else
                start = observations
                    .Where(item => masterLineage.Contains(item.RevisionId))
                    .OrderBy(item => item.RevisionOrder)
                    .First();

            if (!string.Equals(start.StateId, endpoint.StateId, StringComparison.Ordinal)) return false;
            return observations.Any(item =>
                masterLineage.Contains(item.RevisionId) &&
                item.RevisionOrder > start.RevisionOrder &&
                item.RevisionOrder < endpoint.RevisionOrder &&
                !string.Equals(item.StateId, start.StateId, StringComparison.Ordinal));
        }

        private static AspectState Extract(
            RevisionSnapshot snapshot,
            string subjectKey,
            CellChangeAspect aspect)
        {
            bool present = snapshot.TryGetCell(subjectKey, out CanonicalCellState cell);
            cell = cell ?? CanonicalCellState.Blank(subjectKey);
            string text = null;
            CanonicalValue value = null;
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
                    payload = text == null ? "<no-formula-metadata>" : "formula-metadata:" + text;
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
                    throw new ArgumentOutOfRangeException(nameof(aspect));
            }

            string stateId = StableId.Compute(
                "history-aspect-state-v1",
                ((int)aspect).ToString(CultureInfo.InvariantCulture),
                payload);
            return new AspectState(stateId, present, text, value);
        }

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

        private static HashSet<string> BuildLineage(
            HistoryRevision revision,
            IReadOnlyDictionary<string, HistoryRevision> byId)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            HistoryRevision cursor = revision;
            while (cursor != null)
            {
                result.Add(cursor.RevisionId);
                cursor = cursor.ParentRevisionId.Length == 0 ? null : byId[cursor.ParentRevisionId];
            }
            return result;
        }

        private static List<HistoryRevision> ValidateAndOrder(IEnumerable<HistoryRevision> revisions)
        {
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));
            List<HistoryRevision> ordered = revisions.ToList();
            if (ordered.Count < 2)
                throw new ArgumentException("History requires at least two revision occurrences.", nameof(revisions));
            if (ordered.Any(revision => revision == null))
                throw new ArgumentException("History revisions cannot contain null entries.", nameof(revisions));

            var byId = new Dictionary<string, HistoryRevision>(StringComparer.Ordinal);
            var orders = new HashSet<int>();
            foreach (HistoryRevision revision in ordered)
            {
                if (byId.ContainsKey(revision.RevisionId))
                    throw new ArgumentException("Duplicate revision ID: " + revision.RevisionId, nameof(revisions));
                byId.Add(revision.RevisionId, revision);
                if (!orders.Add(revision.Order))
                    throw new ArgumentException("Revision order values must be unique.", nameof(revisions));
                if (string.Equals(revision.RevisionId, revision.ParentRevisionId, StringComparison.Ordinal))
                    throw new ArgumentException("A revision cannot be its own parent.", nameof(revisions));
            }

            if (ordered.Count(revision => revision.IsMaster) != 1)
                throw new ArgumentException("History requires exactly one explicit master revision.", nameof(revisions));
            if (ordered.Count(revision => revision.IsTrustedBaseline) > 1)
                throw new ArgumentException("History accepts at most one trusted baseline.", nameof(revisions));
            if (ordered.Count(revision => revision.ParentRevisionId.Length == 0) != 1)
                throw new ArgumentException("History revisions must form one rooted tree.", nameof(revisions));

            foreach (HistoryRevision revision in ordered)
            {
                if (revision.ParentRevisionId.Length == 0) continue;
                if (!byId.TryGetValue(revision.ParentRevisionId, out HistoryRevision parent))
                    throw new ArgumentException(
                        "Unknown parent revision ID: " + revision.ParentRevisionId,
                        nameof(revisions));
                if (parent.Order >= revision.Order)
                    throw new ArgumentException(
                        "A parent revision must precede its child in the confirmed order.",
                        nameof(revisions));
            }

            ordered.Sort((left, right) =>
            {
                int order = left.Order.CompareTo(right.Order);
                return order != 0 ? order : string.CompareOrdinal(left.RevisionId, right.RevisionId);
            });
            return ordered;
        }

        private static void Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A stable identity component is required.", name);
        }

        private sealed class EdgeChange
        {
            internal EdgeChange(
                string edgeId,
                HistoryRevision before,
                HistoryRevision after,
                AtomicCellChange change,
                AspectState beforeState,
                AspectState afterState,
                bool isOnMasterLineage)
            {
                EdgeId = edgeId;
                Before = before;
                After = after;
                Change = change;
                BeforeState = beforeState;
                AfterState = afterState;
                IsOnMasterLineage = isOnMasterLineage;
            }

            internal string EdgeId { get; }
            internal HistoryRevision Before { get; }
            internal HistoryRevision After { get; }
            internal AtomicCellChange Change { get; }
            internal AspectState BeforeState { get; }
            internal AspectState AfterState { get; }
            internal bool IsOnMasterLineage { get; }
        }

        private sealed class AspectState
        {
            internal AspectState(string stateId, bool isPresent, string text, CanonicalValue value)
            {
                StateId = stateId;
                IsPresent = isPresent;
                Text = text;
                Value = value;
            }

            internal string StateId { get; }
            internal bool IsPresent { get; }
            internal string Text { get; }
            internal CanonicalValue Value { get; }
        }

        private sealed class SubjectAspect : IEquatable<SubjectAspect>
        {
            internal SubjectAspect(string subjectKey, CellChangeAspect aspect)
            {
                SubjectKey = subjectKey;
                Aspect = aspect;
            }

            internal string SubjectKey { get; }
            internal CellChangeAspect Aspect { get; }

            public bool Equals(SubjectAspect other) => other != null &&
                Aspect == other.Aspect &&
                string.Equals(SubjectKey, other.SubjectKey, StringComparison.Ordinal);

            public override bool Equals(object obj) => Equals(obj as SubjectAspect);

            public override int GetHashCode()
            {
                unchecked
                {
                    return (StringComparer.Ordinal.GetHashCode(SubjectKey) * 397) ^ (int)Aspect;
                }
            }
        }
    }
}
