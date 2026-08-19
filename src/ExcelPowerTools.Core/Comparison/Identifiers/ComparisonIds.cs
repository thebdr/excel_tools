using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Identifiers
{
    /// <summary>Typed constructors for durable comparison identities.</summary>
    public static class ComparisonIds
    {
        public static string Edge(string projectId, string beforeRevisionId, string afterRevisionId)
        {
            Require(projectId, nameof(projectId));
            Require(beforeRevisionId, nameof(beforeRevisionId));
            Require(afterRevisionId, nameof(afterRevisionId));
            return StableId.Compute("revision-edge-v1", projectId, beforeRevisionId, afterRevisionId);
        }

        public static string Subject(
            string logicalSheetId,
            string rowIdentity,
            string columnIdentity,
            ComparisonAspect aspect)
        {
            Require(logicalSheetId, nameof(logicalSheetId));
            Require(rowIdentity, nameof(rowIdentity));
            Require(columnIdentity, nameof(columnIdentity));
            return StableId.Compute(
                "logical-subject-v1",
                logicalSheetId,
                rowIdentity,
                columnIdentity,
                Number(aspect));
        }

        public static string Observation(
            string subjectId,
            string revisionId,
            ObservationExactType exactType,
            string contentHash)
        {
            Require(subjectId, nameof(subjectId));
            Require(revisionId, nameof(revisionId));
            Require(contentHash, nameof(contentHash));
            return StableId.Compute(
                "observation-v1",
                subjectId,
                revisionId,
                Number(exactType),
                contentHash);
        }

        /// <summary>
        /// Before/after IDs may be null for an addition/removal. Null remains distinct from an empty, malformed
        /// identifier so corrupted data cannot accidentally acquire a valid transition identity.
        /// </summary>
        public static string Transition(
            string edgeId,
            string subjectId,
            string beforeObservationId,
            string afterObservationId)
        {
            Require(edgeId, nameof(edgeId));
            Require(subjectId, nameof(subjectId));
            RequireOptional(beforeObservationId, nameof(beforeObservationId));
            RequireOptional(afterObservationId, nameof(afterObservationId));
            if (beforeObservationId == null && afterObservationId == null)
                throw new ArgumentException("A transition requires at least one observation endpoint.");
            return StableId.Compute(
                "transition-v1",
                edgeId,
                subjectId,
                beforeObservationId,
                afterObservationId);
        }

        public static string Group(string generationId, string groupingRule, params string[] orderedSubjectIds)
        {
            Require(generationId, nameof(generationId));
            Require(groupingRule, nameof(groupingRule));

            if (orderedSubjectIds == null || orderedSubjectIds.Length == 0)
                throw new ArgumentException("A change group requires at least one logical subject.", nameof(orderedSubjectIds));

            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < orderedSubjectIds.Length; i++)
            {
                Require(orderedSubjectIds[i], nameof(orderedSubjectIds));
                if (!seen.Add(orderedSubjectIds[i]))
                    throw new ArgumentException("A change group cannot contain duplicate logical subjects.", nameof(orderedSubjectIds));
            }

            var components = new string[orderedSubjectIds.Length + 2];
            components[0] = generationId;
            components[1] = groupingRule;
            Array.Copy(orderedSubjectIds, 0, components, 2, orderedSubjectIds.Length);
            return StableId.Compute("change-group-v1", components);
        }

        private static string Number<T>(T value) where T : struct
        {
            return Convert.ToInt32(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
        }

        private static void Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A stable identity component is required.", name);
        }

        private static void RequireOptional(string value, string name)
        {
            if (value != null && string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("An optional stable identity component must be null or non-blank.", name);
        }
    }
}
