using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>One exact revision occurrence contributing bytes to a review evidence generation.</summary>
    public sealed class ReviewEvidenceRevisionIdentity
    {
        public ReviewEvidenceRevisionIdentity(
            string revisionId,
            int order,
            string byteContentId)
        {
            if (string.IsNullOrWhiteSpace(revisionId))
                throw new ArgumentException("A revision occurrence ID is required.", nameof(revisionId));
            if (order < 0) throw new ArgumentOutOfRangeException(nameof(order));
            if (!StableId.IsSha256(byteContentId))
                throw new ArgumentException(
                    "A revision byte-content identity must be a SHA-256 value.",
                    nameof(byteContentId));
            RevisionId = revisionId;
            Order = order;
            ByteContentId = byteContentId.ToUpperInvariant();
        }

        public string RevisionId { get; }
        public int Order { get; }
        public string ByteContentId { get; }
    }

    /// <summary>
    /// Stable identity for the exact ordered revision bytes interpreted by one effective mapping generation.
    /// Enumeration order is ignored; explicit revision Order is authoritative and must be unique.
    /// </summary>
    public sealed class ReviewEvidenceGeneration
    {
        public ReviewEvidenceGeneration(
            string projectId,
            string effectiveMappingGenerationId,
            IEnumerable<ReviewEvidenceRevisionIdentity> revisions)
        {
            if (string.IsNullOrWhiteSpace(projectId))
                throw new ArgumentException("A project ID is required.", nameof(projectId));
            if (string.IsNullOrWhiteSpace(effectiveMappingGenerationId))
                throw new ArgumentException(
                    "An effective mapping generation ID is required.",
                    nameof(effectiveMappingGenerationId));
            if (revisions == null) throw new ArgumentNullException(nameof(revisions));

            var ordered = new List<ReviewEvidenceRevisionIdentity>(revisions);
            if (ordered.Count == 0)
                throw new ArgumentException(
                    "An evidence generation requires at least one revision occurrence.",
                    nameof(revisions));
            if (ordered.Any(item => item == null))
                throw new ArgumentException(
                    "Evidence revisions cannot contain null entries.",
                    nameof(revisions));
            if (ordered.Select(item => item.RevisionId).Distinct(StringComparer.Ordinal).Count() != ordered.Count)
                throw new ArgumentException(
                    "Evidence revision occurrence IDs must be unique.",
                    nameof(revisions));
            if (ordered.Select(item => item.Order).Distinct().Count() != ordered.Count)
                throw new ArgumentException(
                    "Evidence revision orders must be unique.",
                    nameof(revisions));
            ordered.Sort((left, right) => left.Order.CompareTo(right.Order));

            ProjectId = projectId;
            EffectiveMappingGenerationId = effectiveMappingGenerationId;
            Revisions = new ReadOnlyCollection<ReviewEvidenceRevisionIdentity>(ordered);

            var components = new List<string>
            {
                ProjectId,
                EffectiveMappingGenerationId,
                Revisions.Count.ToString(CultureInfo.InvariantCulture),
            };
            foreach (ReviewEvidenceRevisionIdentity revision in Revisions)
            {
                components.Add(revision.Order.ToString(CultureInfo.InvariantCulture));
                components.Add(revision.RevisionId);
                components.Add(revision.ByteContentId);
            }
            GenerationId = StableId.Compute(
                "review-evidence-generation-v1",
                components.ToArray());
        }

        public string ProjectId { get; }
        public string EffectiveMappingGenerationId { get; }
        public IReadOnlyList<ReviewEvidenceRevisionIdentity> Revisions { get; }
        public string GenerationId { get; }
    }
}
