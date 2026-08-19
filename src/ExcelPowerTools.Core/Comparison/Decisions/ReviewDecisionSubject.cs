using System;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.Core.Comparison.Decisions
{
    /// <summary>
    /// Explicit authority granted to reviewer choices for one logical subject. Unspecified exists only so an
    /// older subject-ID-only payload can still hydrate safely; it never authorizes observation selection.
    /// </summary>
    public enum ReviewDecisionSubjectPolicy
    {
        Unspecified = 0,
        ReviewOnly = 1,
        Selectable = 2,
        SelectableAndMergeable = 3,
    }

    /// <summary>
    /// Immutable semantic identity and selection policy for one independently reviewable subject. Formula and
    /// value are the only aspects that may currently authorize an observation selection or future content merge.
    /// </summary>
    public sealed class ReviewDecisionSubject
    {
        public ReviewDecisionSubject(
            string subjectId,
            string logicalSubjectKey = "",
            ComparisonAspect aspect = ComparisonAspect.Unknown,
            ReviewDecisionSubjectPolicy policy = ReviewDecisionSubjectPolicy.Unspecified)
        {
            SubjectId = Require(subjectId, nameof(subjectId));
            if (!Enum.IsDefined(typeof(ComparisonAspect), aspect))
                throw new ArgumentOutOfRangeException(nameof(aspect));
            if (!Enum.IsDefined(typeof(ReviewDecisionSubjectPolicy), policy))
                throw new ArgumentOutOfRangeException(nameof(policy));

            string key = logicalSubjectKey ?? "";
            if (policy == ReviewDecisionSubjectPolicy.Unspecified)
            {
                if (aspect != ComparisonAspect.Unknown)
                    throw new ArgumentException(
                        "Legacy unspecified subject metadata cannot claim an aspect.",
                        nameof(aspect));
                if (key.Length != 0 && string.IsNullOrWhiteSpace(key))
                    throw new ArgumentException(
                        "A logical subject key must be empty or non-blank.",
                        nameof(logicalSubjectKey));
                key = key.Length == 0 ? SubjectId : key;
            }
            else
            {
                key = Require(key, nameof(logicalSubjectKey));
                if (aspect == ComparisonAspect.Unknown)
                    throw new ArgumentException(
                        "Explicit decision metadata requires a known subject aspect.",
                        nameof(aspect));
                if ((policy == ReviewDecisionSubjectPolicy.Selectable ||
                     policy == ReviewDecisionSubjectPolicy.SelectableAndMergeable) &&
                    aspect != ComparisonAspect.Formula &&
                    aspect != ComparisonAspect.Value)
                    throw new ArgumentException(
                        "Only formula and value subjects are selectable in the current review contract.",
                        nameof(policy));
            }

            LogicalSubjectKey = key;
            Aspect = aspect;
            Policy = policy;
        }

        public string SubjectId { get; }
        public string LogicalSubjectKey { get; }
        public ComparisonAspect Aspect { get; }
        public ReviewDecisionSubjectPolicy Policy { get; }
        public bool HasExplicitMetadata => Policy != ReviewDecisionSubjectPolicy.Unspecified;
        public bool CanSelectObservation =>
            Policy == ReviewDecisionSubjectPolicy.Selectable ||
            Policy == ReviewDecisionSubjectPolicy.SelectableAndMergeable;
        public bool CanMergeSelection => Policy == ReviewDecisionSubjectPolicy.SelectableAndMergeable;

        private static string Require(string value, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A non-blank durable subject identity is required.", name);
            return value;
        }
    }
}
