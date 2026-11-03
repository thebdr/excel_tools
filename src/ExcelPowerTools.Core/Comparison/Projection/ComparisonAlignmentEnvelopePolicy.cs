using System;
using ExcelPowerTools.Core.Comparison.Adapters.OpenXml;

namespace ExcelPowerTools.Core.Comparison.Projection
{
    public enum ComparisonAlignmentEnvelopeLimit
    {
        None = 0,
        WorkbookStoredCells = 1,
        AggregateStoredCells = 2,
        AggregateRetainedStoredCells = 3,
        AggregateGridCells = 4,
    }

    public sealed class ComparisonAlignmentEnvelopeAssessment
    {
        internal ComparisonAlignmentEnvelopeAssessment(ComparisonAlignmentEnvelopeLimit exceededLimit)
        {
            ExceededLimit = exceededLimit;
        }

        public ComparisonAlignmentEnvelopeLimit ExceededLimit { get; }
        public bool CanAnalyze => ExceededLimit == ComparisonAlignmentEnvelopeLimit.None;
    }

    /// <summary>
    /// Admission policy for the materialized alignment work. The physically stored count is accepted for
    /// diagnostics only: the OOXML adapter has already applied its package safety envelope, and compacted blank
    /// tails do not remain in the project snapshots or expand an alignment grid.
    /// </summary>
    public static class ComparisonAlignmentEnvelopePolicy
    {
        public const long MaximumWorkbookStoredCells =
            OpenXmlWorkbookSnapshotAdapter.MaxPhysicalCellElementsPerWorkbook;
        public const long MaximumAggregateStoredCells = 10000000L;
        public const long MaximumAggregateRetainedStoredCells = 2000000L;
        public const long MaximumAggregateGridCells = 10000000L;

        public static ComparisonAlignmentEnvelopeAssessment Evaluate(
            long workbookStoredCellCount,
            long aggregateStoredCellCount,
            long aggregateRetainedStoredCellCount,
            long aggregateGridCellCount)
        {
            ValidateNonNegative(workbookStoredCellCount, nameof(workbookStoredCellCount));
            ValidateNonNegative(aggregateStoredCellCount, nameof(aggregateStoredCellCount));
            ValidateNonNegative(
                aggregateRetainedStoredCellCount,
                nameof(aggregateRetainedStoredCellCount));
            ValidateNonNegative(aggregateGridCellCount, nameof(aggregateGridCellCount));

            ComparisonAlignmentEnvelopeLimit limit = workbookStoredCellCount > MaximumWorkbookStoredCells
                ? ComparisonAlignmentEnvelopeLimit.WorkbookStoredCells
                : aggregateStoredCellCount > MaximumAggregateStoredCells
                    ? ComparisonAlignmentEnvelopeLimit.AggregateStoredCells
                    : aggregateRetainedStoredCellCount > MaximumAggregateRetainedStoredCells
                        ? ComparisonAlignmentEnvelopeLimit.AggregateRetainedStoredCells
                        : aggregateGridCellCount > MaximumAggregateGridCells
                            ? ComparisonAlignmentEnvelopeLimit.AggregateGridCells
                            : ComparisonAlignmentEnvelopeLimit.None;
            return new ComparisonAlignmentEnvelopeAssessment(limit);
        }

        private static void ValidateNonNegative(long value, string parameterName)
        {
            if (value < 0) throw new ArgumentOutOfRangeException(parameterName);
        }
    }
}
