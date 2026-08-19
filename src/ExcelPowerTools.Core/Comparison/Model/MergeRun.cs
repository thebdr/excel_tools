using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>A frozen, auditable merge execution against a specific evidence and mapping generation.</summary>
    public sealed class MergeRun
    {
        public string MergeId { get; set; } = "";
        public string ProjectId { get; set; } = "";
        public string SourceMasterRevisionId { get; set; } = "";
        public string OutputRevisionId { get; set; } = "";
        public string GenerationId { get; set; } = "";
        public string MappingGenerationId { get; set; } = "";
        public string MergePlanHash { get; set; } = "";
        public IList<string> SelectedDecisionIds { get; set; } = new List<string>();
        public MergeRunState State { get; set; }
        public IList<MergeCheckResult> PreconditionResults { get; set; } = new List<MergeCheckResult>();
        public string OutputHash { get; set; } = "";
        public IList<MergeCheckResult> ValidationResults { get; set; } = new List<MergeCheckResult>();
        public DateTimeOffset? PreparedUtc { get; set; }
        public DateTimeOffset? ValidatedUtc { get; set; }
        public DateTimeOffset? OutputPublishedUtc { get; set; }
        public DateTimeOffset? AuditCommittedUtc { get; set; }
        public string FailureFindingUid { get; set; } = "";
    }

    public sealed class MergeCheckResult
    {
        public string Code { get; set; } = "";
        public MergeCheckStatus Status { get; set; }
        public string ExpectedHash { get; set; } = "";
        public string ActualHash { get; set; } = "";
        public string Detail { get; set; } = "";
    }
}
