using System;
using System.Collections.Generic;

namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>
    /// Durable project manifest. Large canonical observations may be stored in versioned chunks; this object
    /// owns their generation identifiers and the revision graph, not a UI projection of the evidence.
    /// </summary>
    public sealed class ComparisonProject
    {
        public string ProjectId { get; set; } = "";
        public string SchemaVersion { get; set; } = "1";
        public DateTimeOffset CreatedUtc { get; set; }
        public DateTimeOffset UpdatedUtc { get; set; }
        public string OptionsHash { get; set; } = "";
        public string ThemeVersion { get; set; } = "1";
        public string EvidenceGenerationId { get; set; } = "";
        public string MappingGenerationId { get; set; } = "";
        public string ProjectStateRoot { get; set; } = "";
        public IList<Revision> Revisions { get; set; } = new List<Revision>();
    }
}
