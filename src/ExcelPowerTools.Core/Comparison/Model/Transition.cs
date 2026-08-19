using System.Collections.Generic;

namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>A comparison result across one directed revision edge.</summary>
    public sealed class Transition
    {
        public string TransitionId { get; set; } = "";
        public string EdgeId { get; set; } = "";
        public string SubjectId { get; set; } = "";
        public string BeforeObservationId { get; set; } = "";
        public string AfterObservationId { get; set; } = "";
        public TransitionChangeKind ChangeKind { get; set; }
        public double AlignmentConfidence { get; set; }
        public string GroupId { get; set; } = "";
        public IList<string> EquivalenceTags { get; set; } = new List<string>();
    }
}
