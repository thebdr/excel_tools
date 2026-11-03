using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    public sealed class RowEvidenceIdentity : IEquatable<RowEvidenceIdentity>
    {
        public RowEvidenceIdentity(string artifactId, string sheet, int row)
        {
            if (string.IsNullOrWhiteSpace(artifactId) || string.IsNullOrWhiteSpace(sheet) || row < 1 || row > 1048576)
                throw new ArgumentException("An identity needs an artifact, sheet and valid physical row.");
            ArtifactId = artifactId; Sheet = sheet; Row = row;
        }
        public string ArtifactId { get; }
        public string Sheet { get; }
        public int Row { get; }
        public bool Equals(RowEvidenceIdentity other) => other != null && ArtifactId == other.ArtifactId && Sheet == other.Sheet && Row == other.Row;
        public override bool Equals(object obj) => Equals(obj as RowEvidenceIdentity);
        public override int GetHashCode() => (ArtifactId + "\0" + Sheet + "\0" + Row).GetHashCode();
        public override string ToString() => ArtifactId + "/" + Sheet + "/" + Row;
    }

    /// <summary>Accepted pair evidence is not an ancestry tree or automatic merge authority.</summary>
    public sealed class RowIdentityLedger
    {
        private readonly List<HashSet<RowEvidenceIdentity>> groups = new List<HashSet<RowEvidenceIdentity>>();
        private readonly List<Tuple<RowEvidenceIdentity, RowEvidenceIdentity>> direct = new List<Tuple<RowEvidenceIdentity, RowEvidenceIdentity>>();
        public ReadOnlyCollection<ReadOnlyCollection<RowEvidenceIdentity>> Groups => Array.AsReadOnly(groups
            .Select(g => Array.AsReadOnly(g.OrderBy(i => i.ArtifactId, StringComparer.Ordinal).ThenBy(i => i.Sheet, StringComparer.Ordinal).ThenBy(i => i.Row).ToArray())).ToArray());

        public bool TryAccept(RowEvidenceIdentity a, RowEvidenceIdentity b, out ReviewIssue issue)
        {
            if (a == null || b == null) throw new ArgumentNullException();
            var ag = groups.FirstOrDefault(g => g.Contains(a)); var bg = groups.FirstOrDefault(g => g.Contains(b));
            var union = new HashSet<RowEvidenceIdentity>(ag ?? new HashSet<RowEvidenceIdentity> { a });
            union.UnionWith(bg ?? new HashSet<RowEvidenceIdentity> { b });
            var conflict = union.GroupBy(i => i.ArtifactId, StringComparer.Ordinal).FirstOrDefault(g => g.Count() > 1);
            if (conflict != null)
            {
                issue = ReviewIssue.Create("SR.IDENTITY.NON_TRANSITIVE", "The proposed row history would contain two different observations from the same revision. Keep these histories separate and review the conflicting pair.",
                    new Dictionary<string, string> { ["before"] = a.ToString(), ["after"] = b.ToString(),
                        ["conflictingObservations"] = string.Join(", ", conflict.Select(i => i.ToString())) });
                return false;
            }
            // All validation precedes mutation, including the direct-evidence list.
            if (ag != null) groups.Remove(ag);
            if (bg != null && !ReferenceEquals(ag, bg)) groups.Remove(bg);
            groups.Add(union);
            if (!HasDirectEvidence(a, b)) direct.Add(Tuple.Create(a, b));
            issue = null; return true;
        }
        public bool HasDirectEvidence(RowEvidenceIdentity a, RowEvidenceIdentity b) => direct.Any(e =>
            (e.Item1.Equals(a) && e.Item2.Equals(b)) || (e.Item1.Equals(b) && e.Item2.Equals(a)));
    }
}
