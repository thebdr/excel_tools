using System;
using System.Collections.Generic;
using ExcelPowerTools.Core.Comparison.Identifiers;

namespace ExcelPowerTools.Core.Comparison.Model
{
    /// <summary>
    /// One occurrence in the revision graph. RevisionId is occurrence identity; ByteContentId identifies exact
    /// bytes and may therefore be shared by multiple occurrences.
    /// </summary>
    public sealed class Revision
    {
        private string _byteContentId = "";

        public string RevisionId { get; set; } = "";
        public string ByteContentId
        {
            get => _byteContentId;
            set => AssignByteContentId(value, nameof(ByteContentId));
        }
        public string Label { get; set; } = "";
        public int Order { get; set; }
        public string ParentRevisionId { get; set; } = "";
        public bool IsMaster { get; set; }
        public bool IsTrustedBaseline { get; set; }

        public RevisionFormatKind Format { get; set; }
        public string SourceLabel { get; set; } = "";
        public string OriginalBaseName { get; set; } = "";
        public SourcePathClass PathClass { get; set; }
        public string LocalLocatorKey { get; set; } = "";
        public long ByteLength { get; set; }
        /// <summary>
        /// Compatibility alias for <see cref="ByteContentId"/>. Both serialized names converge on the same
        /// validated backing value, so legacy payloads cannot create contradictory byte identities.
        /// </summary>
        public string Sha256
        {
            get => _byteContentId;
            set => AssignByteContentId(value, nameof(Sha256));
        }
        public ArtifactTimestamps SourceTimestamps { get; set; } = new ArtifactTimestamps();
        public string Description { get; set; } = "";

        public string DocumentCreator { get; set; } = "";
        public string LastModifiedBy { get; set; } = "";
        public string ApplicationVersion { get; set; } = "";
        public IList<MetadataEntry> CloudVersionMetadata { get; set; } = new List<MetadataEntry>();
        public ProvenanceReliability ProvenanceReliability { get; set; }

        public string ParserVersion { get; set; } = "";
        public string ImportProfileId { get; set; } = "";
        public RevisionCoverage Coverage { get; set; } = new RevisionCoverage();

        private void AssignByteContentId(string value, string propertyName)
        {
            value = value ?? "";
            if (value.Length == 0)
            {
                if (_byteContentId.Length != 0)
                    throw new InvalidOperationException(
                        propertyName + " cannot clear an already established byte-content identity.");
                return;
            }
            if (!StableId.IsSha256(value))
                throw new ArgumentException("A byte-content identity must be a SHA-256 hexadecimal digest.", propertyName);

            string normalized = value.ToUpperInvariant();
            if (_byteContentId.Length != 0 &&
                !string.Equals(_byteContentId, normalized, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "ByteContentId and Sha256 cannot identify different artifact bytes.");
            }
            _byteContentId = normalized;
        }
    }

    public sealed class ArtifactTimestamps
    {
        public DateTimeOffset? CreatedUtc { get; set; }
        public DateTimeOffset? ModifiedUtc { get; set; }
        public DateTimeOffset? LastAccessedUtc { get; set; }
    }

    public sealed class MetadataEntry
    {
        public string Key { get; set; } = "";
        public string Value { get; set; } = "";
        public ProvenanceReliability Reliability { get; set; }
    }

    /// <summary>Coverage facts prevent an unsupported package change from being reported as identical.</summary>
    public sealed class RevisionCoverage
    {
        public bool FormulaAndValuesComplete { get; set; }
        public bool StructureComplete { get; set; }
        public string SupportedContentHash { get; set; } = "";
        public string UnsupportedContentHash { get; set; } = "";
        public IList<string> UnsupportedFeatures { get; set; } = new List<string>();
    }
}
