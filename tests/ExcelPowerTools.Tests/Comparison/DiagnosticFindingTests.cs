using System;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.Identifiers;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class DiagnosticFindingTests
    {
        [Fact]
        public void Finding_uid_is_stable_when_evidence_arrives_in_a_different_order()
        {
            string first = FindingUidBuilder.Build(
                "CMP-SHEET-MAPPING-AMBIGUOUS",
                "logical-sheet:orders",
                new FindingEvidence("candidate", "sheet-b"),
                new FindingEvidence("threshold", "0.80"));
            string second = FindingUidBuilder.Build(
                "cmp-sheet-mapping-ambiguous",
                "logical-sheet:orders",
                new FindingEvidence("threshold", "0.80"),
                new FindingEvidence("candidate", "sheet-b"));

            Assert.Equal(first, second);
            Assert.True(StableId.IsSha256(first));
        }

        [Fact]
        public void Finding_uid_changes_with_semantic_evidence_but_has_no_severity_or_message_input()
        {
            string first = FindingUidBuilder.Build(
                "CMP-HEADER-WEAK",
                "revision-1/sheet-1",
                new FindingEvidence("candidate-row", "3"));
            string second = FindingUidBuilder.Build(
                "CMP-HEADER-WEAK",
                "revision-1/sheet-1",
                new FindingEvidence("candidate-row", "4"));

            Assert.NotEqual(first, second);

            var warning = new DiagnosticFinding
            {
                FindingUid = first,
                Severity = DiagnosticSeverity.Warning,
                MessageKey = "HeaderWeak"
            };
            var blocking = new DiagnosticFinding
            {
                FindingUid = first,
                Severity = DiagnosticSeverity.ArtifactBlocking,
                MessageKey = "HeaderWeakVerbose"
            };
            Assert.Equal(warning.FindingUid, blocking.FindingUid);
        }

        [Fact]
        public void Exception_projection_preserves_the_complete_inner_chain()
        {
            var outer = new InvalidOperationException(
                "outer",
                new ArgumentException("middle", new FormatException("inner")));

            DiagnosticExceptionInfo projected = DiagnosticExceptionInfo.FromException(outer);

            Assert.Equal(typeof(InvalidOperationException).FullName, projected.Type);
            Assert.Equal(typeof(ArgumentException).FullName, projected.InnerException.Type);
            Assert.Equal(typeof(FormatException).FullName, projected.InnerException.InnerException.Type);
            Assert.Null(projected.InnerException.InnerException.InnerException);
        }

        [Fact]
        public void Diagnostic_collections_are_structured_and_initialized()
        {
            var finding = new DiagnosticFinding();

            Assert.NotNull(finding.Artifacts);
            Assert.NotNull(finding.Locations);
            Assert.NotNull(finding.MessageArguments);
            Assert.NotNull(finding.Inference.Candidates);
            Assert.NotNull(finding.SuggestedActions);
        }
    }
}
