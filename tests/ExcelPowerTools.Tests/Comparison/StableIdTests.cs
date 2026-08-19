using System;
using System.Text;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public class StableIdTests
    {
        [Fact]
        public void Structural_id_is_deterministic_and_sha256_shaped()
        {
            string first = StableId.Compute("test-domain-v1", "alpha", "βeta", "");
            string second = StableId.Compute("test-domain-v1", "alpha", "βeta", "");

            Assert.Equal(first, second);
            Assert.Equal("E3BB95083C69EC3555A78EA617A95DFF6C9D35929EB3EB9F927BAC2933F4E02F", first);
            Assert.True(StableId.IsSha256(first));
            Assert.Equal(first.ToUpperInvariant(), first);
        }

        [Fact]
        public void Length_prefixes_protect_component_boundaries()
        {
            Assert.NotEqual(
                StableId.Compute("test-domain-v1", "ab", "c"),
                StableId.Compute("test-domain-v1", "a", "bc"));
        }

        [Fact]
        public void Null_empty_domain_and_unicode_are_distinct()
        {
            Assert.NotEqual(
                StableId.Compute("test-domain-v1", new string[] { null! }),
                StableId.Compute("test-domain-v1", ""));
            Assert.NotEqual(
                StableId.Compute("first-domain-v1", "same"),
                StableId.Compute("second-domain-v1", "same"));
            Assert.NotEqual(
                StableId.Compute("test-domain-v1", "é"),
                StableId.Compute("test-domain-v1", "e\u0301"));
        }

        [Fact]
        public void Malformed_utf16_components_are_rejected_instead_of_replacement_colliding()
        {
            Assert.Throws<ArgumentException>(() => StableId.Compute("test-domain-v1", "\uD800"));
            Assert.Throws<ArgumentException>(() => StableId.Compute("test-domain-v1", "\uD801"));
        }

        [Fact]
        public void Exact_byte_content_id_uses_raw_bytes()
        {
            Assert.Equal(
                "B5D4045C3F466FA91FE2CC6ABE79232A1A57CDF104F7A26E716E0A1E2789DF78",
                StableId.ComputeByteContentId(Encoding.UTF8.GetBytes("ABC")));
            Assert.NotEqual(
                StableId.ComputeByteContentId(Encoding.UTF8.GetBytes("ABC")),
                StableId.ComputeByteContentId(Encoding.UTF8.GetBytes("abc")));
            Assert.Equal(
                StableId.ForBytes(Encoding.UTF8.GetBytes("ABC")),
                StableId.ComputeByteContentId(Encoding.UTF8.GetBytes("ABC")));
            Assert.Throws<ArgumentNullException>(() => StableId.ForBytes(null!));
            Assert.Throws<ArgumentNullException>(() => StableId.ComputeByteContentId(null!));
        }

        [Fact]
        public void Typed_identifiers_include_every_semantic_component()
        {
            string formula = ComparisonIds.Subject("sheet-1", "row:customer-7", "column:total", ComparisonAspect.Formula);
            string value = ComparisonIds.Subject("sheet-1", "row:customer-7", "column:total", ComparisonAspect.Value);
            string edge = ComparisonIds.Edge("project-1", "revision-1", "revision-2");
            string added = ComparisonIds.Transition(edge, formula, null, "observation-2");

            Assert.NotEqual(formula, value);
            Assert.True(StableId.IsSha256(added));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Subject("", "row", "column", ComparisonAspect.Value));
        }

        [Fact]
        public void Transition_and_group_ids_reject_malformed_membership()
        {
            string edge = ComparisonIds.Edge("project", "before", "after");

            Assert.True(StableId.IsSha256(ComparisonIds.Transition(edge, "subject", null, "after-observation")));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Transition(edge, "subject", null, null));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Transition(edge, "subject", "", "after-observation"));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Transition(edge, "subject", "before-observation", " "));

            Assert.True(StableId.IsSha256(ComparisonIds.Group("generation", "rule", "subject-a")));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Group("generation", "rule"));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Group("generation", "rule", Array.Empty<string>()));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Group("generation", "rule", "subject-a", ""));
            Assert.Throws<ArgumentException>(() => ComparisonIds.Group("generation", "rule", "subject-a", "subject-a"));
        }
    }
}
