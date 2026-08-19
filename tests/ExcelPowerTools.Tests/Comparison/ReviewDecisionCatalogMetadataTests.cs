using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using ExcelPowerTools.Core.Comparison.Canonical;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Model;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class ReviewDecisionCatalogMetadataTests
    {
        private static readonly DateTimeOffset When = new DateTimeOffset(
            2026, 7, 17, 12, 0, 0, TimeSpan.Zero);

        [Fact]
        public void Evidence_generation_is_deterministic_over_explicit_order_hashes_and_effective_mapping()
        {
            string firstHash = Hash("first");
            string secondHash = Hash("second");
            var left = new ReviewEvidenceGeneration(
                "project",
                "effective-mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("revision-b", 1, secondHash.ToLowerInvariant()),
                    new ReviewEvidenceRevisionIdentity("revision-a", 0, firstHash),
                });
            var right = new ReviewEvidenceGeneration(
                "project",
                "effective-mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("revision-a", 0, firstHash.ToLowerInvariant()),
                    new ReviewEvidenceRevisionIdentity("revision-b", 1, secondHash),
                });

            Assert.Equal(left.GenerationId, right.GenerationId);
            Assert.True(StableId.IsSha256(left.GenerationId));
            Assert.Equal(new[] { "revision-a", "revision-b" },
                left.Revisions.Select(item => item.RevisionId));
            Assert.All(left.Revisions, item => Assert.Equal(item.ByteContentId.ToUpperInvariant(), item.ByteContentId));
            Assert.Throws<NotSupportedException>(() =>
                ((IList<ReviewEvidenceRevisionIdentity>)left.Revisions).Clear());

            Assert.NotEqual(left.GenerationId, new ReviewEvidenceGeneration(
                "project",
                "other-mapping",
                left.Revisions).GenerationId);
            Assert.NotEqual(left.GenerationId, new ReviewEvidenceGeneration(
                "project",
                "effective-mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("revision-a", 1, firstHash),
                    new ReviewEvidenceRevisionIdentity("revision-b", 0, secondHash),
                }).GenerationId);
            Assert.NotEqual(left.GenerationId, new ReviewEvidenceGeneration(
                "project",
                "effective-mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("revision-a", 0, firstHash),
                    new ReviewEvidenceRevisionIdentity("revision-b", 1, Hash("changed")),
                }).GenerationId);
        }

        [Fact]
        public void Evidence_generation_rejects_malformed_or_ambiguous_revision_identity()
        {
            string firstHash = Hash("first");
            string secondHash = Hash("second");
            Assert.Throws<ArgumentException>(() =>
                new ReviewEvidenceRevisionIdentity("revision", 0, "not-a-hash"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new ReviewEvidenceRevisionIdentity("revision", -1, firstHash));
            Assert.Throws<ArgumentException>(() => new ReviewEvidenceGeneration(
                "project", "mapping", Array.Empty<ReviewEvidenceRevisionIdentity>()));
            Assert.Throws<ArgumentException>(() => new ReviewEvidenceGeneration(
                "project",
                "mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("same", 0, firstHash),
                    new ReviewEvidenceRevisionIdentity("same", 1, secondHash),
                }));
            Assert.Throws<ArgumentException>(() => new ReviewEvidenceGeneration(
                "project",
                "mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("first", 0, firstHash),
                    new ReviewEvidenceRevisionIdentity("second", 0, secondHash),
                }));
        }

        [Fact]
        public void Catalog_freezes_explicit_subject_metadata_and_binds_it_into_v2_identity()
        {
            ReviewEvidenceGeneration evidence = Evidence();
            ReviewDecisionSubject formula = Subject(
                "subject-formula",
                "logical-cell/formula",
                ComparisonAspect.Formula,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            ReviewDecisionSubject value = Subject(
                "subject-value",
                "logical-cell/value",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.Selectable);
            ReviewDecisionSubject structure = Subject(
                "subject-structure",
                "logical-cell/structure",
                ComparisonAspect.Structure,
                ReviewDecisionSubjectPolicy.ReviewOnly);

            var subjects = new List<ReviewDecisionSubject> { value, structure, formula };
            Observation[] observations = EvidenceObservations(subjects);
            ReviewDecisionCatalog left = new ReviewDecisionCatalog(
                evidence,
                "master",
                subjects,
                observations);
            subjects.Clear();
            ReviewDecisionSubject[] orderedSubjects = { formula, value, structure };
            ReviewDecisionCatalog right = new ReviewDecisionCatalog(
                evidence,
                "master",
                orderedSubjects,
                EvidenceObservations(orderedSubjects));

            Assert.Equal(left.CatalogId, right.CatalogId);
            Assert.Equal(evidence.GenerationId, left.EvidenceGenerationId);
            Assert.Equal(evidence.EffectiveMappingGenerationId, left.MappingGenerationId);
            Assert.Same(evidence, left.EvidenceGeneration);
            Assert.Equal(new[] { "subject-formula", "subject-structure", "subject-value" }, left.SubjectIds);
            Assert.True(left.TryGetSubject("subject-formula", out ReviewDecisionSubject restored));
            Assert.Same(formula, restored);
            Assert.True(restored.CanSelectObservation);
            Assert.True(restored.CanMergeSelection);
            Assert.False(structure.CanSelectObservation);
            Assert.False(structure.CanMergeSelection);
            Assert.Throws<NotSupportedException>(() =>
                ((IList<ReviewDecisionSubject>)left.Subjects).Clear());

            ReviewDecisionCatalog changedPolicy = new ReviewDecisionCatalog(
                evidence,
                "master",
                new[]
                {
                    Subject(
                        "subject-formula",
                        "logical-cell/formula",
                        ComparisonAspect.Formula,
                        ReviewDecisionSubjectPolicy.Selectable),
                    value,
                    structure,
                },
                EvidenceObservations(orderedSubjects));
            Assert.NotEqual(left.CatalogId, changedPolicy.CatalogId);
        }

        [Fact]
        public void Subject_metadata_rejects_partial_unknown_or_non_formula_value_selection_authority()
        {
            Assert.Throws<ArgumentException>(() => new ReviewDecisionSubject(
                "subject", "", ComparisonAspect.Formula, ReviewDecisionSubjectPolicy.ReviewOnly));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionSubject(
                "subject", "key", ComparisonAspect.Unknown, ReviewDecisionSubjectPolicy.ReviewOnly));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionSubject(
                "subject", "key", ComparisonAspect.CachedResult, ReviewDecisionSubjectPolicy.Selectable));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionSubject(
                "subject", "key", ComparisonAspect.Structure, ReviewDecisionSubjectPolicy.SelectableAndMergeable));
            Assert.Throws<ArgumentException>(() => new ReviewDecisionSubject(
                "subject", "key", ComparisonAspect.Formula, ReviewDecisionSubjectPolicy.Unspecified));
            Assert.Throws<ArgumentOutOfRangeException>(() => new ReviewDecisionSubject(
                "subject", "key", ComparisonAspect.Formula, (ReviewDecisionSubjectPolicy)999));

            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                "project",
                "evidence",
                "mapping",
                "master",
                new[]
                {
                    Subject("first", "same-key", ComparisonAspect.Value, ReviewDecisionSubjectPolicy.ReviewOnly),
                    Subject("second", "same-key", ComparisonAspect.Value, ReviewDecisionSubjectPolicy.ReviewOnly),
                },
                Array.Empty<Observation>()));
        }

        [Fact]
        public void Formula_policy_allows_selecting_a_real_no_formula_observation_state()
        {
            ReviewEvidenceGeneration evidence = Evidence();
            var subject = Subject(
                "formula-subject",
                "logical-cell/formula",
                ComparisonAspect.Formula,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            string masterHash = Hash("=A1");
            string masterObservationId = ComparisonIds.Observation(
                subject.SubjectId,
                "master",
                ObservationExactType.Formula,
                masterHash);
            string candidateHash = Hash("no-formula");
            string candidateObservationId = ComparisonIds.Observation(
                subject.SubjectId,
                "candidate",
                ObservationExactType.Blank,
                candidateHash);
            var catalog = new ReviewDecisionCatalog(
                evidence,
                "master",
                new[] { subject },
                new[]
                {
                    new Observation(
                        masterObservationId,
                        subject.SubjectId,
                        "master",
                        ObservationExactType.Formula,
                        masterHash,
                        formula: "=A1"),
                    new Observation(
                        candidateObservationId,
                        subject.SubjectId,
                        "candidate",
                        ObservationExactType.Blank,
                        candidateHash),
                });
            ReviewDecisionGeneration root = ReviewDecisionEngine.Create(catalog);

            ReviewDecisionGeneration selected = ReviewDecisionEngine.Apply(
                root,
                ReviewDecisionRequest.ForSubject(
                    root.ProjectId,
                    root.BaseGenerationId,
                    root.MappingGenerationId,
                    root.GenerationId,
                    subject.SubjectId,
                    ReviewDecisionStatus.UseCandidate,
                    ReviewerSnapshot.Empty,
                    When,
                    selectedObservationId: candidateObservationId));

            Assert.Equal(
                candidateObservationId,
                selected.GetSubjectState(subject.SubjectId).SelectedObservationId);
        }

        [Fact]
        public void Catalog_identity_binds_formula_value_cache_display_and_date_system_payloads()
        {
            ReviewDecisionCatalog formula = PayloadCatalog(
                ComparisonAspect.Formula,
                FormulaPayload("=A1", "1", "shown"));
            Assert.NotEqual(
                formula.CatalogId,
                PayloadCatalog(
                    ComparisonAspect.Formula,
                    FormulaPayload("=B1", "1", "shown")).CatalogId);
            Assert.NotEqual(
                formula.CatalogId,
                PayloadCatalog(
                    ComparisonAspect.Formula,
                    FormulaPayload("=A1", "2", "shown")).CatalogId);
            Assert.NotEqual(
                formula.CatalogId,
                PayloadCatalog(
                    ComparisonAspect.Formula,
                    FormulaPayload("=A1", "1", "different display")).CatalogId);

            ReviewDecisionCatalog value = PayloadCatalog(
                ComparisonAspect.Value,
                ValuePayload("alpha"));
            Assert.NotEqual(
                value.CatalogId,
                PayloadCatalog(ComparisonAspect.Value, ValuePayload("beta")).CatalogId);

            ReviewDecisionCatalog date1900 = PayloadCatalog(
                ComparisonAspect.Value,
                DatePayload(ExcelDateSystem.Date1900));
            Assert.NotEqual(
                date1900.CatalogId,
                PayloadCatalog(
                    ComparisonAspect.Value,
                    DatePayload(ExcelDateSystem.Date1904)).CatalogId);
        }

        [Fact]
        public void Review_only_and_legacy_unspecified_subjects_block_selection_but_allow_review_statuses()
        {
            var reviewOnly = Subject(
                "cached-subject",
                "logical-cell/cache",
                ComparisonAspect.CachedResult,
                ReviewDecisionSubjectPolicy.ReviewOnly);
            var explicitCatalog = new ReviewDecisionCatalog(
                "project",
                "evidence",
                "mapping",
                "master",
                new[] { reviewOnly },
                Observations(reviewOnly.SubjectId));
            ReviewDecisionGeneration explicitRoot = ReviewDecisionEngine.Create(explicitCatalog);
            Assert.Throws<InvalidOperationException>(() => ReviewDecisionEngine.Apply(
                explicitRoot,
                SelectCandidate(explicitRoot, reviewOnly.SubjectId)));
            ReviewDecisionGeneration reviewed = ReviewDecisionEngine.Apply(
                explicitRoot,
                ReviewDecisionRequest.ForSubject(
                    explicitRoot.ProjectId,
                    explicitRoot.BaseGenerationId,
                    explicitRoot.MappingGenerationId,
                    explicitRoot.GenerationId,
                    reviewOnly.SubjectId,
                    ReviewDecisionStatus.NeedsInvestigation,
                    ReviewerSnapshot.Empty,
                    When));
            Assert.Equal(ReviewDecisionStatus.NeedsInvestigation,
                reviewed.GetSubjectState(reviewOnly.SubjectId).Status);

            var legacyCatalog = new ReviewDecisionCatalog(
                "legacy-project",
                "legacy-evidence",
                "legacy-mapping",
                "master",
                new[] { "legacy-subject" },
                Observations("legacy-subject"));
            Assert.Null(legacyCatalog.EvidenceGeneration);
            Assert.True(legacyCatalog.TryGetSubject("legacy-subject", out ReviewDecisionSubject legacy));
            Assert.Equal(ReviewDecisionSubjectPolicy.Unspecified, legacy.Policy);
            Assert.Equal(ComparisonAspect.Unknown, legacy.Aspect);
            Assert.Equal("legacy-subject", legacy.LogicalSubjectKey);
            ReviewDecisionGeneration legacyRoot = ReviewDecisionEngine.Create(legacyCatalog);
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                ReviewDecisionEngine.Apply(legacyRoot, SelectCandidate(legacyRoot, "legacy-subject")));
            Assert.Contains("Legacy subject metadata", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Subject_and_provenance_json_round_trip_while_legacy_payloads_get_safe_defaults()
        {
            var subject = Subject(
                "subject",
                "logical-key",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            ReviewDecisionSubject restoredSubject = JsonSerializer.Deserialize<ReviewDecisionSubject>(
                JsonSerializer.Serialize(subject))!;
            Assert.Equal(subject.SubjectId, restoredSubject.SubjectId);
            Assert.Equal(subject.LogicalSubjectKey, restoredSubject.LogicalSubjectKey);
            Assert.Equal(subject.Aspect, restoredSubject.Aspect);
            Assert.Equal(subject.Policy, restoredSubject.Policy);

            ReviewDecisionSubject legacySubject = JsonSerializer.Deserialize<ReviewDecisionSubject>(
                "{\"SubjectId\":\"legacy\"}")!;
            Assert.Equal("legacy", legacySubject.LogicalSubjectKey);
            Assert.Equal(ComparisonAspect.Unknown, legacySubject.Aspect);
            Assert.Equal(ReviewDecisionSubjectPolicy.Unspecified, legacySubject.Policy);
            Assert.False(legacySubject.CanSelectObservation);

            var provenance = new ObservationProvenance(
                "revision.xlsx",
                "Data",
                "D17",
                17,
                4,
                "logical-sheet");
            ObservationProvenance restoredProvenance = JsonSerializer.Deserialize<ObservationProvenance>(
                JsonSerializer.Serialize(provenance))!;
            Assert.Equal("logical-sheet", restoredProvenance.LogicalSheetId);
            Assert.Equal("D17", restoredProvenance.Address);

            ObservationProvenance legacyProvenance = JsonSerializer.Deserialize<ObservationProvenance>(
                "{\"SourceLabel\":\"old.xlsx\",\"SheetName\":\"Data\",\"Address\":\"A1\"," +
                "\"PhysicalRow\":1,\"PhysicalColumn\":1}")!;
            Assert.Empty(legacyProvenance.LogicalSheetId);
            Assert.Equal("old.xlsx", legacyProvenance.SourceLabel);
        }

        [Fact]
        public void Catalog_identity_includes_logical_sheet_provenance()
        {
            ReviewDecisionSubject subject = Subject(
                "subject",
                "logical-key",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            ReviewDecisionCatalog left = CatalogWithLogicalSheet(subject, "logical-sheet-a");
            ReviewDecisionCatalog right = CatalogWithLogicalSheet(subject, "logical-sheet-b");

            Assert.NotEqual(left.CatalogId, right.CatalogId);
        }

        [Fact]
        public void Evidence_bound_catalog_rejects_a_master_outside_the_exact_revision_set()
        {
            Assert.Throws<ArgumentException>(() => new ReviewDecisionCatalog(
                Evidence(),
                "unknown-master",
                new[]
                {
                    Subject(
                        "subject",
                        "logical-key",
                        ComparisonAspect.Value,
                        ReviewDecisionSubjectPolicy.SelectableAndMergeable),
                },
                Array.Empty<Observation>()));
        }

        [Fact]
        public void Evidence_bound_catalog_requires_the_complete_subject_revision_matrix()
        {
            ReviewEvidenceGeneration evidence = Evidence();
            ReviewDecisionSubject subject = Subject(
                "subject",
                "logical-key",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                new ReviewDecisionCatalog(
                    evidence,
                    "master",
                    new[] { subject },
                    EvidenceObservations(new[] { subject }).Take(1)));

            Assert.Contains("every subject and evidence revision", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Evidence_bound_catalog_rejects_an_out_of_generation_candidate_before_selection()
        {
            ReviewEvidenceGeneration evidence = Evidence();
            ReviewDecisionSubject subject = Subject(
                "subject",
                "logical-key",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            Observation[] exactMatrix = EvidenceObservations(new[] { subject });
            string outsideHash = Hash("outside-candidate");
            var outsideCandidate = new Observation(
                ComparisonIds.Observation(
                    subject.SubjectId,
                    "emailed-copy",
                    ObservationExactType.String,
                    outsideHash),
                subject.SubjectId,
                "emailed-copy",
                ObservationExactType.String,
                outsideHash,
                value: new CanonicalScalar(ObservationExactType.String, "outside"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                new ReviewDecisionCatalog(
                    evidence,
                    "master",
                    new[] { subject },
                    exactMatrix.Concat(new[] { outsideCandidate })));

            Assert.Contains("exact evidence generation", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Evidence_bound_catalog_rejects_a_forged_observation_id()
        {
            ReviewEvidenceGeneration evidence = Evidence();
            ReviewDecisionSubject subject = Subject(
                "subject",
                "logical-key",
                ComparisonAspect.Value,
                ReviewDecisionSubjectPolicy.SelectableAndMergeable);
            Observation[] observations = EvidenceObservations(new[] { subject });
            Observation candidate = observations.Single(item => item.RevisionId == "candidate");
            observations[Array.IndexOf(observations, candidate)] = new Observation(
                "forged-observation-id",
                candidate.SubjectId,
                candidate.RevisionId,
                candidate.ExactType,
                candidate.ContentHash,
                value: candidate.Value);

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                new ReviewDecisionCatalog(
                    evidence,
                    "master",
                    new[] { subject },
                    observations));

            Assert.Contains("observation ID disagrees", exception.Message, StringComparison.Ordinal);
        }

        private static ReviewDecisionCatalog CatalogWithLogicalSheet(
            ReviewDecisionSubject subject,
            string logicalSheetId) =>
            new ReviewDecisionCatalog(
                "project",
                "evidence",
                "mapping",
                "master",
                new[] { subject },
                new[]
                {
                    new Observation(
                        "observation",
                        subject.SubjectId,
                        "master",
                        ObservationExactType.String,
                        Hash("value"),
                        value: new CanonicalScalar(ObservationExactType.String, "value"),
                        provenance: new ObservationProvenance(
                            "master.xlsx",
                            "Data",
                            "A1",
                            1,
                            1,
                            logicalSheetId)),
                });

        private static ReviewDecisionRequest SelectCandidate(
            ReviewDecisionGeneration generation,
            string subjectId) =>
            ReviewDecisionRequest.ForSubject(
                generation.ProjectId,
                generation.BaseGenerationId,
                generation.MappingGenerationId,
                generation.GenerationId,
                subjectId,
                ReviewDecisionStatus.UseCandidate,
                ReviewerSnapshot.Empty,
                When,
                selectedObservationId: "candidate");

        private static Observation[] Observations(string subjectId) => new[]
        {
            new Observation(
                "master-observation",
                subjectId,
                "master",
                ObservationExactType.String,
                Hash("master"),
                value: new CanonicalScalar(ObservationExactType.String, "master")),
            new Observation(
                "candidate",
                subjectId,
                "candidate",
                ObservationExactType.String,
                Hash("candidate"),
                value: new CanonicalScalar(ObservationExactType.String, "candidate")),
        };

        private static Observation[] EvidenceObservations(
            IEnumerable<ReviewDecisionSubject> subjects) =>
            subjects.SelectMany(subject => new[]
            {
                EvidenceObservation(subject, "candidate"),
                EvidenceObservation(subject, "master"),
            }).ToArray();

        private static Observation EvidenceObservation(
            ReviewDecisionSubject subject,
            string revisionId)
        {
            string contentHash = Hash(subject.SubjectId + "-" + revisionId);
            if (subject.Aspect == ComparisonAspect.Formula)
            {
                string observationId = ComparisonIds.Observation(
                    subject.SubjectId,
                    revisionId,
                    ObservationExactType.Formula,
                    contentHash);
                return new Observation(
                    observationId,
                    subject.SubjectId,
                    revisionId,
                    ObservationExactType.Formula,
                    contentHash,
                    formula: "=" + revisionId);
            }
            string scalarObservationId = ComparisonIds.Observation(
                subject.SubjectId,
                revisionId,
                ObservationExactType.String,
                contentHash);
            return new Observation(
                scalarObservationId,
                subject.SubjectId,
                revisionId,
                ObservationExactType.String,
                contentHash,
                value: new CanonicalScalar(ObservationExactType.String, revisionId));
        }

        private static ReviewDecisionCatalog PayloadCatalog(
            ComparisonAspect aspect,
            Observation observation) =>
            new ReviewDecisionCatalog(
                "payload-project",
                "payload-generation",
                "payload-mapping",
                "master",
                new[]
                {
                    Subject(
                        observation.SubjectId,
                        "payload-logical-key",
                        aspect,
                        ReviewDecisionSubjectPolicy.SelectableAndMergeable),
                },
                new[] { observation });

        private static Observation FormulaPayload(
            string formula,
            string cachedResult,
            string displayText) =>
            new Observation(
                "payload-observation",
                "payload-subject",
                "master",
                ObservationExactType.Formula,
                Hash("deliberately-identical-content-identity"),
                formula: formula,
                cachedResult: new CanonicalScalar(ObservationExactType.Number, cachedResult),
                cachedResultTrust: CachedResultTrust.Trusted,
                displayText: displayText);

        private static Observation ValuePayload(string value) =>
            new Observation(
                "payload-observation",
                "payload-subject",
                "master",
                ObservationExactType.String,
                Hash("deliberately-identical-content-identity"),
                value: new CanonicalScalar(ObservationExactType.String, value));

        private static Observation DatePayload(ExcelDateSystem dateSystem) =>
            new Observation(
                "payload-observation",
                "payload-subject",
                "master",
                ObservationExactType.DateSerial,
                Hash("deliberately-identical-content-identity"),
                value: new CanonicalScalar(ObservationExactType.DateSerial, "1", dateSystem));

        private static ReviewDecisionSubject Subject(
            string subjectId,
            string logicalSubjectKey,
            ComparisonAspect aspect,
            ReviewDecisionSubjectPolicy policy) =>
            new ReviewDecisionSubject(subjectId, logicalSubjectKey, aspect, policy);

        private static ReviewEvidenceGeneration Evidence() =>
            new ReviewEvidenceGeneration(
                "project",
                "effective-mapping",
                new[]
                {
                    new ReviewEvidenceRevisionIdentity("candidate", 0, Hash("candidate-file")),
                    new ReviewEvidenceRevisionIdentity("master", 1, Hash("master-file")),
                });

        private static string Hash(string value) =>
            StableId.ForBytes(Encoding.UTF8.GetBytes(value));
    }
}
