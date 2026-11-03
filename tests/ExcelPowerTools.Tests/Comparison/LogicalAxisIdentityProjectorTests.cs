using System;
using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Identifiers;
using ExcelPowerTools.Core.Comparison.Projection;
using Xunit;

namespace ExcelPowerTools.Tests.Comparison
{
    public sealed class LogicalAxisIdentityProjectorTests
    {
        [Fact]
        public void Root_assignments_are_deterministic_namespaced_and_immutable()
        {
            LogicalAxisIdentityProjection first = LogicalAxisIdentityProjector.CreateRoot(
                "project-1", "sheet-input", "revision-root", LogicalAxisKind.Row, 3);
            LogicalAxisIdentityProjection repeated = LogicalAxisIdentityProjector.CreateRoot(
                "project-1", "sheet-input", "revision-root", LogicalAxisKind.Row, 3);
            LogicalAxisIdentityProjection otherAxis = LogicalAxisIdentityProjector.CreateRoot(
                "project-1", "sheet-input", "revision-root", LogicalAxisKind.Column, 3);
            LogicalAxisIdentityProjection otherSheet = LogicalAxisIdentityProjector.CreateRoot(
                "project-1", "sheet-output", "revision-root", LogicalAxisKind.Row, 3);

            Assert.Equal(first.ProjectionId, repeated.ProjectionId);
            Assert.Equal(first.IdentityByPhysicalIndex, repeated.IdentityByPhysicalIndex);
            Assert.NotEqual(first.IdentityByPhysicalIndex[0], otherAxis.IdentityByPhysicalIndex[0]);
            Assert.NotEqual(first.IdentityByPhysicalIndex[0], otherSheet.IdentityByPhysicalIndex[0]);
            Assert.Equal(AlignmentConfidence.Confirmed, first.Confidence);
            Assert.True(first.IsFullyResolved);
            Assert.Empty(first.StructuralEvents);
            Assert.All(first.AssignmentsByPhysicalIndex.Values, assignment =>
            {
                Assert.Equal(LogicalAxisIdentitySource.Root, assignment.Source);
                Assert.Equal(AlignmentConfidence.Confirmed, assignment.Confidence);
                Assert.True(StableId.IsSha256(assignment.AssignmentId));
                Assert.True(StableId.IsSha256(assignment.LogicalIdentityId));
            });
            Assert.True(StableId.IsSha256(first.ProjectionId));

            IDictionary<int, string> mutableView = Assert.IsAssignableFrom<IDictionary<int, string>>(
                first.IdentityByPhysicalIndex);
            Assert.Throws<NotSupportedException>(() => mutableView.Add(9, "fabricated"));
        }

        [Fact]
        public void Resolved_matches_inherit_while_replacement_emits_deletion_and_insertion()
        {
            AlignmentTable before = Table(
                new[] { "A" },
                new[] { "B" },
                new[] { "C" });
            AlignmentTable after = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "C" });
            AxisAlignmentResult alignment = AlignRowsByConfirmedKey(before, after);
            LogicalAxisIdentityProjection root = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, before.RowCount);

            LogicalAxisIdentityProjection child = LogicalAxisIdentityProjector.Propagate(
                root, "r2", after.RowCount, alignment);
            LogicalAxisIdentityProjection repeated = LogicalAxisIdentityProjector.Propagate(
                root, "r2", after.RowCount, alignment);

            Assert.Equal(root.IdentityByPhysicalIndex[0], child.IdentityByPhysicalIndex[0]);
            Assert.Equal(root.IdentityByPhysicalIndex[2], child.IdentityByPhysicalIndex[2]);
            Assert.DoesNotContain(root.IdentityByPhysicalIndex.Values,
                identity => identity == child.IdentityByPhysicalIndex[1]);
            Assert.Equal(LogicalAxisIdentitySource.Inserted,
                child.AssignmentsByPhysicalIndex[1].Source);

            LogicalAxisStructuralEvent insertion = Assert.Single(child.StructuralEvents,
                item => item.Kind == LogicalAxisStructuralEventKind.Insertion);
            LogicalAxisStructuralEvent deletion = Assert.Single(child.StructuralEvents,
                item => item.Kind == LogicalAxisStructuralEventKind.Deletion);
            Assert.Equal(child.IdentityByPhysicalIndex[1], insertion.LogicalIdentityId);
            Assert.Equal(root.IdentityByPhysicalIndex[1], deletion.LogicalIdentityId);
            Assert.Equal(1, insertion.AfterIndex);
            Assert.Equal(1, deletion.BeforeIndex);
            Assert.All(child.StructuralEvents, item => Assert.True(StableId.IsSha256(item.EventId)));
            Assert.Empty(child.Ambiguities);
            Assert.True(child.IsFullyResolved);
            Assert.Equal(child.ProjectionId, repeated.ProjectionId);
            Assert.Equal(
                child.StructuralEvents.Select(item => item.EventId),
                repeated.StructuralEvents.Select(item => item.EventId));
        }

        [Fact]
        public void Reordered_matches_keep_identity_and_emit_only_reorder_events()
        {
            AlignmentTable before = Table(new[] { "A" }, new[] { "B" }, new[] { "C" });
            AlignmentTable after = Table(new[] { "B" }, new[] { "A" }, new[] { "C" });
            AxisAlignmentResult alignment = TabularAlignmentEngine.AlignRows(before, after);
            LogicalAxisIdentityProjection root = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, before.RowCount);

            LogicalAxisIdentityProjection child = LogicalAxisIdentityProjector.Propagate(
                root, "r2", after.RowCount, alignment);

            Assert.Equal(root.IdentityByPhysicalIndex[1], child.IdentityByPhysicalIndex[0]);
            Assert.Equal(root.IdentityByPhysicalIndex[0], child.IdentityByPhysicalIndex[1]);
            Assert.Equal(root.IdentityByPhysicalIndex[2], child.IdentityByPhysicalIndex[2]);
            Assert.Equal(2, child.StructuralEvents.Count);
            Assert.All(child.StructuralEvents,
                item => Assert.Equal(LogicalAxisStructuralEventKind.Reorder, item.Kind));
            Assert.DoesNotContain(child.AssignmentsByPhysicalIndex.Values,
                item => item.Source == LogicalAxisIdentitySource.Inserted);
        }

        [Fact]
        public void Ambiguous_blocks_preserve_candidates_but_never_assign_child_identities()
        {
            AlignmentTable before = Table(
                new[] { "A" },
                new[] { "X" },
                new[] { "X" },
                new[] { "Y" },
                new[] { "B" });
            AlignmentTable after = Table(
                new[] { "A" },
                new[] { "Y" },
                new[] { "X" },
                new[] { "X" },
                new[] { "B" });
            AxisAlignmentResult alignment = TabularAlignmentEngine.AlignRows(before, after);
            LogicalAxisIdentityProjection root = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, before.RowCount);

            LogicalAxisIdentityProjection child = LogicalAxisIdentityProjector.Propagate(
                root, "r2", after.RowCount, alignment);

            LogicalAxisIdentityAmbiguity ambiguity = Assert.Single(child.Ambiguities);
            Assert.Equal(LogicalAxisAmbiguityKind.AlignmentBlock, ambiguity.Kind);
            Assert.Equal(new[] { 1, 2 }, ambiguity.BeforeIndices);
            Assert.Equal(new[] { 2, 3 }, ambiguity.AfterIndices);
            Assert.Equal(root.IdentityByPhysicalIndex[1], ambiguity.KnownParentIdentities[1]);
            Assert.Equal(root.IdentityByPhysicalIndex[2], ambiguity.KnownParentIdentities[2]);
            Assert.Equal(new[] { 2, 3 }, child.UnmappedPhysicalIndices);
            Assert.All(ambiguity.AfterIndices,
                index => Assert.False(child.TryGetLogicalIdentity(index, out _)));
            Assert.Equal(AlignmentConfidence.None, ambiguity.Confidence);
            Assert.Equal(AlignmentConfidence.None, child.Confidence);
            Assert.False(child.IsFullyResolved);
            Assert.True(StableId.IsSha256(ambiguity.AmbiguityId));
            Assert.Contains(ambiguity.Reasons,
                reason => reason.Code == "identity.ambiguous-block-no-propagation");
        }

        [Fact]
        public void Later_deletion_of_unresolved_indices_does_not_invent_their_identities()
        {
            AlignmentTable rootTable = Table(
                new[] { "A" }, new[] { "X" }, new[] { "X" }, new[] { "Y" }, new[] { "B" });
            AlignmentTable ambiguousTable = Table(
                new[] { "A" }, new[] { "Y" }, new[] { "X" }, new[] { "X" }, new[] { "B" });
            AlignmentTable finalTable = Table(new[] { "A" }, new[] { "Y" }, new[] { "B" });
            LogicalAxisIdentityProjection root = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, rootTable.RowCount);
            LogicalAxisIdentityProjection ambiguous = LogicalAxisIdentityProjector.Propagate(
                root,
                "r2",
                ambiguousTable.RowCount,
                TabularAlignmentEngine.AlignRows(rootTable, ambiguousTable));

            LogicalAxisIdentityProjection final = LogicalAxisIdentityProjector.Propagate(
                ambiguous,
                "r3",
                finalTable.RowCount,
                TabularAlignmentEngine.AlignRows(ambiguousTable, finalTable));

            Assert.Equal(ambiguous.IdentityByPhysicalIndex[0], final.IdentityByPhysicalIndex[0]);
            Assert.Equal(ambiguous.IdentityByPhysicalIndex[1], final.IdentityByPhysicalIndex[1]);
            Assert.Equal(ambiguous.IdentityByPhysicalIndex[4], final.IdentityByPhysicalIndex[2]);
            LogicalAxisStructuralEvent[] unresolvedDeletions = final.StructuralEvents
                .Where(item => item.Kind == LogicalAxisStructuralEventKind.Deletion &&
                               item.LogicalIdentityId.Length == 0)
                .ToArray();
            Assert.Equal(2, unresolvedDeletions.Length);
            Assert.All(unresolvedDeletions,
                item => Assert.Equal(AlignmentConfidence.None, item.Confidence));
            Assert.Equal(2, final.Ambiguities.Count(
                item => item.Kind == LogicalAxisAmbiguityKind.MissingParentIdentity));
            Assert.DoesNotContain(final.IdentityByPhysicalIndex.Values,
                identity => identity == root.IdentityByPhysicalIndex[1] ||
                            identity == root.IdentityByPhysicalIndex[2]);
        }

        [Fact]
        public void Inserted_identity_is_inherited_by_later_revisions()
        {
            AlignmentTable rootTable = Table(new[] { "A" }, new[] { "B" });
            AlignmentTable childTable = Table(new[] { "A" }, new[] { "X" }, new[] { "B" });
            LogicalAxisIdentityProjection root = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, rootTable.RowCount);
            LogicalAxisIdentityProjection child = LogicalAxisIdentityProjector.Propagate(
                root,
                "r2",
                childTable.RowCount,
                TabularAlignmentEngine.AlignRows(rootTable, childTable));

            LogicalAxisIdentityProjection grandchild = LogicalAxisIdentityProjector.Propagate(
                child,
                "r3",
                childTable.RowCount,
                TabularAlignmentEngine.AlignRows(childTable, childTable));

            Assert.Equal(child.IdentityByPhysicalIndex[1], grandchild.IdentityByPhysicalIndex[1]);
            LogicalAxisIdentityAssignment inherited = grandchild.AssignmentsByPhysicalIndex[1];
            Assert.Equal(LogicalAxisIdentitySource.Inherited, inherited.Source);
            Assert.Equal("r2", inherited.OriginRevisionId);
            Assert.Equal(1, inherited.OriginPhysicalIndex);
            Assert.Equal(1, inherited.ParentPhysicalIndex);
        }

        [Fact]
        public void Column_alignment_uses_the_same_contract_without_row_identity_collisions()
        {
            AlignmentTable before = Table(
                new[] { "ID", "Name", "Amount" },
                new[] { "A", "Alpha", "10" });
            AlignmentTable after = Table(
                new[] { "ID", "Notes", "Name", "Amount" },
                new[] { "A", "new", "Alpha", "10" });
            AxisAlignmentResult alignment = TabularAlignmentEngine.AlignColumns(
                before, after, ConfirmedHeader(before), ConfirmedHeader(after));
            LogicalAxisIdentityProjection columns = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Column, before.ColumnCount);
            LogicalAxisIdentityProjection rows = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, before.ColumnCount);

            LogicalAxisIdentityProjection child = LogicalAxisIdentityProjector.Propagate(
                columns, "r2", after.ColumnCount, alignment);

            Assert.Equal(columns.IdentityByPhysicalIndex[0], child.IdentityByPhysicalIndex[0]);
            Assert.Equal(columns.IdentityByPhysicalIndex[1], child.IdentityByPhysicalIndex[2]);
            Assert.Equal(columns.IdentityByPhysicalIndex[2], child.IdentityByPhysicalIndex[3]);
            Assert.Equal(LogicalAxisIdentitySource.Inserted, child.AssignmentsByPhysicalIndex[1].Source);
            Assert.NotEqual(rows.IdentityByPhysicalIndex[0], columns.IdentityByPhysicalIndex[0]);
        }

        [Fact]
        public void Uncovered_child_indices_remain_explicit_and_unmapped()
        {
            AlignmentTable single = Table(new[] { "A" });
            LogicalAxisIdentityProjection root = LogicalAxisIdentityProjector.CreateRoot(
                "project", "sheet", "r1", LogicalAxisKind.Row, 1);
            AxisAlignmentResult alignment = TabularAlignmentEngine.AlignRows(single, single);

            LogicalAxisIdentityProjection child = LogicalAxisIdentityProjector.Propagate(
                root, "r2", 2, alignment);

            Assert.Equal(new[] { 1 }, child.UnmappedPhysicalIndices);
            LogicalAxisIdentityAmbiguity gap = Assert.Single(child.Ambiguities);
            Assert.Equal(LogicalAxisAmbiguityKind.UncoveredPhysicalIndex, gap.Kind);
            Assert.Empty(gap.BeforeIndices);
            Assert.Equal(new[] { 1 }, gap.AfterIndices);
            Assert.False(child.TryGetLogicalIdentity(1, out _));
        }

        private static AxisAlignmentResult AlignRowsByConfirmedKey(
            AlignmentTable before,
            AlignmentTable after)
        {
            HeaderDetectionResult beforeHeader = HeaderDetector.Detect(before);
            HeaderDetectionResult afterHeader = HeaderDetector.Detect(after);
            KeyInferenceResult key = KeyColumnInference.Infer(
                before,
                after,
                beforeHeader,
                afterHeader,
                new KeyInferenceOptions
                {
                    ConfirmedBeforeColumns = new[] { 0 },
                    ConfirmedAfterColumns = new[] { 0 },
                });
            return TabularAlignmentEngine.AlignRows(before, after, beforeHeader, afterHeader, key);
        }

        private static HeaderDetectionResult ConfirmedHeader(AlignmentTable table) =>
            HeaderDetector.Detect(
                table,
                options: new HeaderDetectionOptions { ConfirmedHeaderRowIndex = 0 });

        private static AlignmentTable Table(params string[][] rows) => new AlignmentTable(rows);
    }
}
