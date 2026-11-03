using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ExcelPowerTools.Core.SpreadsheetReview;
using Xunit;

namespace ExcelPowerTools.Tests.SpreadsheetReview;

public class ReviewKernelTests
{
    private static CellObservation Value(string text, bool struck = false, string? fill = null) => Observe(CellContent.Value(ReviewScalar.Text(text)), struck, fill);
    private static CellObservation Observe(CellContent content, bool struck = false, string? fill = null) => CellObservation.Observed(content,
        ChannelEvidence<CellInterior>.Observed(fill == null ? CellInterior.NoFill : new CellInterior("solid", fill, null)),
        ChannelEvidence<StrikeProfile>.Observed(new StrikeProfile(content.ToString(), struck)));
    private static ReviewSheet Sheet(string id, params string[][] rows) => new(id, "Data", rows.SelectMany((row, r) =>
        row.Select((v, c) => new ReviewCell(r + 1, c + 1, Value(v)))), (r, c) => Observe(CellContent.Value(ReviewScalar.Blank)));
    private static ReviewRowAlignment Align(string[][] a, string[][] b, int cap = 1) => new OrderedRowAlignment().Align(Sheet("A", a), Sheet("B", b), ResolvedReviewScope.SameColumns(cap));

    [Theory]
    [InlineData("1", "1.0")]
    [InlineData("1E3", "1000")]
    [InlineData("-0", "0")]
    public void Numeric_encoding_is_retained_but_not_a_value_change(string a, string b)
    {
        var x = ReviewScalar.Number(a); var y = ReviewScalar.Number(b);
        Assert.Equal(x, y); Assert.Equal(x.GetHashCode(), y.GetHashCode()); Assert.Equal(a, x.OriginalEncoding);
        Assert.Empty(CellHistoryComparison.Compare(new[] { Observe(CellContent.Value(x)), Observe(CellContent.Value(y)) }).Transitions);
    }

    [Theory]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("1,000")]
    [InlineData("1e999")]
    public void Invalid_numeric_evidence_is_not_coerced(string value) => Assert.Throws<FormatException>(() => ReviewScalar.Number(value));

    [Fact]
    public void Typed_values_remain_distinct()
    {
        Assert.NotEqual(ReviewScalar.Number("1"), ReviewScalar.Text("1"));
        Assert.NotEqual(ReviewScalar.Boolean(true), ReviewScalar.Text("TRUE"));
        Assert.NotEqual(ReviewScalar.Error("#N/A"), ReviewScalar.Text("#N/A"));
    }

    [Fact]
    public void Formula_returning_empty_is_still_content_and_missing_cache_is_not_blank()
    {
        var a = Observe(CellContent.Expression("=\"\"", ChannelEvidence<ReviewScalar>.Observed(ReviewScalar.Text(""))));
        var b = Observe(CellContent.Expression("=\"\""));
        var result = CellHistoryComparison.Compare(new[] { a, b });
        Assert.True(result.IsRelevant); Assert.True(result.HasIncompleteEvidence);
        Assert.Equal(ReviewCellChange.IncompleteEvidence, Assert.Single(result.Transitions).Changes);
        Assert.Contains(result.Transitions[0].Reasons, r => r.Contains("saved formula result"));
    }

    [Fact]
    public void Same_formula_different_saved_result_is_secondary()
    {
        var a = Observe(CellContent.Expression("A1+1", ChannelEvidence<ReviewScalar>.Observed(ReviewScalar.Number("3"))));
        var b = Observe(CellContent.Expression("A1+1", ChannelEvidence<ReviewScalar>.Observed(ReviewScalar.Number("4"))));
        Assert.Equal(ReviewCellChange.SavedResult, Assert.Single(CellHistoryComparison.Compare(new[] { a, b }).Transitions).Changes);
    }

    [Fact]
    public void Formatting_only_on_never_populated_cells_is_silent()
    {
        var result = CellHistoryComparison.Compare(new[] { Value("", fill: "FFFF0000"), Value("", true, "FF00FF00") });
        Assert.False(result.IsRelevant); Assert.Empty(result.Transitions);
    }

    [Fact]
    public void Blank_formatting_survives_the_ever_content_union()
    {
        var result = CellHistoryComparison.Compare(new[] { Value("data"), Value("", fill: "FFFF0000"), Value("", fill: "FF00FF00") });
        Assert.True(result.IsRelevant); Assert.Equal(2, result.Transitions.Count);
        Assert.Equal(ReviewCellChange.Interior, result.Transitions[1].Changes);
    }

    [Fact]
    public void Intermediate_only_changes_are_visible_when_endpoints_agree()
    {
        var result = CellHistoryComparison.Compare(new[] { Value("first"), Value("middle", true), Value("first") });
        Assert.True(result.HasIntermediateOnlyChanges); Assert.Equal(2, result.Transitions.Count);
    }

    [Theory]
    [InlineData(CellObservationState.Unresolved)]
    [InlineData(CellObservationState.Unavailable)]
    [InlineData(CellObservationState.Failed)]
    [InlineData(CellObservationState.Excluded)]
    public void Missing_evidence_never_becomes_an_empty_cell(CellObservationState state)
    {
        var result = CellHistoryComparison.Compare(new[] { Value(""), CellObservation.Missing(state, "Reason") });
        Assert.False(result.IsRelevant); Assert.True(result.HasIncompleteEvidence);
        Assert.Equal(ReviewCellChange.IncompleteEvidence, Assert.Single(result.Transitions).Changes);
    }

    [Fact]
    public void Confirmed_absence_is_different_from_failed_read()
    {
        var result = CellHistoryComparison.Compare(new[] { Value("data"), CellObservation.Missing(CellObservationState.AbsentRow, "Deleted row") });
        Assert.Equal(ReviewCellChange.Content | ReviewCellChange.Presence, Assert.Single(result.Transitions).Changes);
        Assert.False(result.HasIncompleteEvidence);
    }

    [Theory]
    [InlineData("a", "A")]
    [InlineData("a b", "a  b")]
    [InlineData("a\tb", "a b")]
    [InlineData("a\r\nb", "a\nb")]
    [InlineData("=A1", "= A1")]
    [InlineData("", "added")]
    [InlineData("removed", "")]
    [InlineData("A😀B", "A😁B")]
    [InlineData("é", "é")]
    public void Exact_character_diff_round_trips_both_sides(string a, string b)
    {
        var spans = ExactTextDifference.Compare(a, b);
        Assert.Equal(a, string.Concat(spans.Where(s => s.Kind != ReviewTextSpanKind.Added).Select(s => s.Text)));
        Assert.Equal(b, string.Concat(spans.Where(s => s.Kind != ReviewTextSpanKind.Removed).Select(s => s.Text)));
        Assert.Contains(spans, s => s.Kind != ReviewTextSpanKind.Equal);
        Assert.All(spans, s => { Assert.False(char.IsLowSurrogate(s.Text[0])); Assert.False(char.IsHighSurrogate(s.Text[^1])); });
    }

    [Fact]
    public void Rich_text_storage_partitions_do_not_change_effective_strike()
    {
        var a = new StrikeProfile("abcd", false, new[] { new StrikeRun(0, 4, true) });
        var b = new StrikeProfile("abcd", true);
        var c = new StrikeProfile("abcd", false, new[] { new StrikeRun(0, 2, true), new StrikeRun(2, 2, true) });
        Assert.Equal(a, b); Assert.Equal(a.GetHashCode(), b.GetHashCode()); Assert.Equal(b, c);
        Assert.NotEqual(new StrikeProfile("", false), new StrikeProfile("", true));
    }

    [Fact]
    public void Inserting_text_before_a_struck_character_is_not_a_strike_change()
    {
        CellObservation Rich(string text, int start) => CellObservation.Observed(CellContent.Value(ReviewScalar.Text(text)),
            ChannelEvidence<CellInterior>.Observed(CellInterior.NoFill),
            ChannelEvidence<StrikeProfile>.Observed(new StrikeProfile(text, false, new[] { new StrikeRun(start, 1, true) })));
        var result = CellHistoryComparison.Compare(new[] { Rich("abc", 1), Rich("aXbc", 2) });
        Assert.Equal(ReviewCellChange.Content, Assert.Single(result.Transitions).Changes);
    }

    [Fact]
    public void Replaced_rich_text_with_different_effective_strike_is_reported()
    {
        var a = CellObservation.Observed(CellContent.Value(ReviewScalar.Text("abc")), ChannelEvidence<CellInterior>.Observed(CellInterior.NoFill),
            ChannelEvidence<StrikeProfile>.Observed(new StrikeProfile("abc", false, new[] { new StrikeRun(0, 3, true) })));
        var result = CellHistoryComparison.Compare(new[] { a, Value("xyz") });
        Assert.Equal(ReviewCellChange.Content | ReviewCellChange.Strikethrough, Assert.Single(result.Transitions).Changes);
    }

    [Fact]
    public void An_inserted_row_does_not_cascade_into_changed_rows()
    {
        var result = Align(new[] { new[] { "head" }, new[] { "alpha" }, new[] { "beta" } },
            new[] { "head" }, new[] { "insert" }, new[] { "alpha" }, new[] { "beta" });
        Assert.Equal(new[] { (1, 1), (2, 3), (3, 4) }, result.Matches.Select(m => (m.BeforeRow, m.AfterRow)));
        var region = Assert.Single(result.Regions); Assert.Equal(0, region.BeforeCount); Assert.Equal(1, region.AfterCount);
        AssertCoverage(result);
    }

    private static ReviewRowAlignment Align(string[][] a, params string[][] b) => Align(a, b, 1);

    [Fact]
    public void Duplicate_rows_are_not_arbitrarily_paired()
    {
        var result = Align(new[] { new[] { "top" }, new[] { "X" }, new[] { "X" }, new[] { "end" } },
            new[] { new[] { "top" }, new[] { "X" }, new[] { "end" } }, 1);
        Assert.DoesNotContain(result.Matches, m => m.BeforeRow == 2 || m.BeforeRow == 3);
        var region = Assert.Single(result.Regions); Assert.Equal(2, region.BeforeCount); Assert.Equal(1, region.AfterCount);
        Assert.Equal(2, region.Suggestions.Count); AssertCoverage(result);
    }

    [Fact]
    public void Partial_matches_remain_suggestions_and_every_row_is_retained()
    {
        var result = Align(new[] { new[] { "a", "original" }, new[] { "b", "steady" } },
            new[] { new[] { "a", "edited" }, new[] { "b", "steady" } }, 2);
        Assert.Single(result.Matches); Assert.Single(Assert.Single(result.Regions).Suggestions); AssertCoverage(result);
    }

    [Fact]
    public void Unofficial_columns_cannot_change_alignment()
    {
        var a = new[] { new[] { "a", "one" }, new[] { "b", "two" } };
        var b = new[] { new[] { "a", "two" }, new[] { "b", "one" } };
        var result = Align(a, b, 1);
        Assert.Equal(2, result.Matches.Count); Assert.Empty(result.Regions); Assert.Empty(result.Issues);
    }

    [Fact]
    public void Official_Z_is_inclusive_and_AA_cannot_expand_or_steer_alignment()
    {
        var a = new ReviewSheet("A", "R1.S1.C1", new[] {
            new ReviewCell(1, 26, Value("first")), new ReviewCell(2, 26, Value("second")),
            new ReviewCell(1, 27, Value("notes-a")), new ReviewCell(2, 27, Value("notes-b")),
            new ReviewCell(100000, 27, Value("unofficial tail")) });
        var b = new ReviewSheet("B", "R1.S1.C1", new[] {
            new ReviewCell(1, 26, Value("first")), new ReviewCell(2, 26, Value("inserted")), new ReviewCell(3, 26, Value("second")),
            new ReviewCell(1, 27, Value("notes-b")), new ReviewCell(3, 27, Value("notes-a")),
            new ReviewCell(200000, 27, Value("different unofficial tail")) });

        var result = new OrderedRowAlignment().Align(a, b, ResolvedReviewScope.SameColumns(26));

        Assert.Equal(2, result.BeforeRows); Assert.Equal(3, result.AfterRows);
        Assert.Equal(new[] { (1, 1), (2, 3) }, result.Matches.Select(m => (m.BeforeRow, m.AfterRow)));
        var insertion = Assert.Single(result.Regions);
        Assert.Equal(0, insertion.BeforeCount); Assert.Equal(2, insertion.AfterStart); Assert.Equal(1, insertion.AfterCount);
        Assert.Empty(result.Issues); AssertCoverage(result);
    }

    [Fact]
    public void Explicit_column_mapping_handles_inserted_official_column()
    {
        var scope = new ResolvedReviewScope(2, 3, new[] { new ReviewColumnPair(1, 1), new ReviewColumnPair(null, 2), new ReviewColumnPair(2, 3) });
        var result = new OrderedRowAlignment().Align(Sheet("A", new[] { "id", "value" }), Sheet("B", new[] { "id", "new", "value" }), scope);
        Assert.Single(result.Matches); Assert.Empty(result.Regions);
    }

    [Fact]
    public void Invalid_or_incomplete_scope_cannot_enter_alignment()
    {
        Assert.Throws<ArgumentException>(() => new ResolvedReviewScope(2, 2, new[] { new ReviewColumnPair(1, 1) }));
        Assert.Throws<ArgumentException>(() => new ResolvedReviewScope(1, 2, new[] { new ReviewColumnPair(1, 1), new ReviewColumnPair(1, 2) }));
    }

    [Fact]
    public void Partial_cell_shift_requires_review_not_whole_row_inference()
    {
        var a = Sheet("A", new[] { "a1", "b1" }, new[] { "a2", "b2" }, new[] { "a3", "b3" }, new[] { "a4", "b4" });
        var b = Sheet("B", new[] { "a1", "new" }, new[] { "a2", "b1" }, new[] { "a3", "b2" }, new[] { "a4", "b3" }, new[] { "", "b4" });
        var result = new OrderedRowAlignment().Align(a, b, ResolvedReviewScope.SameColumns(2));
        Assert.Empty(result.Matches); Assert.Contains(result.Issues, i => i.Code == "SR.ALIGN.PARTIAL_SHIFT"); AssertCoverage(result);
        var fixedResult = new OrderedRowAlignment().Align(a, b, ResolvedReviewScope.SameColumns(2), explicitFixedPosition: true);
        Assert.Equal(4, fixedResult.Matches.Count); Assert.All(fixedResult.Matches, m => Assert.Equal(RowMatchBasis.ExplicitPosition, m.Basis));
    }

    [Fact]
    public void Formatting_tails_and_excluded_data_do_not_expand_the_row_envelope()
    {
        var sheet = new ReviewSheet("A", "Data", new[] { new ReviewCell(1, 1, Value("data")),
            new ReviewCell(1048576, 1, Value("", true)), new ReviewCell(100000, 2, Value("unofficial")) });
        Assert.Equal(1, sheet.LastContentRow(1));
    }

    [Fact]
    public void Different_offsets_in_disjoint_sheet_regions_are_not_partial_shift_conflicts()
    {
        var a = Sheet("A", new[] { "a1", "" }, new[] { "a2", "" }, new[] { "a3", "" }, new[] { "a4", "" },
            new[] { "", "b1" }, new[] { "", "b2" }, new[] { "", "b3" }, new[] { "", "b4" });
        var b = Sheet("B", new[] { "a1", "" }, new[] { "a2", "" }, new[] { "a3", "" }, new[] { "a4", "" },
            new[] { "inserted", "" }, new[] { "", "b1" }, new[] { "", "b2" }, new[] { "", "b3" }, new[] { "", "b4" });
        var result = new OrderedRowAlignment().Align(a, b, ResolvedReviewScope.SameColumns(2));
        Assert.Equal(8, result.Matches.Count); Assert.Empty(result.Issues); AssertCoverage(result);
    }

    [Fact]
    public void Known_formula_storage_does_not_hide_independent_constant_anchors()
    {
        var formula = CellContent.UnexpandedExpression("shared index 2", "Formula expansion pending", ChannelEvidence<ReviewScalar>.Observed(ReviewScalar.Number("7")));
        var a = new ReviewSheet("A", "Data", new[] { new ReviewCell(1, 1, Value("unique")), new ReviewCell(1, 2, Observe(formula)) });
        var b = new ReviewSheet("B", "Data", new[] { new ReviewCell(2, 1, Value("unique")), new ReviewCell(2, 2, Observe(formula)) });
        var result = new OrderedRowAlignment().Align(a, b, ResolvedReviewScope.SameColumns(2));
        Assert.Equal((1, 2), (Assert.Single(result.Matches).BeforeRow, result.Matches[0].AfterRow));
        Assert.True(CellHistoryComparison.Compare(new[] { a.Cell(1, 2), b.Cell(2, 2) }).HasIncompleteEvidence);
    }

    [Fact]
    public void Budget_exhaustion_keeps_raw_regions_and_actionable_issue()
    {
        var result = new OrderedRowAlignment().Align(Sheet("A", new[] { "duplicate" }, new[] { "duplicate" }),
            Sheet("B", new[] { "duplicate" }), ResolvedReviewScope.SameColumns(1), candidateBudget: 0);
        AssertCoverage(result); var issue = Assert.Single(result.Issues);
        Assert.Equal("SR.ALIGN.CANDIDATE_BUDGET", issue.Code); Assert.NotEmpty(issue.IssueId);
        Assert.EndsWith("OrderedRowAlignment.cs", issue.SourceFile); Assert.True(issue.SourceLine > 0);
        Assert.Contains(issue.IssueId, issue.ToString()); Assert.Contains("beforeArtifact: A", issue.ToString());
    }

    [Fact]
    public void Cancelled_alignment_does_not_return_a_partial_success()
    {
        Assert.Throws<OperationCanceledException>(() => new OrderedRowAlignment().Align(Sheet("A", new[] { "a" }),
            Sheet("B", new[] { "b" }), ResolvedReviewScope.SameColumns(1), cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void Conflicting_revision_chain_is_rejected_atomically()
    {
        var ledger = new RowIdentityLedger(); var a = new RowEvidenceIdentity("A", "Data", 1); var b = new RowEvidenceIdentity("B", "Data", 2);
        var c = new RowEvidenceIdentity("C", "Data", 3); var otherA = new RowEvidenceIdentity("A", "Data", 5);
        Assert.True(ledger.TryAccept(a, b, out _)); Assert.True(ledger.TryAccept(b, c, out _));
        Assert.False(ledger.HasDirectEvidence(a, c));
        Assert.False(ledger.TryAccept(c, otherA, out var issue)); Assert.Equal("SR.IDENTITY.NON_TRANSITIVE", issue.Code);
        Assert.Equal(3, Assert.Single(ledger.Groups).Count); Assert.False(ledger.HasDirectEvidence(c, otherA));
    }

    [Fact]
    public void Repeated_file_names_and_version_labels_are_not_identity()
    {
        var ledger = new RowIdentityLedger();
        Assert.True(ledger.TryAccept(new RowEvidenceIdentity("hash-one", "Data", 1), new RowEvidenceIdentity("hash-two", "Data", 1), out _));
        Assert.Equal(2, Assert.Single(ledger.Groups).Count);
    }

    [Fact]
    public void Empty_and_one_sided_sheets_have_complete_coverage()
    {
        AssertCoverage(Align(Array.Empty<string[]>(), Array.Empty<string[]>(), 1));
        AssertCoverage(Align(Array.Empty<string[]>(), new[] { new[] { "a" }, new[] { "b" } }, 1));
        AssertCoverage(Align(new[] { new[] { "a" } }, Array.Empty<string[]>(), 1));
    }

    [Fact]
    public void Insert_delete_metamorphic_trials_preserve_unique_rows()
    {
        for (int seed = 0; seed < 50; seed++)
        {
            var random = new Random(seed); var a = Enumerable.Range(1, 40).Select(i => new[] { "original-" + i }).ToArray();
            var b = a.Where(_ => random.Next(5) != 0).ToList();
            for (int i = 0; i < 5; i++) b.Insert(random.Next(b.Count + 1), new[] { "insert-" + i });
            var result = Align(a, b.ToArray(), 1); AssertCoverage(result);
            foreach (var match in result.Matches) Assert.Equal(a[match.BeforeRow - 1][0], b[match.AfterRow - 1][0]);
            Assert.Equal(b.Count(r => r[0].StartsWith("original-")), result.Matches.Count);
        }
    }

    private static void AssertCoverage(ReviewRowAlignment result)
    {
        var a = result.Matches.Select(m => m.BeforeRow).Concat(result.Regions.SelectMany(r => Enumerable.Range(r.BeforeStart, r.BeforeCount))).OrderBy(r => r);
        var b = result.Matches.Select(m => m.AfterRow).Concat(result.Regions.SelectMany(r => Enumerable.Range(r.AfterStart, r.AfterCount))).OrderBy(r => r);
        Assert.Equal(Enumerable.Range(1, result.BeforeRows), a); Assert.Equal(Enumerable.Range(1, result.AfterRows), b);
    }
}
