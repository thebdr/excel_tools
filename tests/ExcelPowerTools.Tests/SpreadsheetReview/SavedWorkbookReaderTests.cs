using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelPowerTools.Core.SpreadsheetReview;
using Xunit;
using Xunit.Abstractions;

namespace ExcelPowerTools.Tests.SpreadsheetReview;

public class SavedWorkbookReaderTests
{
    // User-confirmed scope for the five sorter trial artifacts, not a product-wide default.
    private const string SorterMainSheet = "R1.S1.C1";
    private const int SorterLastOfficialColumn = 26; // Z, inclusive.
    private readonly ITestOutputHelper output;
    public SavedWorkbookReaderTests(ITestOutputHelper output) => this.output = output;

    [Fact]
    public void Reads_saved_values_formulas_and_manual_formatting_without_modifying_source()
    {
        using var fixture = new SavedFixture("""
            <row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1"><f>A2+1</f><v>3</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#REF!</v></c></row>
            <row r="2"><c r="A2"><v>2.0</v></c></row>
            <row r="50000"><c r="A50000" s="1"/></row>
            """, sharedStrings: "<si><r><rPr><strike/></rPr><t>ab</t></r><r><rPr><strike val=\"0\"/></rPr><t>cd</t></r><rPh sb=\"0\" eb=\"1\"><t>phonetic-not-data</t></rPh></si>");
        string before = Fingerprint(fixture.Path); var events = new List<ReviewReadProgress>();
        var result = new SavedWorkbookReader().Read(fixture.Path, "revision-a", events.Add);
        var sheet = Assert.Single(result.Sheets);
        Assert.Equal(before, Fingerprint(fixture.Path)); Assert.Equal(before, result.Sha256);
        Assert.Equal("abcd", sheet.Cell(1, 1).Content.Constant.OriginalEncoding);
        Assert.Equal(new[] { true, true, false, false }, sheet.Cell(1, 1).Strike.Value.CharacterStrike);
        Assert.Equal("FFFF0000", sheet.Cell(1, 1).Interior.Value.ForegroundArgb);
        Assert.Equal("=A2+1", sheet.Cell(1, 2).Content.Formula); Assert.Equal("3", sheet.Cell(1, 2).Content.SavedResult.Value.OriginalEncoding);
        Assert.Equal(ReviewScalarKind.Boolean, sheet.Cell(1, 3).Content.Constant.Kind); Assert.Equal(ReviewScalarKind.Error, sheet.Cell(1, 4).Content.Constant.Kind);
        Assert.Equal(2, sheet.LastContentRow(4)); Assert.Equal(6, result.StoredCellCounts["Data"]);
        Assert.True(sheet.Cell(50000, 1).Strike.Value.CellStrike); Assert.Empty(result.Issues); Assert.NotEmpty(events);
    }

    [Fact]
    public void Blank_manual_styles_are_available_outside_content_extent()
    {
        using var fixture = new SavedFixture("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>data</t></is></c></row><row r=\"20\" s=\"1\" customFormat=\"1\"/>",
            columns: "<cols><col min=\"2\" max=\"2\" style=\"1\"/></cols>");
        var sheet = Assert.Single(new SavedWorkbookReader().Read(fixture.Path).Sheets);
        Assert.Equal(1, sheet.LastContentRow(2)); Assert.False(sheet.Cell(20, 1).HasContent);
        Assert.Equal("FFFF0000", sheet.Cell(20, 1).Interior.Value.ForegroundArgb);
        Assert.Equal("FFFF0000", sheet.Cell(100, 2).Interior.Value.ForegroundArgb);
        Assert.Equal(CellInterior.NoFill, sheet.Cell(100, 1).Interior.Value);
    }

    [Fact]
    public void Shared_formula_followers_are_unresolved_not_saved_value_substitutes()
    {
        using var fixture = new SavedFixture("<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"0\" ref=\"A1:A2\">B1+1</f><v>3</v></c></row><row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/><v>4</v></c></row>");
        var result = new SavedWorkbookReader().Read(fixture.Path); var sheet = Assert.Single(result.Sheets);
        Assert.Equal("=B1+1", sheet.Cell(1, 1).Content.Formula);
        Assert.Equal(CellObservationState.Observed, sheet.Cell(2, 1).State);
        Assert.True(sheet.Cell(2, 1).Content.IsFormula); Assert.False(sheet.Cell(2, 1).Content.FormulaEvidence.IsAvailable);
        Assert.Equal("4", sheet.Cell(2, 1).Content.SavedResult.Value.OriginalEncoding);
        Assert.True(CellHistoryComparison.Compare(new[] { sheet.Cell(2, 1), sheet.Cell(2, 1) }).HasIncompleteEvidence);
        Assert.Equal("SR.READ.FORMULA_EXPANSION", Assert.Single(result.Issues).Code);
        Assert.Equal(2, sheet.LastContentRow(1));
    }

    [Fact]
    public void Invalid_shared_string_is_failed_not_blank()
    {
        using var fixture = new SavedFixture("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>999</v></c></row>");
        var result = new SavedWorkbookReader().Read(fixture.Path);
        Assert.Equal(CellObservationState.Failed, result.Sheets[0].Cell(1, 1).State);
        Assert.Equal("SR.READ.CELL", Assert.Single(result.Issues).Code);
    }

    [Fact]
    public void External_relationships_and_macros_are_not_executed_or_followed()
    {
        using var fixture = new SavedFixture("<row r=\"1\"><c r=\"A1\"><f>'https://invalid.example/[other.xlsx]Sheet1'!A1</f><v>9</v></c></row>",
            extraRelationship: "<Relationship Id=\"external\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" Target=\"https://invalid.example/never-contact\" TargetMode=\"External\"/>");
        var result = new SavedWorkbookReader().Read(fixture.Path);
        Assert.Contains("https://invalid.example", result.Sheets[0].Cell(1, 1).Content.Formula);
        Assert.Equal("9", result.Sheets[0].Cell(1, 1).Content.SavedResult.Value.OriginalEncoding);
    }

    [Fact]
    public void Duplicate_package_parts_are_rejected()
    {
        using var fixture = new SavedFixture("", duplicatePart: true);
        Assert.Throws<InvalidDataException>(() => new SavedWorkbookReader().Read(fixture.Path));
    }

    [Fact]
    public void Date_system_is_retained_and_unsupported_iso_date_is_not_coerced()
    {
        using var fixture = new SavedFixture("<row r=\"1\"><c r=\"A1\" t=\"d\"><v>2026-09-10T00:00:00</v></c><c r=\"B1\"><v>39267</v></c></row>", date1904: true);
        var result = new SavedWorkbookReader().Read(fixture.Path);
        Assert.True(result.Date1904); Assert.Equal(CellObservationState.Failed, result.Sheets[0].Cell(1, 1).State);
        Assert.Equal("39267", result.Sheets[0].Cell(1, 2).Content.Constant.OriginalEncoding);
    }

    [Fact]
    public void Numeric_serialization_change_does_not_change_alignment()
    {
        using var a = new SavedFixture("<row r=\"1\"><c r=\"A1\"><v>1.0</v></c></row>");
        using var b = new SavedFixture("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");
        var reader = new SavedWorkbookReader(); var x = reader.Read(a.Path); var y = reader.Read(b.Path);
        Assert.Single(new OrderedRowAlignment().Align(x.Sheets[0], y.Sheets[0], ResolvedReviewScope.SameColumns(1)).Matches);
    }

    [Fact]
    public void Spreadsheet_escape_sequences_and_literal_escapes_are_distinct()
    {
        using var fixture = new SavedFixture("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"str\"><v>_x005F_x000D_</v></c></row>",
            sharedStrings: "<si><r><t>a_x000D_</t></r><r><rPr><strike/></rPr><t>b</t></r></si>");
        var sheet = new SavedWorkbookReader().Read(fixture.Path).Sheets[0];
        Assert.Equal("a\rb", sheet.Cell(1, 1).Content.Constant.OriginalEncoding);
        Assert.Equal(new[] { false, false, true }, sheet.Cell(1, 1).Strike.Value.CharacterStrike);
        Assert.Equal("_x000D_", sheet.Cell(1, 2).Content.Constant.OriginalEncoding);
    }

    [RealReviewTrial]
    public void Real_saved_revisions_readonly_alignment_trial()
    {
        string root = Environment.GetEnvironmentVariable("EXCEL_REVIEW_TRIAL_ROOT")!;
        string[] names = { @"old\8FVZ_IO0001_R1.0 - sorter.xlsx", @"old\8FVZ_IO0001_R1.1 - sorter.xlsx",
            @"old\8FVZ_IO0001_R1.2 - sorter.xlsx", "8FVZ_IO0001_R1.3 - sorter.xlsx", @"_latest\8FVZ_IO0001_R2.3 - sorter.xlsx" };
        SavedReviewWorkbook? previous = null;
        foreach (string name in names)
        {
            string path = System.IO.Path.Combine(root, name); Assert.True(File.Exists(path), "Real trial source is missing: " + path);
            string hashBefore = Fingerprint(path); var watch = Stopwatch.StartNew();
            var current = new SavedWorkbookReader().Read(path);
            Assert.Equal(hashBefore, current.Sha256); Assert.Equal(hashBefore, Fingerprint(path));
            output.WriteLine($"SOURCE {name}: {current.ByteLength:N0} bytes; {watch.Elapsed.TotalSeconds:F2}s; {current.Sheets.Count} sheets; {current.Issues.Count} reader issue categories; SHA256 {current.Sha256}");
            foreach (var sheet in current.Sheets)
            {
                output.WriteLine($"  SHEET {sheet.Name}: {current.StoredCellCounts[sheet.Name]:N0} stored, {sheet.Cells.Count:N0} retained content/typed observations, last content {sheet.LastContentRow(16384)}x{sheet.LastContentColumn}");
                if (sheet.Name == SorterMainSheet)
                    output.WriteLine($"  CONFIRMED SCOPE {sheet.Name}: A:Z inclusive; last in-scope content row {sheet.LastContentRow(SorterLastOfficialColumn)}; AA onward excluded from matching and differences.");
            }
            foreach (var issue in current.Issues) output.WriteLine(issue.ToString());
            if (previous != null)
            {
                foreach (var a in previous.Sheets)
                {
                    var b = current.Sheets.SingleOrDefault(s => s.Name == a.Name);
                    if (b == null) { output.WriteLine($"  UNPAIRED before sheet: {a.Name}"); continue; }
                    bool confirmedSorterScope = a.Name == SorterMainSheet;
                    int cap = confirmedSorterScope ? SorterLastOfficialColumn : Math.Max(a.LastContentColumn, b.LastContentColumn);
                    if (cap == 0) continue;
                    // Other sheets remain discovery coverage. Fixed physical column pairing is still provisional on all sheets.
                    if (cap > 200 || Math.Max(a.LastContentRow(cap), b.LastContentRow(cap)) > 20000)
                    { output.WriteLine($"  SCOPE REVIEW NEEDED {a.Name}: content bounds exceed the initial alignment trial; no truncation/alignment attempted."); continue; }
                    watch.Restart();
                    var aligned = new OrderedRowAlignment().Align(a, b, ResolvedReviewScope.SameColumns(cap));
                    Assert.Equal(aligned.BeforeRows, aligned.Matches.Count + aligned.Regions.Sum(r => r.BeforeCount));
                    Assert.Equal(aligned.AfterRows, aligned.Matches.Count + aligned.Regions.Sum(r => r.AfterCount));
                    int changed = 0, incomplete = 0;
                    foreach (var match in aligned.Matches)
                        for (int c = 1; c <= cap; c++)
                        {
                            var history = CellHistoryComparison.Compare(new[] { a.Cell(match.BeforeRow, c), b.Cell(match.AfterRow, c) });
                            if (history.IsRelevant && history.Transitions.Any(t => (t.Changes & ~ReviewCellChange.IncompleteEvidence) != 0)) changed++;
                            if (history.HasIncompleteEvidence) incomplete++;
                        }
                    string scopeLabel = confirmedSorterScope ? "confirmed A:Z scope; provisional physical pairing" : $"provisional fixed columns 1..{cap}";
                    output.WriteLine($"  PAIR {a.Name} ({scopeLabel}): {aligned.Matches.Count} exact-constant anchors, {aligned.Regions.Count} raw unresolved/one-sided regions, {changed} changed cells on anchors, {incomplete} incomplete cells; {watch.Elapsed.TotalSeconds:F2}s");
                    foreach (var issue in aligned.Issues) output.WriteLine(issue.ToString());
                }
                foreach (var b in current.Sheets.Where(s => previous.Sheets.All(a => a.Name != s.Name))) output.WriteLine($"  UNPAIRED after sheet: {b.Name}");
            }
            previous = current;
        }
    }

    private static string Fingerprint(string path)
    { using var stream = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(); }

    private sealed class RealReviewTrialAttribute : FactAttribute
    {
        public RealReviewTrialAttribute()
        { if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("EXCEL_REVIEW_TRIAL_ROOT"))) Skip = "Set EXCEL_REVIEW_TRIAL_ROOT to the user-approved sorter revision folder for the read-only trial."; }
    }

    private sealed class SavedFixture : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PowerTools-reader-" + Guid.NewGuid().ToString("N") + ".xlsx");
        public SavedFixture(string rows, string sharedStrings = "", string columns = "", string extraRelationship = "", bool date1904 = false, bool duplicatePart = false)
        {
            using var file = File.Create(Path); using var zip = new ZipArchive(file, ZipArchiveMode.Create);
            Add("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            Add("xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"" + (date1904 ? "1" : "0") + "\"/><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"sheet\"/></sheets></workbook>");
            Add("xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sheet\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"strings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/><Relationship Id=\"styles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" + extraRelationship + "</Relationships>");
            Add("xl/sharedStrings.xml", "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + sharedStrings + "</sst>");
            Add("xl/styles.xml", "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts><font/><font><strike/></font></fonts><fills><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFF0000\"/><bgColor indexed=\"64\"/></patternFill></fill></fills><cellStyleXfs><xf fontId=\"0\" fillId=\"0\"/></cellStyleXfs><cellXfs><xf fontId=\"0\" fillId=\"0\" xfId=\"0\"/><xf fontId=\"1\" fillId=\"1\" xfId=\"0\" applyFill=\"1\" applyFont=\"1\"/></cellXfs></styleSheet>");
            Add("xl/worksheets/sheet1.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:XFD1048576\"/>" + columns + "<sheetData>" + rows + "</sheetData></worksheet>");
            if (duplicatePart) Add("xl/sharedStrings.xml", "<duplicate/>");
            void Add(string name, string xml) { using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(xml); }
        }
        public void Dispose() => File.Delete(Path);
    }
}
