using System;
using System.Text;
using ExcelPowerTools.Core.Csv;
using Xunit;

namespace ExcelPowerTools.Tests.Csv
{
    public class CsvWorkspaceLogicTests
    {
        [Theory]
        [InlineData("a.csv", true)]
        [InlineData("A.TSV", true)]
        [InlineData("a.TxT", true)]
        [InlineData("a.csv.bak", false)]
        [InlineData("a.xlsx", false)]
        public void Supported_file_extensions_are_case_insensitive(string path, bool expected) =>
            Assert.Equal(expected, CsvWorkspaceLogic.IsSupportedPath(path));

        [Fact]
        public void Duplicate_basenames_get_unique_suffixes()
        {
            Assert.Equal("signals", CsvWorkspaceLogic.UniqueSheetName(@"C:\one\signals.csv", new[] { "Index" }));
            Assert.Equal("signals (2)", CsvWorkspaceLogic.UniqueSheetName(
                @"D:\two\signals.csv", new[] { "Index", "signals" }));
            Assert.Equal("signals (3)", CsvWorkspaceLogic.UniqueSheetName(
                @"D:\two\signals.tsv", new[] { "Index", "signals", "signals (2)" }));
        }

        [Fact]
        public void Long_name_is_truncated_before_suffix()
        {
            string stem = new string('x', 40);
            string name = CsvWorkspaceLogic.UniqueSheetName(stem + ".csv", new[] { stem.Substring(0, 31) });
            Assert.Equal(31, name.Length);
            Assert.EndsWith(" (2)", name);
        }

        [Fact]
        public void Invalid_characters_are_replaced_and_index_is_reserved_by_collision()
        {
            Assert.Equal("a_b_c", CsvWorkspaceLogic.UniqueSheetName("a[b]c.csv", new string[0]));
            Assert.Equal("Index (2)", CsvWorkspaceLogic.UniqueSheetName("Index.csv", new[] { "Index" }));
        }

        [Fact]
        public void Hyperlink_subaddress_escapes_apostrophes() =>
            Assert.Equal("'Joe''s data'!A1", CsvWorkspaceLogic.EscapeSheetNameForSubAddress("Joe's data"));

        [Fact]
        public void Sheet_anchor_formula_escapes_apostrophes() =>
            Assert.Equal("='Joe''s data'!$1:$1048576", CsvWorkspaceLogic.SheetAnchorFormula("Joe's data"));

        [Theory]
        [InlineData(1, "='Joe''s data'!$A:$A")]
        [InlineData(26, "='Joe''s data'!$Z:$Z")]
        [InlineData(27, "='Joe''s data'!$AA:$AA")]
        [InlineData(16384, "='Joe''s data'!$XFD:$XFD")]
        public void Sheet_column_anchor_formula_uses_invariant_excel_columns(int column, string expected) =>
            Assert.Equal(expected, CsvWorkspaceLogic.SheetColumnAnchorFormula("Joe's data", column));

        [Theory]
        [InlineData(0)]
        [InlineData(16385)]
        public void Sheet_column_anchor_rejects_invalid_columns(int column) =>
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                CsvWorkspaceLogic.SheetColumnAnchorFormula("Sheet1", column));

        [Theory]
        [InlineData("='My Sheet'!#REF!", "My Sheet")]
        [InlineData("='O''Brien'!#REF!", "O'Brien")]
        [InlineData("=Data!#REF!", "Data")]
        public void Broken_sheet_anchor_can_still_identify_its_worksheet(
            string refersTo, string sheetName) =>
            Assert.True(CsvWorkspaceLogic.BrokenSheetAnchorTargetsSheet(refersTo, sheetName));

        [Theory]
        [InlineData("='Other'!#REF!", "Data")]
        [InlineData("='Data'!$D:$D", "Data")]
        [InlineData("=#REF!", "Data")]
        public void Broken_sheet_anchor_rejects_unrelated_references(
            string refersTo, string sheetName) =>
            Assert.False(CsvWorkspaceLogic.BrokenSheetAnchorTargetsSheet(refersTo, sheetName));

        [Theory]
        [InlineData(false, false, false, false)]
        [InlineData(true, true, false, false)]
        [InlineData(true, true, true, true)]
        [InlineData(true, false, false, true)]
        public void Closing_workbook_sheet_identity_fallback_ignores_live_cross_workbook_copies(
            bool sheetIdentityMatches,
            bool registeredWorkbookIsOpen,
            bool registeredWorkbookIsClosing,
            bool expected) =>
            Assert.Equal(expected, CsvWorkspaceLogic.UseSheetIdentityForClosingWorkbook(
                sheetIdentityMatches, registeredWorkbookIsOpen, registeredWorkbookIsClosing));

        [Theory]
        [InlineData(false, false, false)]
        [InlineData(true, false, true)]
        [InlineData(false, true, true)]
        [InlineData(true, true, true)]
        public void Deferred_cleanup_preserves_sessions_while_either_runtime_owner_is_open(
            bool workspaceIdOwnerIsOpen, bool rememberedNameOwnerIsOpen, bool expected) =>
            Assert.Equal(expected, CsvWorkspaceLogic.KeepRuntimeSessionDuringCleanup(
                workspaceIdOwnerIsOpen, rememberedNameOwnerIsOpen));

        [Theory]
        [InlineData("__PT_CSV_SHEET_0123456789abcdef0123456789abcdef")]
        [InlineData("Book1!__PT_CSV_SHEET_0123456789ABCDEF0123456789ABCDEF")]
        [InlineData("'Book!One.xlsx'!__PT_CSV_SHEET_0123456789abcdef0123456789abcdef")]
        public void Sheet_anchor_ids_parse_from_short_or_qualified_names(string name) =>
            Assert.Equal("0123456789abcdef0123456789abcdef", CsvWorkspaceLogic.TryParseSheetAnchorId(name));

        [Theory]
        [InlineData("__PT_CSV_SHEET_bad")]
        [InlineData("Sheet1!UnrelatedName")]
        [InlineData("")]
        public void Invalid_sheet_anchor_names_do_not_parse(string name) =>
            Assert.Null(CsvWorkspaceLogic.TryParseSheetAnchorId(name));

        [Fact]
        public void Excel_reserved_history_name_is_disambiguated() =>
            Assert.Equal("History (2)", CsvWorkspaceLogic.UniqueSheetName("History.csv", new string[0]));

        [Fact]
        public void Truncation_does_not_split_a_unicode_surrogate_pair()
        {
            string stem = new string('x', 30) + "😀tail";
            string name = CsvWorkspaceLogic.UniqueSheetName(stem + ".csv", new string[0]);
            Assert.Equal(30, name.Length);
            Assert.DoesNotContain('\uFFFD', name);
            Assert.False(char.IsSurrogate(name[name.Length - 1]));
        }

        [Fact]
        public void Canonical_path_identity_distinguishes_duplicate_basenames()
        {
            string relative = @"folder\signals.csv";
            Assert.True(CsvWorkspaceLogic.SameSourcePath(relative, CsvWorkspaceLogic.CanonicalPath(relative)));
            Assert.False(CsvWorkspaceLogic.SameSourcePath(@"C:\one\signals.csv", @"D:\two\signals.csv"));
        }

        [Fact]
        public void Empty_or_single_column_tsv_keeps_tab_hint()
        {
            var empty = CsvFile.LoadForPath(new byte[0], "empty.tsv");
            var oneColumn = CsvFile.LoadForPath(Encoding.UTF8.GetBytes("a\r\nb\r\n"), "one.tsv");
            Assert.Equal('\t', empty.Format.Delimiter);
            Assert.Equal('\t', oneColumn.Format.Delimiter);
        }

        [Fact]
        public void Multi_column_tsv_uses_tabs_and_round_trips()
        {
            byte[] input = Encoding.UTF8.GetBytes("name\tvalue\r\na,b\t1,2\r\n");
            var doc = CsvFile.LoadForPath(input, "data.tsv");

            Assert.Equal('\t', doc.Format.Delimiter);
            Assert.Equal(new[] { "a,b", "1,2" }, doc.Records[1].Fields);
            Assert.Equal(input, doc.Save());
        }

        [Fact]
        public void Sheet_name_collisions_are_case_insensitive() =>
            Assert.Equal("signals (2)",
                CsvWorkspaceLogic.UniqueSheetName("signals.csv", new[] { "SIGNALS" }));

        [Fact]
        public void Sep_directive_overrides_tsv_extension_hint()
        {
            var doc = CsvFile.LoadForPath(Encoding.UTF8.GetBytes("sep=;\r\na;b\r\n"), "data.tsv");
            Assert.Equal(';', doc.Format.Delimiter);
        }

        [Theory]
        [InlineData("0123456789abcdef0123456789abcdef", true)]
        [InlineData("0123456789ABCDEF0123456789ABCDEF", true)]
        [InlineData("Sheet1!$A$1", false)]
        [InlineData("0123456789abcdef", false)]
        [InlineData("g123456789abcdef0123456789abcdef", false)]
        public void Workspace_ids_are_exact_32_digit_hex_values(string value, bool expected) =>
            Assert.Equal(expected, CsvWorkspaceLogic.IsWorkspaceId(value));
    }
}
