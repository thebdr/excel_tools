using System.Collections.Generic;
using System.Linq;
using ExcelPowerTools.Core.Experimental.IoBridge;
using Xunit;

namespace ExcelPowerTools.Tests.Experimental.IoBridge
{
    public class IoBridgeColumnLayoutTests
    {
        private static ColumnDef Col(string key, string side = "IO") =>
            new ColumnDef { NormalizedHeader = key, SampleHeader = key.ToUpperInvariant(), Side = side };

        private static IoBridgeColumnConfig Config()
        {
            var config = new IoBridgeColumnConfig();
            config.Groups.Add(new GroupSpec { Name = "Ident", Order = 1 });
            config.Groups.Add(new GroupSpec { Name = "Detail", Order = 2, HiddenByDefault = true });
            config.ColumnGroups["mnemonic"] = "Ident";
            config.ColumnGroups["bit"] = "Ident";
            config.ColumnGroups["manufacturer"] = "Detail";
            config.ColumnGroups["partno."] = "Detail";
            return config;
        }

        [Fact]
        public void Columns_never_move_and_groups_become_runs_in_source_order()
        {
            // Source order deliberately interleaves the groups — positions must survive untouched,
            // so the workspace reads side by side with the original document.
            var columns = new[]
            {
                Col("manufacturer"), Col("mnemonic"), Col("free1"), Col("bit"), Col("partno."), Col("free2"),
            };

            ColumnLayout layout = IoBridgeColumnLayout.ComputeLayout(columns, Config());

            Assert.Equal(new[] { "manufacturer", "mnemonic", "free1", "bit", "partno.", "free2" },
                layout.Columns.Select(c => c.Def.NormalizedHeader));

            Assert.Equal(
                new[]
                {
                    ("Detail", 0, 0, true),    // manufacturer
                    ("Ident", 1, 1, false),    // mnemonic
                    ("", 2, 2, false),         // free1
                    ("Ident", 3, 3, false),    // bit — a second run of the same group
                    ("Detail", 4, 4, true),    // partno.
                    ("", 5, 5, false),         // free2
                },
                layout.Spans.Select(s => (s.Name, s.FirstIndex, s.LastIndex, s.Hidden)));
        }

        [Fact]
        public void Contiguous_same_group_columns_form_one_span()
        {
            var config = Config();
            ColumnLayout layout = IoBridgeColumnLayout.ComputeLayout(
                new[] { Col("mnemonic"), Col("bit"), Col("free1"), Col("manufacturer"), Col("partno.") }, config);

            Assert.Equal(3, layout.Spans.Count);
            Assert.Equal(("Ident", 0, 1, false),
                (layout.Spans[0].Name, layout.Spans[0].FirstIndex, layout.Spans[0].LastIndex, layout.Spans[0].Hidden));
            Assert.Equal(("", 2, 2, false),
                (layout.Spans[1].Name, layout.Spans[1].FirstIndex, layout.Spans[1].LastIndex, layout.Spans[1].Hidden));
            Assert.Equal(("Detail", 3, 4, true),
                (layout.Spans[2].Name, layout.Spans[2].FirstIndex, layout.Spans[2].LastIndex, layout.Spans[2].Hidden));
        }

        [Fact]
        public void A_group_referenced_but_not_configured_keeps_position_and_is_never_hidden()
        {
            var config = Config();
            config.ColumnGroups["free1"] = "Mystery";   // no GroupSpec for it

            ColumnLayout layout = IoBridgeColumnLayout.ComputeLayout(
                new[] { Col("free1"), Col("mnemonic"), Col("free2") }, config);

            Assert.Equal(new[] { "free1", "mnemonic", "free2" },
                layout.Columns.Select(c => c.Def.NormalizedHeader));
            GroupSpan mystery = layout.Spans.Single(s => s.Name == "Mystery");
            Assert.Equal((0, 0, false), (mystery.FirstIndex, mystery.LastIndex, mystery.Hidden));
        }

        [Fact]
        public void The_configured_group_order_never_moves_columns()
        {
            var config = new IoBridgeColumnConfig();
            config.Groups.Add(new GroupSpec { Name = "B", Order = 1 });
            config.Groups.Add(new GroupSpec { Name = "A", Order = 2 });
            config.ColumnGroups["x"] = "B";
            config.ColumnGroups["y"] = "A";
            config.ColumnGroups["z"] = "B";

            ColumnLayout layout = IoBridgeColumnLayout.ComputeLayout(
                new[] { Col("x"), Col("y"), Col("z") }, config);

            Assert.Equal(new[] { "x", "y", "z" }, layout.Columns.Select(c => c.Def.NormalizedHeader));
            Assert.Equal(new[] { "B", "A", "B" }, layout.Spans.Select(s => s.Name));
        }

        [Fact]
        public void Empty_config_yields_one_visible_span_in_source_order()
        {
            ColumnLayout layout = IoBridgeColumnLayout.ComputeLayout(
                new[] { Col("a"), Col("b") }, new IoBridgeColumnConfig());

            Assert.Equal(new[] { "a", "b" }, layout.Columns.Select(c => c.Def.NormalizedHeader));
            GroupSpan only = Assert.Single(layout.Spans);
            Assert.Equal("", only.Name);
            Assert.False(only.Hidden);
        }

        [Fact]
        public void Columns_sheet_round_trips_the_configuration()
        {
            var known = new[] { Col("mnemonic"), Col("manufacturer"), Col("free1") };
            List<string[]> rows = IoBridgeColumnLayout.RenderColumnsSheet(Config(), known);

            Assert.Equal(IoBridgeColumnLayout.ColumnsSheetHeaders, rows[0]);
            Assert.Equal(4, rows.Count);

            IoBridgeColumnConfig back = IoBridgeColumnLayout.ParseColumnsSheet(rows);
            Assert.Equal("Ident", back.ColumnGroups["mnemonic"]);
            Assert.Equal("Detail", back.ColumnGroups["manufacturer"]);
            Assert.Equal("", back.ColumnGroups["free1"]);

            GroupSpec detail = back.FindGroup("Detail")!;
            Assert.True(detail.HiddenByDefault);
            Assert.Equal(2, detail.Order);
        }

        [Fact]
        public void Parse_takes_the_first_row_of_a_group_and_tolerates_noise()
        {
            var rows = new List<string[]>
            {
                (string[])IoBridgeColumnLayout.ColumnsSheetHeaders.Clone(),
                new[] { "bit", "BIT", "IO", "Ident", "3", "yes" },
                new[] { "mnemonic", "MN", "IO", "Ident", "9", "" },     // later row must not redefine Ident
                new[] { "", "stray note below the table" },
                new[] { "id", "ID", "IO", "", "", "" },
            };

            IoBridgeColumnConfig config = IoBridgeColumnLayout.ParseColumnsSheet(rows);

            GroupSpec ident = config.FindGroup("Ident")!;
            Assert.Equal(3, ident.Order);
            Assert.True(ident.HiddenByDefault);
            Assert.Equal("", config.ColumnGroups["id"]);
        }

        [Theory]
        [InlineData("1", true)]
        [InlineData("x", true)]
        [InlineData("YES", true)]
        [InlineData("true", true)]
        [InlineData("☑", true)]
        [InlineData("", false)]
        [InlineData("0", false)]
        [InlineData("no", false)]
        public void Hidden_flag_accepts_the_obvious_spellings(string text, bool expected)
        {
            var rows = new List<string[]>
            {
                new[] { "col", "COL", "IO", "G", "1", text },
            };
            Assert.Equal(expected, IoBridgeColumnLayout.ParseColumnsSheet(rows).FindGroup("G")!.HiddenByDefault);
        }
    }
}
