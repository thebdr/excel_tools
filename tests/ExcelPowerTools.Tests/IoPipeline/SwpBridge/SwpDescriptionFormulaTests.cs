using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Xunit;

namespace ExcelPowerTools.Tests.IoPipeline.SwpBridge
{
    /// <summary>
    /// The Q/R rebuild: template truncated at its last placeholder, then " - " and the T/U cell.
    /// The fixtures are the real IL03_BELT ones — template "[1]DB_IL_BELT!$K$122", %1 = $B$6,
    /// %2 = F679, suffix T778 — so what is asserted here is what the field file must hold.
    /// </summary>
    public class SwpDescriptionFormulaTests
    {
        private const string BeltFormula =
            "SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%1\",$B$6),\"%2\",F679)";

        private const string BeltRebuilt =
            "_xlfn.LET(_xlpm.t,[1]DB_IL_BELT!$K$122,_xlpm.p,FIND(CHAR(1),SUBSTITUTE(_xlpm.t,\"%\",CHAR(1)," +
            "LEN(_xlpm.t)-LEN(SUBSTITUTE(_xlpm.t,\"%\",\"\")))),SUBSTITUTE(SUBSTITUTE(LEFT(_xlpm.t,IFERROR(_xlpm.p,LEN(_xlpm.t))+1)," +
            "\"%1\",$B$6),\"%2\",F679)&\" - \"&T778)";

        private static SwpDescriptionFormula.Parsed Parse(string formula)
        {
            Assert.True(SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed parsed),
                "should have parsed: " + formula);
            return parsed;
        }

        [Fact]
        public void The_plain_swp_shape_yields_the_canonical_rebuild()
        {
            Assert.Equal(BeltRebuilt, SwpDescriptionFormula.Build(Parse(BeltFormula), "T778"));
        }

        // ---- what a kept formula shows ------------------------------------------------------

        /// <summary>The field CABINET row (FVZ G1, ILGRP01_CABINET row 216): a CONFIG template ending
        /// at "…%2 - ", concatenated with the row's own description cell.</summary>
        private const string CabinetFormula =
            "SUBSTITUTE(SUBSTITUTE([1]DB_ILGRP_CABINET!$L$57,\"%1\",$B$6),\"%2\",F181)&\" \"&U216";

        private const string CabinetSuffix = "CIRCUIT BREAKER TRIPPED 24V IM 155 AND I/O MODULE";

        private const string CabinetShown =
            "INTRODUCTION GRP01 - ARMOIRE DE GROUPE =S1+MS1.CC1 -  " + CabinetSuffix;

        [Fact]
        public void A_cabinet_formula_shows_its_description_after_a_template_with_no_words_of_its_own()
        {
            Assert.True(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                CabinetFormula, CabinetShown, "U216", CabinetSuffix, "INTRODUCTION GRP01", "=S1+MS1.CC1"));
        }

        [Fact]
        public void An_empty_description_cell_still_reads_as_the_alarms_only_name()
        {
            Assert.True(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                CabinetFormula, "INTRODUCTION GRP01 - ARMOIRE DE GROUPE =S1+MS1.CC1 -  ", "U216", "",
                "INTRODUCTION GRP01", "=S1+MS1.CC1"));
        }

        [Fact]
        public void A_template_that_names_the_alarm_itself_is_not_a_bare_template()
        {
            // The same concatenation, but the template's own tail says something: a description
            // written into U would stack a second sentence onto it.
            Assert.False(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                CabinetFormula, "INTRODUCTION GRP01 - BELT =S1+MS1.CC1 - MOTOR OVERLOAD =S1+MS1.CC1-F22001",
                "U216", "=S1+MS1.CC1-F22001", "INTRODUCTION GRP01", "=S1+MS1.CC1"));
        }

        [Fact]
        public void A_formula_not_showing_the_cell_is_never_a_description_slot()
        {
            Assert.False(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                BeltFormula, "INTRODUCTION =S1+IN11 - TAPIS =S1+IN11.08 - ", "T778", "",
                "INTRODUCTION =S1+IN11", "=S1+IN11.08"));
        }

        [Fact]
        public void A_built_formula_shows_its_description_after_the_head_by_construction()
        {
            Assert.True(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(BeltRebuilt, "", "T778", "", "", ""));
        }

        [Fact]
        public void An_unreadable_display_answers_false()
        {
            // No cached text; a cache not ending with the description (patched, not recalculated);
            // no anchor in the text. Nothing is known, so nothing is claimed.
            Assert.False(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                CabinetFormula, "", "U216", CabinetSuffix, "INTRODUCTION GRP01", "=S1+MS1.CC1"));
            Assert.False(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                CabinetFormula, CabinetShown, "U216", "DISJONCTEUR DÉCLENCHÉ", "INTRODUCTION GRP01", "=S1+MS1.CC1"));
            Assert.False(SwpDescriptionFormula.ShowsSuffixAfterBareTemplate(
                CabinetFormula, CabinetShown, "U216", CabinetSuffix, "SOMEWHERE ELSE", "=S9+XX9"));
        }

        [Fact]
        public void Cell_references_are_bounded_on_both_sides()
        {
            Assert.True(SwpDescriptionFormula.ReferencesCell(CabinetFormula, "U216"));
            Assert.True(SwpDescriptionFormula.ReferencesCell("SUBSTITUTE(A1&$U$216,\"%1\",F181)", "U216"));
            Assert.True(SwpDescriptionFormula.ReferencesCell("x&u216", "U216"));
            Assert.False(SwpDescriptionFormula.ReferencesCell("x&AU216", "U216"));
            Assert.False(SwpDescriptionFormula.ReferencesCell("x&U2160", "U216"));
            Assert.False(SwpDescriptionFormula.ReferencesCell(CabinetFormula, "T216"));
            Assert.False(SwpDescriptionFormula.ReferencesCell("", "U216"));
        }

        /// <summary>What the first build emitted: bare LET parameter names. Excel's loader removed
        /// this formula on open in the field G2 ("Removed Records: Formula from …").</summary>
        private const string BeltRebuiltUnloadable =
            "_xlfn.LET(t,[1]DB_IL_BELT!$K$122,p,FIND(CHAR(1),SUBSTITUTE(t,\"%\",CHAR(1)," +
            "LEN(t)-LEN(SUBSTITUTE(t,\"%\",\"\")))),SUBSTITUTE(SUBSTITUTE(LEFT(t,IFERROR(p,LEN(t))+1)," +
            "\"%1\",$B$6),\"%2\",F679)&\" - \"&T778)";

        /// <summary>A lone lower-case t or p that is not the tail of "_xlpm." — a bare parameter.</summary>
        private const string BareParameter = @"(?<![\w.])[tp](?![\w(])";

        [Fact]
        public void Every_LET_parameter_is_spelled_for_the_file()
        {
            // Inside LET the loader reads a bare name as an unknown reference and deletes the
            // whole formula on open — so every occurrence, declarations AND uses, carries _xlpm.
            // A single bare one is enough to lose the cell.
            string built = SwpDescriptionFormula.Build(Parse(BeltFormula), "T778");

            Assert.Equal(0, System.Text.RegularExpressions.Regex.Matches(built, BareParameter).Count);
            Assert.Equal(6, System.Text.RegularExpressions.Regex.Matches(built, @"_xlpm\.t\b").Count);
            Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(built, @"_xlpm\.p\b").Count);

            // …and the check itself can see the bug: the unloadable spelling trips it.
            Assert.NotEqual(0, System.Text.RegularExpressions.Regex.Matches(BeltRebuiltUnloadable, BareParameter).Count);
        }

        [Fact]
        public void The_unloadable_first_build_is_recognized_and_rebuilt_loadable()
        {
            // A file patched by the first build but not yet opened still holds the bare names, and
            // Excel would drop the formula the moment it opens it. Counting that as "already
            // rebuilt" would leave the loss armed: it parses back and rebuilds in the file spelling.
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(BeltRebuiltUnloadable));
            Assert.False(SwpDescriptionFormula.IsCurrentForm(BeltRebuiltUnloadable));
            Assert.True(SwpDescriptionFormula.IsCurrentForm(BeltRebuilt));

            SwpDescriptionFormula.Parsed parsed = Parse(BeltRebuiltUnloadable);
            Assert.Equal("[1]DB_IL_BELT!$K$122", parsed.TemplateRef);
            Assert.Equal("$B$6", parsed.Param1);
            Assert.Equal("F679", parsed.Param2);
            Assert.Equal(BeltRebuilt, SwpDescriptionFormula.Build(parsed, "T778"));
        }

        [Theory]
        // Starts like the first build but lacks the head it always emitted — nothing to cut on a guess.
        [InlineData("_xlfn.LET(t,[1]DB_IL_BELT!$K$122,p,1,t&T778)")]
        // A LET this class never wrote.
        [InlineData("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)")]
        public void Other_LET_formulas_are_left_alone(string formula)
        {
            Assert.False(SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed parsed));
            Assert.Null(parsed);
        }

        [Fact]
        public void LET_is_emitted_as_a_future_function()
        {
            // A bare modern name in the XML is read back as _xludf.<name> — an unknown
            // user-defined function — and the cell shows #NAME?. That is what the previous
            // wholesale CONCAT rewrite did to IL03_BELT!Q114/R114/R778.
            Assert.StartsWith("_xlfn.LET(", SwpDescriptionFormula.Build(Parse(BeltFormula), "T778"));
            Assert.DoesNotContain("_xludf.", SwpDescriptionFormula.Build(Parse(BeltFormula), "T778"));

            // …and nothing else in the shape is a future function, so nothing else is prefixed.
            foreach (string classic in new[] { "SUBSTITUTE(", "FIND(", "LEFT(", "LEN(", "CHAR(", "IFERROR(" })
                Assert.DoesNotContain("_xlfn." + classic, SwpDescriptionFormula.Build(Parse(BeltFormula), "T778"));
        }

        [Fact]
        public void Commas_separate_the_arguments()
        {
            // These strings only ever reach the XLSX XML, which is locale-independent — a
            // semicolon there is a broken formula whatever the content language of the file.
            Assert.DoesNotContain(";", SwpDescriptionFormula.Build(Parse(BeltFormula), "T778"));
        }

        [Fact]
        public void A_single_placeholder_template_drops_the_outer_substitute()
        {
            Assert.Equal(
                "_xlfn.LET(_xlpm.t,[1]DB_CABINETS_DIAG!$K$57,_xlpm.p,FIND(CHAR(1),SUBSTITUTE(_xlpm.t,\"%\",CHAR(1)," +
                "LEN(_xlpm.t)-LEN(SUBSTITUTE(_xlpm.t,\"%\",\"\")))),SUBSTITUTE(LEFT(_xlpm.t,IFERROR(_xlpm.p,LEN(_xlpm.t))+1),\"%1\",F181)" +
                "&\" - \"&T216)",
                SwpDescriptionFormula.Build(
                    Parse("SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57,\"%1\",F181)"), "T216"));
        }

        [Fact]
        public void The_label_extension_is_parsed_off_rather_than_kept()
        {
            // A CABINET row already carrying &" "&T rebuilds to the same shape as a plain one:
            // the built formula names the suffix cell itself, so keeping the tail would double it.
            SwpDescriptionFormula.Parsed parsed = Parse(BeltFormula + "&\" \"&T778");
            Assert.Equal("[1]DB_IL_BELT!$K$122", parsed.TemplateRef);
            Assert.Equal(BeltRebuilt, SwpDescriptionFormula.Build(parsed, "T778"));
        }

        [Fact]
        public void A_template_concatenated_inside_the_substitute_keeps_only_the_template()
        {
            SwpDescriptionFormula.Parsed parsed =
                Parse("SUBSTITUTE([1]DB_CABINETS_DIAG!$K$57&T216,\"%1\",F181)");
            Assert.Equal("[1]DB_CABINETS_DIAG!$K$57", parsed.TemplateRef);
            Assert.Equal("F181", parsed.Param1);
            Assert.Equal("", parsed.Param2);
        }

        [Fact]
        public void Either_nesting_order_parses()
        {
            SwpDescriptionFormula.Parsed parsed =
                Parse("SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%2\",F679),\"%1\",$B$6)");
            Assert.Equal("[1]DB_IL_BELT!$K$122", parsed.TemplateRef);
            Assert.Equal("$B$6", parsed.Param1);
            Assert.Equal("F679", parsed.Param2);
        }

        [Theory]
        // A wholesale-rewritten cell: the template reference is gone for good.
        [InlineData("_xlfn.CONCAT(UPPER($B$6),\" \",$F679,\" - \",T778)")]
        [InlineData("CONCAT(UPPER($B$6),\" \",$F679,\" - \",T778)")]
        // A pasted-over literal turned into a formula — a text head, not a template.
        [InlineData("\"CABINET =S1+MC1.CC1 - CIRCUIT BREAKER\"&\" \"&T216")]
        // Already rebuilt: a second rebuild would nest the LET inside itself.
        [InlineData(BeltRebuilt)]
        // Not a template substitution at all.
        [InlineData("T778")]
        [InlineData("")]
        [InlineData(null)]
        // A computed template operand is not something to truncate blindly.
        [InlineData("SUBSTITUTE(INDIRECT(\"[1]DB!K\"&A1),\"%1\",F181)")]
        // A placeholder marker that is not %1/%2.
        [InlineData("SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%3\",F679)")]
        public void Unrecoverable_shapes_are_refused(string formula)
        {
            Assert.False(SwpDescriptionFormula.TryParse(formula, out SwpDescriptionFormula.Parsed parsed));
            Assert.Null(parsed);
        }

        [Fact]
        public void An_unbalanced_formula_is_refused_rather_than_half_read()
        {
            Assert.False(SwpDescriptionFormula.TryParse(
                "SUBSTITUTE(SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%1\",$B$6),\"%2\",F679", out _));
        }

        [Fact]
        public void A_quoted_comma_does_not_split_an_argument()
        {
            SwpDescriptionFormula.Parsed parsed =
                Parse("SUBSTITUTE([1]DB_IL_BELT!$K$122,\"%1\",\"A, B\")");
            Assert.Equal("\"A, B\"", parsed.Param1);
        }

        [Fact]
        public void The_rendered_head_stops_at_the_last_placeholder_value()
        {
            // The block label CONTAINS the sheet description ("=S1+IN22" inside "=S1+IN22.05"),
            // so the later anchor must win or the head would be cut one placeholder too early.
            Assert.Equal("INDUCTION =S1+IN22 - BELT =S1+IN22.05",
                SwpDescriptionFormula.TruncatedRendered(
                    "INDUCTION =S1+IN22 - BELT =S1+IN22.05 - SPARE ALARM.63", "=S1+IN22", "=S1+IN22.05"));
        }

        [Fact]
        public void A_rendered_text_holding_neither_anchor_is_left_whole()
        {
            // Nothing is cut on a guess: better a stale tail in the preview than a mangled one.
            Assert.Equal("SOMETHING ELSE ENTIRELY",
                SwpDescriptionFormula.TruncatedRendered("SOMETHING ELSE ENTIRELY", "=S1+IN22", "=S1+IN22.05"));
            Assert.Equal("", SwpDescriptionFormula.TruncatedRendered(null, "", ""));
        }

        [Fact]
        public void The_rendered_final_is_head_separator_suffix()
        {
            Assert.Equal("INDUCTION =S1+IN22 - BELT =S1+IN22.05 - FIELD BUS FAILURE - DRIVE",
                SwpDescriptionFormula.RenderedFinal(
                    "INDUCTION =S1+IN22 - BELT =S1+IN22.05 - SPARE ALARM.63",
                    "=S1+IN22", "=S1+IN22.05", "FIELD BUS FAILURE - DRIVE"));
        }

        [Fact]
        public void Only_a_built_formula_is_recognized_as_one()
        {
            Assert.True(SwpDescriptionFormula.IsTruncatingForm(BeltRebuilt));
            Assert.False(SwpDescriptionFormula.IsTruncatingForm(BeltFormula));
            Assert.False(SwpDescriptionFormula.IsTruncatingForm("LET(t,1,t)"));
            Assert.False(SwpDescriptionFormula.IsTruncatingForm(null));
        }
    }
}
