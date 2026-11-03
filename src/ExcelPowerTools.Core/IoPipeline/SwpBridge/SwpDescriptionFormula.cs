using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The Q/R final-text formula: CONFIG_PLANT's template with the row's own description put
    /// where the template's trailing text was.
    ///
    /// A template reads "INDUCTION %1 - BELT %2 - SPARE ALARM.63": a head naming the place
    /// (%1 = the sheet description B6, %2 = the block's electrical label) and a tail naming the
    /// bit generically. The SWP's own T/U cell names the SAME bit in the plant's words
    /// ("FIELD BUS FAILURE - DRIVE [n0115-in22-05-t1 - 192.168.3.115]"), so the useful text is
    /// the template's HEAD followed by T/U — the tail is filler that the head already earned.
    ///
    /// So the emitted formula truncates the template at its LAST placeholder and appends the
    /// suffix cell. The truncation happens BEFORE substitution and must: once %1/%2 are replaced
    /// the markers are gone, and cutting the rendered text at " - " breaks on every template
    /// whose own tail contains one ("FIELD BUS FAILURE - DRIVE").
    ///
    /// Two rules the shape depends on:
    /// • LET is a post-2019 function, so the FILE must spell it <c>_xlfn.LET</c>. Writing a bare
    ///   modern name into the XML is what made Excel read the previous wholesale rewrite as
    ///   <c>_xludf.CONCAT</c> — an unknown user-defined function — and show #NAME? on every row
    ///   it touched. Everything else here (SUBSTITUTE/FIND/LEFT/LEN/CHAR/IFERROR) is classic and
    ///   takes no prefix. The argument separator is the comma: these strings only ever reach the
    ///   XLSX XML, which is locale-independent, never a locale-aware Excel API.
    /// • CHAR(1) is the sentinel that locates the last "%" — safe because the templates are plain
    ///   ASCII text. A template with NO placeholder makes FIND error, so IFERROR falls back to the
    ///   whole template rather than poisoning the cell.
    /// </summary>
    public static class SwpDescriptionFormula
    {
        /// <summary>The separator between the truncated template and the suffix cell.</summary>
        public const string Separator = " - ";

        /// <summary>What <see cref="Build"/> emits first — also how a built formula is recognized.</summary>
        private const string LetPrefix = "_xlfn.LET(";

        private const string SubstituteName = "SUBSTITUTE";

        /// <summary>
        /// A single cell on another workbook's sheet — "[1]DB_IL_BELT!$K$122", or the quoted form
        /// a spaced sheet name takes. Anything computed is refused: the truncation only means
        /// something applied to a template cell.
        /// </summary>
        private static readonly Regex TemplateReference = new Regex(
            @"^(?:\[\d+\])?[^&+*/(),]*![$]?[A-Za-z]{1,3}[$]?[0-9]{1,7}$",
            RegexOptions.CultureInvariant);

        /// <summary>The pieces of a "template with substituted placeholders" Q/R formula.</summary>
        public sealed class Parsed
        {
            /// <summary>The CONFIG_PLANT template cell, e.g. "[1]DB_IL_BELT!$K$122".</summary>
            public string TemplateRef = "";

            /// <summary>What "%1" is replaced with, e.g. "$B$6".</summary>
            public string Param1 = "";

            /// <summary>What "%2" is replaced with, e.g. "F679" ("" when the template has one placeholder).</summary>
            public string Param2 = "";
        }

        /// <summary>
        /// LET's parameter names as the FILE must spell them. Inside LET, the workbook loader reads
        /// a bare name as an unknown reference and REMOVES the whole formula on open ("Removed
        /// Records: Formula from /xl/worksheets/sheetN.xml part"), so only the cached value
        /// survives — measured on a real Excel: <c>_xlfn.LET(x,1,x+1)</c> refuses to open,
        /// <c>_xlfn.LET(_xlpm.x,1,_xlpm.x+1)</c> loads and reads back as <c>=LET(x,1,x+1)</c>.
        /// </summary>
        private const string ParamTemplate = "_xlpm.t";
        private const string ParamPosition = "_xlpm.p";

        /// <summary>What the first build emitted — bare parameter names, which Excel deletes on
        /// open. Recognized so the next build replaces it rather than leaving it in place.</summary>
        private const string LegacyLetPrefix = "_xlfn.LET(t,";
        private const string LegacyHead = "LEFT(t,IFERROR(p,LEN(t))+1)";

        /// <summary>True for a formula this class built, in either spelling — what the spine, the
        /// preview and the patcher's cache classify as a rebuild.</summary>
        public static bool IsTruncatingForm(string formula) =>
            (formula ?? "").TrimStart().StartsWith(LetPrefix, StringComparison.Ordinal);

        /// <summary>True only for the LOADABLE spelling — the one a build leaves alone.</summary>
        public static bool IsCurrentForm(string formula) =>
            (formula ?? "").TrimStart().StartsWith(LetPrefix + ParamTemplate + ",", StringComparison.Ordinal);

        /// <summary>
        /// The formula: the template truncated at its last placeholder, substituted, then
        /// " - " and the suffix cell.
        /// </summary>
        public static string Build(Parsed parsed, string suffixCell)
        {
            if (parsed == null) throw new ArgumentNullException(nameof(parsed));

            // p = the position of the LAST "%": swap that one occurrence for CHAR(1) (the count
            // of "%" is LEN(t) - LEN(t without "%")), then find it. LEFT(t, p+1) keeps the "%"
            // AND its digit, so the head ends exactly at the placeholder.
            string t = ParamTemplate, p = ParamPosition;
            var text = new StringBuilder();
            text.Append(LetPrefix).Append(t).Append(',').Append(parsed.TemplateRef).Append(',').Append(p).Append(',')
                .Append("FIND(CHAR(1),SUBSTITUTE(").Append(t).Append(",\"%\",CHAR(1),LEN(").Append(t)
                .Append(")-LEN(SUBSTITUTE(").Append(t).Append(",\"%\",\"\")))),");

            string head = "LEFT(" + t + ",IFERROR(" + p + ",LEN(" + t + "))+1)";
            string substituted = SubstituteCall(head, "%1", parsed.Param1);
            if (parsed.Param2.Length > 0) substituted = SubstituteCall(substituted, "%2", parsed.Param2);

            text.Append(substituted).Append("&\"").Append(Separator).Append("\"&").Append(suffixCell).Append(')');
            return text.ToString();
        }

        private static string SubstituteCall(string inner, string marker, string replacement) =>
            SubstituteName + "(" + inner + ",\"" + marker + "\"," + replacement + ")";

        /// <summary>
        /// Recovers the template reference and its placeholder sources from a Q/R formula.
        /// Accepts the plain SWP shape <c>SUBSTITUTE(SUBSTITUTE(tpl,"%1",p1),"%2",p2)</c>, the
        /// single-placeholder shape, either nesting order, and any of them carrying a trailing
        /// concatenation (the <c>&amp;" "&amp;T{row}</c> label extension) — that tail is dropped,
        /// because the built formula supersedes it. Also accepts the first build's unloadable
        /// LET (bare parameter names), so a file still holding it gets the loadable one.
        ///
        /// False when the template is not recoverable: a wholesale-rewritten cell (its reference
        /// was destroyed), a literal, a formula already in the current shape. Those rows are left
        /// alone — the reference-SWP formula audit is what restores a lost template.
        /// </summary>
        public static bool TryParse(string formula, out Parsed parsed)
        {
            parsed = null;
            string text = (formula ?? "").Trim();
            if (text.Length == 0 || IsCurrentForm(text)) return false;

            // The unloadable first build: its template is LET's second argument, and its
            // placeholders wrap the head it always emitted. Put the template back where that head
            // is and it is the plain shape again, parsed like any other. Only this exact head —
            // it has exactly one producer, so nothing here is a guess.
            if (text.StartsWith(LegacyLetPrefix, StringComparison.Ordinal))
            {
                if (!TryReadCall(text, 0, out _, out List<string> letArgs, out _) || letArgs.Count != 5)
                    return false;
                string body = letArgs[4];
                int at = body.IndexOf(LegacyHead, StringComparison.Ordinal);
                if (at < 0 || !TemplateReference.IsMatch(letArgs[1])) return false;
                return TryParse(body.Substring(0, at) + letArgs[1] + body.Substring(at + LegacyHead.Length), out parsed);
            }
            if (IsTruncatingForm(text)) return false;   // a LET this class did not write

            // Only the LEADING call is read; a trailing "&…" extension is deliberately ignored,
            // because the built formula carries the suffix cell itself.
            var result = new Parsed();
            string expression = text;
            int guard = 0;
            while (TryReadCall(expression, 0, out string name, out List<string> args, out _) &&
                   string.Equals(name, SubstituteName, StringComparison.OrdinalIgnoreCase) &&
                   args.Count == 3 && guard++ < 4)
            {
                string marker = Unquote(args[1]);
                if (marker == "%1" && result.Param1.Length == 0) result.Param1 = args[2];
                else if (marker == "%2" && result.Param2.Length == 0) result.Param2 = args[2];
                else return false;
                expression = args[0];
            }

            if (result.Param1.Length == 0) return false;

            // The innermost operand is the template — possibly already concatenated with the
            // suffix cell ("tpl&T216", the legacy template-and-description shape), which the
            // built formula expresses itself. Its FIRST operand must be a plain external
            // reference: a literal head (the pasted-value repair) has no template to truncate,
            // and a computed one is not something to cut blindly.
            int ampersand = expression.IndexOf('&');
            if (ampersand >= 0) expression = expression.Substring(0, ampersand).Trim();
            if (!TemplateReference.IsMatch(expression)) return false;

            result.TemplateRef = expression;
            parsed = result;
            return true;
        }

        /// <summary>
        /// The rendered head — what the truncated template DISPLAYS — recovered from the text the
        /// cell shows today. The head ends at the last placeholder, so it ends with that
        /// placeholder's value: the block's electrical label, or the sheet description when the
        /// template has only "%1". Everything after is the tail the truncation drops.
        ///
        /// The later of the two anchors wins, which is also what makes a label that CONTAINS the
        /// sheet description ("=S1+IN22" inside "=S1+IN22.05") cut in the right place.
        /// </summary>
        public static string TruncatedRendered(string renderedFinal, string sheetDescription, string blockLabel)
        {
            string text = renderedFinal ?? "";
            int cut = HeadEnd(text, sheetDescription, blockLabel);
            return cut < 0 ? text : text.Substring(0, cut);
        }

        /// <summary>Just past the later of the head's two anchors in <paramref name="text"/>; -1 when
        /// neither shows.</summary>
        private static int HeadEnd(string text, string sheetDescription, string blockLabel)
        {
            int cut = -1;
            foreach (string anchor in new[] { blockLabel, sheetDescription })
            {
                if (string.IsNullOrEmpty(anchor)) continue;
                int at = text.LastIndexOf(anchor, StringComparison.Ordinal);
                if (at >= 0 && at + anchor.Length > cut) cut = at + anchor.Length;
            }
            return cut;
        }

        /// <summary>A single A1 cell, absolute markers allowed — the column and row, split.</summary>
        private static readonly Regex CellParts = new Regex(
            @"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$", RegexOptions.CultureInvariant);

        /// <summary>
        /// Whether a formula SHOWS the given cell of its row ("U216", "$U$216") — the one test the
        /// append guard (never extend a formula already showing its description), the preview's
        /// recalc simulation and the kept-formula check below all share. Bounded on both sides, so
        /// AU216 and U2160 are other cells.
        /// </summary>
        public static bool ReferencesCell(string formula, string cell)
        {
            Match parts = CellParts.Match(cell ?? "");
            if (!parts.Success || string.IsNullOrEmpty(formula)) return false;
            return Regex.IsMatch(formula,
                @"(?<![A-Za-z0-9_])\$?" + parts.Groups[1].Value + @"\$?" + parts.Groups[2].Value + @"(?![0-9])",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        /// <summary>
        /// True when a Q/R formula already SHOWS its description cell and CONFIG_PLANT's template in
        /// front of it says nothing past its last placeholder — the CABINET shape
        /// <c>SUBSTITUTE(SUBSTITUTE([1]DB_ILGRP_CABINET!$L$57,"%1",$B$6),"%2",F181)&amp;" "&amp;U216</c>
        /// over a template ending "…%2 - ", which displays
        /// "INTRODUCTION GRP01 - ARMOIRE DE GROUPE =S1+MS1.CC1 -  CIRCUIT BREAKER TRIPPED …". That
        /// cell is then the alarm's ONLY name: keeping CONFIG_PLANT's text keeps the formula, and the
        /// description still belongs in the cell it shows. A formula this class built is that shape
        /// by construction — its head ends at the last placeholder.
        ///
        /// Read from what the cell displays: the cached text, minus the description cell's current
        /// text at its end, cut after the head's anchors — what remains must hold no letter or digit.
        /// Whatever cannot be read that way answers false (no cached text; a cache not ending with the
        /// description, as in a file patched and not yet recalculated; no anchor; another formula
        /// shape), and the row keeps the label-only treatment.
        /// </summary>
        public static bool ShowsSuffixAfterBareTemplate(string formula, string cachedText, string suffixCell,
            string suffixText, string sheetDescription, string blockLabel)
        {
            string text = (formula ?? "").Trim();
            if (!ReferencesCell(text, suffixCell)) return false;
            if (IsTruncatingForm(text)) return true;
            if (!TryParse(text, out _)) return false;

            string shown = cachedText ?? "";
            string suffix = suffixText ?? "";
            if (shown.Length == 0 || !shown.EndsWith(suffix, StringComparison.Ordinal)) return false;
            shown = shown.Substring(0, shown.Length - suffix.Length);

            int cut = HeadEnd(shown, sheetDescription, blockLabel);
            if (cut < 0) return false;
            for (int i = cut; i < shown.Length; i++)
                if (char.IsLetterOrDigit(shown[i])) return false;
            return true;
        }

        /// <summary>The whole displayed text of a built formula: truncated head, separator, suffix.</summary>
        public static string RenderedFinal(string renderedFinal, string sheetDescription,
            string blockLabel, string suffix) =>
            TruncatedRendered(renderedFinal, sheetDescription, blockLabel) + Separator + (suffix ?? "");

        // ---- a minimal formula reader (quotes and nesting, nothing else) ---------------------

        /// <summary>
        /// Reads NAME(arg, arg, …) starting at <paramref name="start"/>. <paramref name="end"/>
        /// is the index just past the closing parenthesis, so a caller can see what follows.
        /// </summary>
        private static bool TryReadCall(string text, int start, out string name, out List<string> args, out int end)
        {
            name = null;
            args = null;
            end = start;

            int i = start;
            while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] == '_' || text[i] == '.')) i++;
            if (i == start || i >= text.Length || text[i] != '(') return false;
            name = text.Substring(start, i - start);

            var parts = new List<string>();
            var current = new StringBuilder();
            int depth = 0;
            bool inQuotes = false;
            for (i++; i < text.Length; i++)
            {
                char c = text[i];
                if (inQuotes)
                {
                    current.Append(c);
                    if (c != '"') continue;
                    // "" inside a quoted literal is an escaped quote, not its end.
                    if (i + 1 < text.Length && text[i + 1] == '"') { current.Append('"'); i++; continue; }
                    inQuotes = false;
                    continue;
                }
                if (c == '"') { inQuotes = true; current.Append(c); continue; }
                if (c == '(') { depth++; current.Append(c); continue; }
                if (c == ')')
                {
                    if (depth == 0)
                    {
                        parts.Add(current.ToString().Trim());
                        args = parts;
                        end = i + 1;
                        return true;
                    }
                    depth--;
                    current.Append(c);
                    continue;
                }
                if (c == ',' && depth == 0) { parts.Add(current.ToString().Trim()); current.Clear(); continue; }
                current.Append(c);
            }
            return false;   // unbalanced — not a formula this reader can speak for
        }

        private static string Unquote(string argument)
        {
            string text = (argument ?? "").Trim();
            return text.Length >= 2 && text[0] == '"' && text[text.Length - 1] == '"'
                ? text.Substring(1, text.Length - 2).Replace("\"\"", "\"")
                : text;
        }
    }
}
