using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace ExcelPowerTools.Core.SpreadsheetReview
{
    /// <summary>Resolves ordinary saved manual style evidence. Does not evaluate conditional formatting.</summary>
    internal sealed class SavedStyleReader
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private readonly XElement[] formats, bases, fills, fonts;
        private readonly string[] theme;
        private readonly string[] indexed;
        private readonly bool absentStylePart;
        private readonly Dictionary<int, ChannelEvidence<CellInterior>> interiors = new Dictionary<int, ChannelEvidence<CellInterior>>();
        private static readonly string[] Palette = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
            "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
            "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCCC 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
            "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
            "003366 339966 003300 333300 993300 993366 333399 333333").Split(' ');

        internal SavedStyleReader(XElement styles, XElement themePart)
        {
            if (styles != null && styles.Name != S + "styleSheet") throw new InvalidOperationException("Unexpected saved style-table namespace/root.");
            absentStylePart = styles == null;
            formats = styles?.Element(S + "cellXfs")?.Elements(S + "xf").ToArray() ?? new XElement[0];
            bases = styles?.Element(S + "cellStyleXfs")?.Elements(S + "xf").ToArray() ?? new XElement[0];
            fills = styles?.Element(S + "fills")?.Elements(S + "fill").ToArray() ?? new XElement[0];
            fonts = styles?.Element(S + "fonts")?.Elements(S + "font").ToArray() ?? new XElement[0];
            indexed = styles?.Element(S + "colors")?.Element(S + "indexedColors")?.Elements(S + "rgbColor").Select(c => (string)c.Attribute("rgb")).ToArray();
            string[] names = { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
            var scheme = themePart?.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
            theme = names.Select(n =>
            {
                var color = scheme?.Elements().FirstOrDefault(e => e.Name.LocalName == n)?.Elements().FirstOrDefault();
                return color == null ? null : (string)color.Attribute(color.Name.LocalName == "sysClr" ? "lastClr" : "val");
            }).ToArray();
        }

        internal CellObservation Observe(CellContent content, int style, XElement rich)
        {
            ChannelEvidence<CellInterior> interior;
            lock (interiors)
            {
                if (!interiors.TryGetValue(style, out interior))
                {
                    try { interior = ChannelEvidence<CellInterior>.Observed(Interior(style)); }
                    catch (Exception ex) when (ex is ArgumentException || ex is FormatException || ex is InvalidOperationException)
                    { interior = ChannelEvidence<CellInterior>.Unavailable(ex.Message); }
                    interiors.Add(style, interior);
                }
            }
            ChannelEvidence<StrikeProfile> strike;
            try
            {
                bool cellStrike = false;
                if (!absentStylePart)
                {
                    int fontId = Component(style, "fontId", "applyFont");
                    if (fontId < 0 || fontId >= fonts.Length) throw new InvalidOperationException("Saved font index is outside the style table.");
                    cellStrike = On(fonts[fontId].Element(S + "strike"));
                }
                else if (style != 0) throw new InvalidOperationException("The cell references a missing style table.");
                var runs = new List<StrikeRun>(); int start = 0;
                if (rich != null)
                {
                    foreach (var element in rich.Elements().Where(e => e.Name == S + "t" || e.Name == S + "r"))
                    {
                        string text = SavedWorkbookReader.RichElementText(element);
                        var properties = element.Element(S + "rPr");
                        // A run's explicit strike overrides the cell. Unspecified strike inherits the cell font.
                        var localStrike = properties?.Element(S + "strike");
                        if (localStrike != null && text.Length > 0) runs.Add(new StrikeRun(start, text.Length, On(localStrike)));
                        start += text.Length;
                    }
                }
                strike = ChannelEvidence<StrikeProfile>.Observed(new StrikeProfile(content.ToString(), cellStrike, runs));
            }
            catch (Exception ex) when (ex is ArgumentException || ex is FormatException || ex is InvalidOperationException)
            { strike = ChannelEvidence<StrikeProfile>.Unavailable(ex.Message); }
            return CellObservation.Observed(content, interior, strike, FormattingOrigin.SavedManual);
        }

        private int Component(int style, string component, string apply)
        {
            if (style < 0 || style >= formats.Length) throw new InvalidOperationException("Saved cell style index is outside the style table.");
            var xf = formats[style]; int baseId = Int(xf.Attribute("xfId"), 0);
            var baseXf = baseId >= 0 && baseId < bases.Length ? bases[baseId] : null;
            if (baseXf == null && baseId != 0) throw new InvalidOperationException("The named-style parent is unavailable.");
            bool inherit = (string)xf.Attribute(apply) == "0" || (string)xf.Attribute(apply) == "false";
            return Int((inherit ? baseXf?.Attribute(component) : xf.Attribute(component)) ?? baseXf?.Attribute(component), 0);
        }
        private CellInterior Interior(int style)
        {
            if (absentStylePart && style == 0) return CellInterior.NoFill;
            int fillId = Component(style, "fillId", "applyFill");
            if (fillId < 0 || fillId >= fills.Length) throw new InvalidOperationException("Saved fill index is outside the style table.");
            var pattern = fills[fillId].Element(S + "patternFill");
            if (pattern == null) throw new InvalidOperationException("Gradient or unsupported interior evidence requires the full style reader.");
            string kind = (string)pattern.Attribute("patternType") ?? "none";
            if (kind == "none") return CellInterior.NoFill;
            string fg = Color(pattern.Element(S + "fgColor"));
            // For a solid fill the background is not visible and must not create a false positive.
            string bg = kind == "solid" ? null : Color(pattern.Element(S + "bgColor"));
            return new CellInterior(kind, fg, bg);
        }
        private string Color(XElement color)
        {
            if (color == null) return null;
            string rgb = (string)color.Attribute("rgb");
            if (rgb == null && color.Attribute("theme") != null)
            {
                int i = Int(color.Attribute("theme"), -1);
                if (i < 0 || i >= theme.Length || theme[i] == null) throw new InvalidOperationException("Theme color cannot be resolved from the saved theme.");
                rgb = theme[i];
            }
            if (rgb == null && color.Attribute("indexed") != null)
            {
                int i = Int(color.Attribute("indexed"), -1);
                if (i < 0 || i >= (indexed?.Length ?? Palette.Length)) throw new InvalidOperationException("System or missing indexed interior color is not a fixed saved color.");
                rgb = indexed == null ? Palette[i] : indexed[i];
            }
            if (rgb == null) throw new InvalidOperationException("Automatic interior color has no resolved saved RGB value.");
            if (rgb.Length == 6) rgb = "FF" + rgb;
            if (rgb.Length != 8 || rgb.Any(c => !Uri.IsHexDigit(c))) throw new FormatException("Invalid saved RGB color.");
            double tint = color.Attribute("tint") == null ? 0 : double.Parse((string)color.Attribute("tint"), CultureInfo.InvariantCulture);
            if (tint != 0)
                throw new InvalidOperationException("Tinted colors are retained as unavailable until Excel-compatible tint rounding is verified; no guessed RGB is used.");
            return rgb.ToUpperInvariant();
        }
        private static bool On(XElement element) => element != null && (string)element.Attribute("val") != "0" && (string)element.Attribute("val") != "false";
        private static int Int(XAttribute value, int fallback) => value == null ? fallback : int.Parse(value.Value, NumberStyles.None, CultureInfo.InvariantCulture);
    }
}
