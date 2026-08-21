using System;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The one definition of "this SWP cell already holds that value", shared by the diff (decides
    /// what to propose) and the patcher (three-way check) so they can never disagree. Values are
    /// compared trimmed — legacy-written files carry trailing spaces — and the Export flag is
    /// case-insensitive (the fleet's legacy "x" equals the rules-§3.1 "X" we write now).
    /// </summary>
    public static class SwpFieldCompare
    {
        public static bool Same(SwpField field, string a, string b)
        {
            a = (a ?? "").Trim();
            b = (b ?? "").Trim();
            return string.Equals(a, b,
                field == SwpField.ExportB ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
        }
    }
}
