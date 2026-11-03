using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelPowerTools.Core.IoPipeline.SwpBridge
{
    /// <summary>
    /// The workspace's first-line and group numbers, read from the SWP's OWN sheet names instead
    /// of being typed. The mnemonic table builds destination sheets as
    /// "IL{firstLineId+line-1:00}_…" and "ILGRP{groupIndex:00}_CABINET" (UPS: "UP{firstLineId+sys-1:00}_…"),
    /// so the numbers are already written on the tabs — asking for them again is asking the user
    /// to re-derive what the file states, and getting one wrong sends every row to a sheet that
    /// does not exist.
    ///
    /// The FIRST line is the LOWEST numbered one: the mnemonic's line digit starts at 1 and the
    /// map computes firstLineId + line - 1. The group is whatever ILGRP sheet is present.
    ///
    /// UPS carries no group in its tab names — the conveyor group (ConveyorSys\SysGroupNN) is a
    /// production-tree fact, not an SWP one — so <see cref="Result.HasGroup"/> stays false there
    /// and the typed value stands.
    /// </summary>
    public static class SwpPlantNumbering
    {
        private static readonly Regex IspcLine = new Regex(@"^IL(\d{1,2})_", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex IspcGroup = new Regex(@"^ILGRP(\d{1,2})_", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex UpsSystem = new Regex(@"^UP(\d{1,2})_", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public struct Result
        {
            public bool HasFirstLine;
            public int FirstLineId;
            public bool HasGroup;
            public int GroupIndex;

            /// <summary>The sheet names the numbers were read from — what the dialog shows so the
            /// user can see WHY, instead of a number appearing on its own.</summary>
            public string Evidence;

            public bool Any => HasFirstLine || HasGroup;
        }

        /// <summary>
        /// Reads the numbers a plant type keeps in its tab names. MSC and SAFETY name no line or
        /// group at all (one cabinet table, no per-line sheets), so nothing is derived for them
        /// and the dialog's own values stand.
        /// </summary>
        public static Result FromSheetNames(IEnumerable<string> sheetNames, SwpPlantType plantType)
        {
            var result = new Result { Evidence = "" };
            if (sheetNames == null) return result;
            if (plantType != SwpPlantType.Ispc && plantType != SwpPlantType.Ups) return result;

            int firstLine = int.MaxValue;
            string firstLineSheet = null;
            int group = int.MaxValue;
            string groupSheet = null;

            foreach (string name in sheetNames)
            {
                if (string.IsNullOrEmpty(name)) continue;

                // ILGRP is tested FIRST: "ILGRP02_CABINET" also starts with "IL", and reading its
                // "GR" as a line number is exactly the kind of silent misparse this replaces.
                if (plantType == SwpPlantType.Ispc)
                {
                    Match groupMatch = IspcGroup.Match(name);
                    if (groupMatch.Success)
                    {
                        int value = Number(groupMatch);
                        if (value < group) { group = value; groupSheet = name; }
                        continue;
                    }
                }

                Match lineMatch = plantType == SwpPlantType.Ispc ? IspcLine.Match(name) : UpsSystem.Match(name);
                if (!lineMatch.Success) continue;
                int line = Number(lineMatch);
                if (line < firstLine) { firstLine = line; firstLineSheet = name; }
            }

            var evidence = new List<string>();
            if (firstLineSheet != null)
            {
                result.HasFirstLine = true;
                result.FirstLineId = firstLine;
                evidence.Add(firstLineSheet);
            }
            if (groupSheet != null)
            {
                result.HasGroup = true;
                result.GroupIndex = group;
                evidence.Add(groupSheet);
            }
            result.Evidence = string.Join(", ", evidence.ToArray());
            return result;
        }

        private static int Number(Match match) =>
            int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
    }
}
