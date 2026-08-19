using System;
using System.Collections.Generic;
using System.Globalization;

namespace ExcelPowerTools.Core.Experimental.IoBridge
{
    public sealed class GroupSpec
    {
        public string Name = "";
        public int Order;
        public bool HiddenByDefault;
    }

    /// <summary>Per-user column configuration: which group each column belongs to, and each group's traits.</summary>
    public sealed class IoBridgeColumnConfig
    {
        /// <summary>Normalized header → group name. Unlisted columns fall into the implicit trailing group.</summary>
        public Dictionary<string, string> ColumnGroups =
            new Dictionary<string, string>(StringComparer.Ordinal);

        public List<GroupSpec> Groups = new List<GroupSpec>();

        public GroupSpec FindGroup(string name)
        {
            foreach (GroupSpec g in Groups)
                if (g.Name == name) return g;
            return null;
        }
    }

    public sealed class ColumnDef
    {
        public string NormalizedHeader = "";
        public string SampleHeader = "";     // the header as it reads on the sheet
        public string Side = "";             // "IO", "CSV", "Status"
    }

    public sealed class LayoutColumn
    {
        public ColumnDef Def;
        public string Group = "";
    }

    /// <summary>One contiguous run of columns forming an Excel outline group. Indices are 0-based into the layout.</summary>
    public sealed class GroupSpan
    {
        public string Name = "";
        public int FirstIndex, LastIndex;
        public bool Hidden;
    }

    public sealed class ColumnLayout
    {
        public List<LayoutColumn> Columns = new List<LayoutColumn>();
        public List<GroupSpan> Spans = new List<GroupSpan>();
    }

    /// <summary>
    /// Orders workspace columns group-by-group so Excel outline groups (which must be contiguous) can
    /// express the configuration, and round-trips the _Columns configuration sheet. Pure logic — the
    /// COM grouping calls live in the add-in.
    /// </summary>
    public static class IoBridgeColumnLayout
    {
        /// <summary>
        /// Columns are NEVER reordered — the workspace mirrors the source document's column positions
        /// so both can be read side by side. A group therefore materializes as one outline span per
        /// contiguous RUN of columns assigned to it; a name split across the sheet simply yields
        /// several spans sharing that name and its hidden-by-default flag. The group's configured
        /// Order does not move anything.
        /// </summary>
        public static ColumnLayout ComputeLayout(IReadOnlyList<ColumnDef> columns, IoBridgeColumnConfig config)
        {
            if (columns == null) throw new ArgumentNullException(nameof(columns));
            config = config ?? new IoBridgeColumnConfig();

            var layout = new ColumnLayout();
            string runGroup = null;
            int runStart = -1;

            for (int i = 0; i <= columns.Count; i++)
            {
                string group = null;
                if (i < columns.Count)
                {
                    group = config.ColumnGroups.TryGetValue(columns[i].NormalizedHeader, out string g)
                        ? (g ?? "").Trim()
                        : "";
                    layout.Columns.Add(new LayoutColumn { Def = columns[i], Group = group });
                }

                if (group == runGroup && i < columns.Count) continue;

                if (runStart >= 0)
                {
                    GroupSpec spec = runGroup.Length > 0 ? config.FindGroup(runGroup) : null;
                    layout.Spans.Add(new GroupSpan
                    {
                        Name = runGroup,
                        FirstIndex = runStart,
                        LastIndex = i - 1,
                        Hidden = runGroup.Length > 0 && spec != null && spec.HiddenByDefault,
                    });
                }
                runGroup = group;
                runStart = i;
            }
            return layout;
        }

        public static readonly string[] ColumnsSheetHeaders =
            { "Column (normalized)", "Sample header", "Side", "Group", "Group order", "Hidden by default" };

        /// <summary>The _Columns sheet content for the current config and the columns actually present.</summary>
        public static List<string[]> RenderColumnsSheet(IoBridgeColumnConfig config, IReadOnlyList<ColumnDef> knownColumns)
        {
            config = config ?? new IoBridgeColumnConfig();
            var rows = new List<string[]> { (string[])ColumnsSheetHeaders.Clone() };

            foreach (ColumnDef col in knownColumns ?? new List<ColumnDef>())
            {
                config.ColumnGroups.TryGetValue(col.NormalizedHeader, out string group);
                group = (group ?? "").Trim();
                GroupSpec spec = group.Length > 0 ? config.FindGroup(group) : null;

                rows.Add(new[]
                {
                    col.NormalizedHeader,
                    col.SampleHeader,
                    col.Side,
                    group,
                    spec != null ? spec.Order.ToString(CultureInfo.InvariantCulture) : "",
                    spec != null && spec.HiddenByDefault ? "1" : "",
                });
            }
            return rows;
        }

        /// <summary>
        /// Reads an edited _Columns sheet back into a config. The first row carrying a group defines
        /// that group's order and hidden flag; later rows only need the group name. Rows with an empty
        /// column key are ignored, so stray notes below the table are harmless.
        /// </summary>
        public static IoBridgeColumnConfig ParseColumnsSheet(IReadOnlyList<string[]> rows)
        {
            var config = new IoBridgeColumnConfig();
            if (rows == null) return config;

            bool headerSkipped = false;
            foreach (string[] row in rows)
            {
                if (row == null || row.Length == 0) continue;
                string key = Cell(row, 0).Trim();
                if (key.Length == 0) continue;
                if (!headerSkipped && key == ColumnsSheetHeaders[0]) { headerSkipped = true; continue; }

                string group = Cell(row, 3).Trim();
                config.ColumnGroups[key] = group;

                if (group.Length == 0 || config.FindGroup(group) != null) continue;
                config.Groups.Add(new GroupSpec
                {
                    Name = group,
                    Order = ParseOrder(Cell(row, 4)),
                    HiddenByDefault = ParseFlag(Cell(row, 5)),
                });
            }
            return config;
        }

        private static string Cell(string[] row, int index) =>
            index < row.Length && row[index] != null ? row[index] : "";

        private static int ParseOrder(string text) =>
            int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : 0;

        private static bool ParseFlag(string text)
        {
            switch (text.Trim().ToLowerInvariant())
            {
                case "1":
                case "x":
                case "yes":
                case "true":
                case "☑":
                    return true;
                default:
                    return false;
            }
        }
    }
}
