using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelPowerTools.Core.IoPipeline.IoBridge;
using Microsoft.Win32;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// Per-user persistence in HKCU: last-used paths for the setup form, and the column-group
    /// configuration keyed by normalized header text (so it survives cosmetic header edits across
    /// IO-list revisions and applies to every project).
    /// </summary>
    internal static class IoBridgeRegistry
    {
        // Historical key name kept verbatim so settings saved before the IO Pipeline rename survive.
        private const string RootKey = @"Software\ExcelPowerTools\Experimental\IoBridge";
        private const string ColumnsKey = RootKey + @"\Columns";
        private const string GroupsKey = RootKey + @"\Groups";

        // ---- setup-form memory -----------------------------------------------------------------------

        public static string GetPath(string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
                return key?.GetValue(name) as string;
        }

        public static void SetPath(string name, string value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                key?.SetValue(name, value ?? "");
        }

        // ---- column configuration --------------------------------------------------------------------

        public static IoBridgeColumnConfig LoadColumnConfig()
        {
            var config = new IoBridgeColumnConfig();
            try
            {
                using (RegistryKey columns = Registry.CurrentUser.OpenSubKey(ColumnsKey))
                {
                    if (columns != null)
                        foreach (string valueName in columns.GetValueNames())
                        {
                            string group = columns.GetValue(valueName) as string;
                            if (!string.IsNullOrEmpty(valueName))
                                config.ColumnGroups[valueName] = (group ?? "").Trim();
                        }
                }

                using (RegistryKey groups = Registry.CurrentUser.OpenSubKey(GroupsKey))
                {
                    if (groups != null)
                        foreach (string groupName in groups.GetSubKeyNames())
                            using (RegistryKey one = groups.OpenSubKey(groupName))
                            {
                                if (one == null) continue;
                                config.Groups.Add(new GroupSpec
                                {
                                    Name = groupName,
                                    Order = ParseInt(one.GetValue("Order") as string),
                                    HiddenByDefault = (one.GetValue("Hidden") as string) == "1",
                                });
                            }
                }
            }
            catch
            {
                // A broken registry never blocks a build — worst case every column is visible.
            }
            return config;
        }

        public static void SaveColumnConfig(IoBridgeColumnConfig config)
        {
            if (config == null) return;

            try { Registry.CurrentUser.DeleteSubKeyTree(ColumnsKey, throwOnMissingSubKey: false); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(GroupsKey, throwOnMissingSubKey: false); } catch { }

            using (RegistryKey columns = Registry.CurrentUser.CreateSubKey(ColumnsKey))
            {
                if (columns != null)
                    foreach (KeyValuePair<string, string> pair in config.ColumnGroups)
                        if (!string.IsNullOrEmpty(pair.Key) && !string.IsNullOrEmpty(pair.Value))
                            columns.SetValue(pair.Key, pair.Value);
            }

            foreach (GroupSpec group in config.Groups)
            {
                // A backslash would nest registry keys and shear the round trip.
                if (string.IsNullOrEmpty(group.Name) || group.Name.IndexOf('\\') >= 0) continue;
                using (RegistryKey one = Registry.CurrentUser.CreateSubKey(GroupsKey + @"\" + group.Name))
                {
                    if (one == null) continue;
                    one.SetValue("Order", group.Order.ToString(CultureInfo.InvariantCulture));
                    one.SetValue("Hidden", group.HiddenByDefault ? "1" : "0");
                }
            }
        }

        private static int ParseInt(string text) =>
            int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : 0;
    }
}
