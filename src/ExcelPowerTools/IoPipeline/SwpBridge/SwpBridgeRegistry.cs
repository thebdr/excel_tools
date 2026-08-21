using System.Collections.Generic;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;
using Microsoft.Win32;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// HKCU persistence for what is per-user rather than per-system: the last SWP root the setup
    /// form opened, the INDEX RULES table (per plant type, per diagnosable object — edited on
    /// the _Offsets sheet), and the LANGUAGE table. The language table has no sheet any more:
    /// it quietly supplies the bus-failure constants for GLOSSARY-LESS workspaces and maps
    /// legacy sidecar language names to codes. Everything project-shaped (paths, parameters,
    /// overrides, dictionaries) lives in the system's swp-bridge.json.
    /// </summary>
    internal static class SwpBridgeRegistry
    {
        private const string RootKey = @"Software\ExcelPowerTools\SwpBridge";
        private const char LanguageSeparator = '|';

        public static string GetLastRoot()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
                    return key?.GetValue("LastRoot") as string;
            }
            catch { return null; }
        }

        public static void SetLastRoot(string path)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                    key?.SetValue("LastRoot", path ?? "");
            }
            catch { }
        }

        /// <summary>
        /// The ordered language table: (name, glossary code, bus-failure text). The FIRST row is
        /// language 1 — the language of the SWP's Description 1 column; the CODE is the column
        /// header the glossario workbooks use (EN/IT/FR/PT/…). Seeded with the texts the legacy
        /// tools carried when nothing is stored yet.
        /// </summary>
        public static List<(string Name, string Code, string Text)> GetLanguages()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
                {
                    if (key?.GetValue("Languages") is string[] lines && lines.Length > 0)
                    {
                        var languages = new List<(string, string, string)>();
                        foreach (string line in lines)
                        {
                            string[] parts = (line ?? "").Split(LanguageSeparator);
                            if (parts.Length >= 3 && parts[0].Length > 0)
                                languages.Add((parts[0], parts[1],
                                    string.Join(LanguageSeparator.ToString(), parts, 2, parts.Length - 2)));
                            else if (parts.Length == 2 && parts[0].Length > 0)
                                languages.Add((parts[0], "", parts[1]));   // pre-code format
                        }
                        if (languages.Count > 0) return languages;
                    }
                }
            }
            catch { }
            return DefaultLanguages();
        }

        public static void SetLanguages(List<(string Name, string Code, string Text)> languages)
        {
            try
            {
                var lines = new List<string>();
                foreach ((string name, string code, string text) in languages)
                    if (!string.IsNullOrWhiteSpace(name))
                        lines.Add(name.Trim() + LanguageSeparator + (code ?? "").Trim().ToUpperInvariant() +
                                  LanguageSeparator + (text ?? ""));
                if (lines.Count == 0) return;   // never persist an empty table
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                    key?.SetValue("Languages", lines.ToArray(), RegistryValueKind.MultiString);
            }
            catch { }
        }

        public static List<(string Name, string Code, string Text)> DefaultLanguages() =>
            new List<(string, string, string)>
            {
                ("English", "EN", "FIELD BUS MODULE FAILURE"),
                ("Português", "PT", "AVARIA DO MÓDULO BUS DE CAMPO"),
            };

        /// <summary>
        /// The index-rules table (per plant type + diagnosable object), seeded from the official
        /// "SWP and diagnostic configuration rules" tables when nothing is stored yet.
        /// </summary>
        public static SwpIndexRules GetIndexRules()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
                {
                    if (key?.GetValue("IndexRules") is string[] lines && lines.Length > 0)
                    {
                        SwpIndexRules stored = SwpIndexRules.FromLines(lines);
                        if (stored != null) return stored;
                    }
                }
            }
            catch { }
            return SwpIndexRules.Official();
        }

        public static void SetIndexRules(SwpIndexRules rules)
        {
            try
            {
                List<string> lines = rules?.ToLines();
                if (lines == null || lines.Count == 0) return;   // never persist an empty table
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                    key?.SetValue("IndexRules", lines.ToArray(), RegistryValueKind.MultiString);
            }
            catch { }
        }
    }
}
