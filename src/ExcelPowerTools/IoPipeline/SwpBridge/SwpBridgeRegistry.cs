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

        /// <summary>The session FILE Setup last built or saved — what it comes back to, since a
        /// named session lives wherever it was saved and a root only knows its default sidecar.</summary>
        public static string GetLastSession()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
                    return key?.GetValue("LastSession") as string;
            }
            catch { return null; }
        }

        public static void SetLastSession(string path)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                    key?.SetValue("LastSession", path ?? "");
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

        /// <summary>
        /// The MNEMONIC MAP — what a mnemonic is, where it lands, which producers run, and the
        /// template sets. Same shape as the index rules: stored per user so it applies to every
        /// system and session, seeded from the shipped table when nothing is stored yet.
        ///
        /// A stored table that no longer parses would be worse than no table at all, so a parse
        /// producing nothing usable falls back to the shipped default rather than leaving the
        /// bridge with an empty map that silently claims no mnemonics.
        /// </summary>
        public static SwpMnemonicMap GetMnemonicMap()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RootKey))
                {
                    if (key?.GetValue("MnemonicMap") is string[] lines && lines.Length > 0)
                    {
                        SwpMnemonicMap stored = SwpMnemonicMap.Parse(lines, null);
                        if (stored != null && stored.Rules.Count > 0) return stored;
                    }
                }
            }
            catch { }
            return SwpMnemonicMapDefaults.Official();
        }

        public static void SetMnemonicMap(SwpMnemonicMap map)
        {
            try
            {
                List<string> lines = map?.ToLines();
                if (lines == null || lines.Count <= 1) return;   // header only = never persist
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                    key?.SetValue("MnemonicMap", lines.ToArray(), RegistryValueKind.MultiString);
            }
            catch { }
        }

        /// <summary>Drops the stored table so the next read returns the shipped one — what the
        /// sheet's "restore defaults" needs, and the way out of a table edited into a corner.</summary>
        public static void ClearMnemonicMap()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RootKey))
                    key?.DeleteValue("MnemonicMap", throwOnMissingValue: false);
            }
            catch { }
        }
    }
}
