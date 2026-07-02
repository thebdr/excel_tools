using Microsoft.Win32;

namespace ExcelPowerTools.Settings
{
    /// <summary>
    /// The Change-Field dialog settings. The VBA version stashed these on a hidden "DataStorage"
    /// sheet inside PERSONAL.XLSB; an add-in has no workbook, so we persist them per-user in the
    /// registry (HKCU\Software\ExcelPowerTools\FieldEditor).
    /// </summary>
    public sealed class FieldEditorSettings
    {
        private const string KeyPath = @"Software\ExcelPowerTools\FieldEditor";

        public string Separator { get; set; } = ".";
        public string Separator2 { get; set; } = "";
        public int FieldIndex { get; set; } = 1;      // 1-based, as shown in the dialog
        public string FontColor { get; set; } = "";   // "", "auto", or RRGGBB hex
        public string InteriorColor { get; set; } = ""; // "", "none", or RRGGBB hex

        public static FieldEditorSettings Load()
        {
            var s = new FieldEditorSettings();
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(KeyPath))
            {
                if (key == null) return s;

                s.Separator = (key.GetValue(nameof(Separator)) as string) ?? s.Separator;
                s.Separator2 = (key.GetValue(nameof(Separator2)) as string) ?? s.Separator2;
                s.FontColor = (key.GetValue(nameof(FontColor)) as string) ?? s.FontColor;
                s.InteriorColor = (key.GetValue(nameof(InteriorColor)) as string) ?? s.InteriorColor;

                if (int.TryParse(key.GetValue(nameof(FieldIndex)) as string, out int idx))
                    s.FieldIndex = idx;
            }
            return s;
        }

        public void Save()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(KeyPath))
            {
                if (key == null) return;
                key.SetValue(nameof(Separator), Separator ?? "");
                key.SetValue(nameof(Separator2), Separator2 ?? "");
                key.SetValue(nameof(FieldIndex), FieldIndex.ToString());
                key.SetValue(nameof(FontColor), FontColor ?? "");
                key.SetValue(nameof(InteriorColor), InteriorColor ?? "");
            }
        }
    }
}
