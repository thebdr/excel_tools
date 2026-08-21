using System;
using System.Reflection;
using System.Windows.Forms;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Text rendering for the add-in's WinForms surfaces. WinForms draws Label / CheckBox / Radio /
    /// Button text through GDI+ unless the HOST called Application.SetCompatibleTextRenderingDefault
    /// (false) before any window existed — an application does that in Main; an add-in inside Excel
    /// never gets the chance, so the default stays GDI+. Inside Excel's process that GDI+ text comes
    /// out un-antialiased (jagged) while the native controls (TextBox = EDIT) draw smooth GDI
    /// ClearType. Measured 2026-08-21: the very same controls in a standalone process render
    /// ClearType under every DPI awareness context, so the remedy is to take the GDI path
    /// (TextRenderer) like the native controls do. <see cref="MakeDefault"/> flips the add-in-wide
    /// default — the public API first; when a window already exists in the AppDomain (Excel-DNA's
    /// synchronization window) through the internal setter. <see cref="Apply"/> is the per-surface
    /// net for controls that read the default before it changed.
    /// </summary>
    internal static class GdiText
    {
        private static bool _defaultApplied;

        public static void MakeDefault()
        {
            if (_defaultApplied) return;
            try
            {
                Application.SetCompatibleTextRenderingDefault(false);
                _defaultApplied = true;
                return;
            }
            catch (InvalidOperationException) { }   // a handle exists already — use the internal setter
            catch { }

            try
            {
                PropertyInfo property = typeof(Control).GetProperty("UseCompatibleTextRenderingDefault",
                    BindingFlags.NonPublic | BindingFlags.Static);
                if (property != null && property.CanWrite)
                {
                    property.SetValue(null, false, null);
                    _defaultApplied = true;
                }
            }
            catch { }
        }

        /// <summary>GDI text for every control under root — the ones there now and the ones added later.</summary>
        public static void Apply(Control root)
        {
            if (root == null) return;
            ApplyOne(root);
            foreach (Control child in root.Controls) Apply(child);
            root.ControlAdded -= OnControlAdded;
            root.ControlAdded += OnControlAdded;
        }

        private static void OnControlAdded(object sender, ControlEventArgs e) => Apply(e.Control);

        private static void ApplyOne(Control control)
        {
            switch (control)
            {
                case Label label: label.UseCompatibleTextRendering = false; break;          // LinkLabel included
                case ButtonBase button: button.UseCompatibleTextRendering = false; break;   // Button, CheckBox, RadioButton
                case GroupBox group: group.UseCompatibleTextRendering = false; break;
            }
        }
    }
}
