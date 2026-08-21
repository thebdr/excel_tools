using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    /// <summary>
    /// Cell right-click menu entries for the bridge, installed while a session is live and removed on
    /// teardown. Buttons are Temporary (Excel drops them at exit) and tagged, so re-installation never
    /// duplicates and removal never touches anything else.
    /// </summary>
    internal static class IoBridgeContextMenu
    {
        private const int MsoControlButton = 1;
        private const string Tag = "PT_IOB_CTX";

        private static readonly (string Caption, string Macro)[] Buttons =
        {
            // The macro names must match the [ExcelCommand(Name = …)] attributes in IoBridgeCommands —
            // the IO Pipeline rename dropped the PT_X prefix and these stayed behind.
            ("IO Bridge: row actions…", "PT_IoBridgeRowActions"),
            ("IO Bridge: expanded view", "PT_IoBridgeExpand"),
        };

        public static void Install()
        {
            try
            {
                dynamic bar = ExcelApp.App.CommandBars["Cell"];
                RemoveFrom(bar);

                bool first = true;
                foreach ((string caption, string macro) in Buttons)
                {
                    dynamic button = bar.Controls.Add(MsoControlButton, System.Type.Missing,
                        System.Type.Missing, System.Type.Missing, true /* Temporary */);
                    button.Caption = caption;
                    button.Tag = Tag;
                    button.OnAction = macro;
                    if (first) { button.BeginGroup = true; first = false; }
                }
            }
            catch { /* a locked-down CommandBars model must never break the session */ }
        }

        public static void Remove()
        {
            try
            {
                RemoveFrom(ExcelApp.App.CommandBars["Cell"]);
            }
            catch { }
        }

        private static void RemoveFrom(dynamic bar)
        {
            // Walk backwards — deleting re-indexes the collection.
            for (int i = (int)bar.Controls.Count; i >= 1; i--)
            {
                dynamic control = bar.Controls(i);
                string tag = null;
                try { tag = (string)control.Tag; } catch { }
                if (tag == Tag) try { control.Delete(); } catch { }
            }
        }
    }
}
