using System;

namespace ExcelPowerTools.Interop
{
    /// <summary>Broadcast of Excel selection changes (wired in AddIn) for live UI like the stats panel.</summary>
    internal static class SelectionEvents
    {
        public static event Action Changed;

        public static void Raise()
        {
            try { Changed?.Invoke(); } catch { }
        }
    }
}
