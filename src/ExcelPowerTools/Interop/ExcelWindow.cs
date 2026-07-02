using System;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// Wraps Excel's main-window handle so message boxes and dialogs are owned by Excel
    /// (they appear on top and stay modal to it, instead of getting lost behind the window).
    /// </summary>
    internal sealed class ExcelWindow : IWin32Window
    {
        public IntPtr Handle { get; }

        private ExcelWindow(IntPtr handle) => Handle = handle;

        public static ExcelWindow Current
        {
            get
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    return new ExcelWindow(new IntPtr((int)app.Hwnd));
                }
                catch
                {
                    return new ExcelWindow(IntPtr.Zero);
                }
            }
        }
    }
}
