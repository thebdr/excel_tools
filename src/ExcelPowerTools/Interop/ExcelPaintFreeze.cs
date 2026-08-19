using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// Freezes ALL painting of the Excel main window for the lifetime of the scope — the WM_SETREDRAW
    /// trick. Application.ScreenUpdating suppresses cell repaints but NOT window creation, close, or
    /// z-order churn (Workbooks.Add/Open/Close), which is exactly the flashing a workspace rebuild
    /// produces. Dispose re-enables redraw and forces one full repaint, so the UI goes from "old"
    /// to "finished" in a single frame. Never hold across user dialogs.
    /// </summary>
    internal sealed class ExcelPaintFreeze : IDisposable
    {
        private const int WM_SETREDRAW = 0x000B;
        private const uint RDW_INVALIDATE = 0x0001;
        private const uint RDW_ERASE = 0x0004;
        private const uint RDW_ALLCHILDREN = 0x0080;
        private const uint RDW_UPDATENOW = 0x0100;
        private const uint RDW_FRAME = 0x0400;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        private readonly IntPtr _hwnd;
        private bool _frozen;

        public ExcelPaintFreeze()
        {
            try
            {
                _hwnd = ExcelDnaUtil.WindowHandle;
                if (_hwnd != IntPtr.Zero)
                {
                    SendMessage(_hwnd, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
                    _frozen = true;
                }
            }
            catch { /* a failed freeze degrades to the old flicker, never to a broken build */ }
        }

        public void Dispose()
        {
            if (!_frozen) return;
            _frozen = false;
            try
            {
                SendMessage(_hwnd, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
                RedrawWindow(_hwnd, IntPtr.Zero, IntPtr.Zero,
                    RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW | RDW_FRAME);
            }
            catch { }
        }
    }
}
