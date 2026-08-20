using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace ExcelPowerTools.Interop
{
    /// <summary>
    /// An opaque "working…" curtain over the screen hosting Excel, alive for the scope of a rebuild.
    ///
    /// Why this exists: modern Excel is SDI — Workbooks.Add/Open create brand-new TOP-LEVEL windows,
    /// which neither Application.ScreenUpdating nor a WM_SETREDRAW freeze on the existing window can
    /// suppress. Their creation, painting and closing is the flashing that survives every in-process
    /// trick. The curtain simply covers it: whatever Excel paints happens behind an opaque topmost
    /// window carrying a message and a marquee.
    ///
    /// The form lives on its OWN STA thread with its own message pump, because the Excel thread is
    /// blocked inside the build — an overlay on that thread would freeze unpainted. It never takes
    /// focus (WS_EX_NOACTIVATE), so Excel's activation state stays untouched. Every failure path
    /// degrades to "no curtain", never to a broken build.
    /// </summary>
    internal sealed class ExcelBuildOverlay : IDisposable
    {
        private Thread _thread;
        private OverlayForm _form;

        public ExcelBuildOverlay(string message)
        {
            try
            {
                Rectangle bounds = Screen.FromHandle(ExcelDnaUtil.WindowHandle).WorkingArea;
                using (var shown = new ManualResetEventSlim(false))
                {
                    _thread = new Thread(() =>
                    {
                        try
                        {
                            _form = new OverlayForm(bounds, message);
                            _form.Shown += (s, e) => shown.Set();
                            Application.Run(_form);
                        }
                        catch { shown.Set(); }
                    })
                    {
                        IsBackground = true,
                        Name = "PT-IoBridge-Overlay",
                    };
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();

                    // Wait briefly so the curtain is UP before the flashing work begins.
                    shown.Wait(700);
                }
            }
            catch
            {
                _form = null;
                _thread = null;
            }
        }

        public void Dispose()
        {
            try
            {
                OverlayForm form = _form;
                _form = null;
                if (form != null && !form.IsDisposed)
                    form.BeginInvoke(new Action(() => { try { form.Close(); } catch { } }));
                _thread?.Join(1000);
            }
            catch { }
        }

        private sealed class OverlayForm : Form
        {
            private const int WS_EX_NOACTIVATE = 0x08000000;
            private const int WS_EX_TOPMOST = 0x00000008;

            public OverlayForm(Rectangle bounds, string message)
            {
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;
                Bounds = bounds;
                TopMost = true;
                ShowInTaskbar = false;
                BackColor = Color.FromArgb(32, 34, 36);
                Cursor = Cursors.WaitCursor;

                var label = new Label
                {
                    Text = message,
                    ForeColor = Color.White,
                    Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 16f),
                    AutoSize = true,
                };
                var marquee = new ProgressBar
                {
                    Style = ProgressBarStyle.Marquee,
                    MarqueeAnimationSpeed = 25,
                    Width = 320,
                    Height = 16,
                };
                Controls.Add(label);
                Controls.Add(marquee);

                Load += (s, e) =>
                {
                    label.Location = new Point(
                        (ClientSize.Width - label.Width) / 2, ClientSize.Height / 2 - label.Height - 12);
                    marquee.Location = new Point(
                        (ClientSize.Width - marquee.Width) / 2, ClientSize.Height / 2 + 12);
                };
            }

            protected override bool ShowWithoutActivation => true;

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams parameters = base.CreateParams;
                    parameters.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOPMOST;
                    return parameters;
                }
            }
        }
    }
}
