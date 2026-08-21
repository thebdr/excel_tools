using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// The non-modal SWP Bridge dashboard: phase + progress, live tallies, the per-sheet
    /// breakdown (click → filter the spine), the findings tail (click → jump to the row), and
    /// the Refresh / Apply / Cancel buttons.
    ///
    /// The form lives on its OWN STA thread (the ExcelBuildOverlay pattern): during a build or
    /// apply the Excel thread is blocked inside the macro, and only a separately-pumped window
    /// keeps painting. The rules that make this safe: Excel-side code talks to the form ONLY
    /// through BeginInvoke (never blocking the build on the UI), and the form NEVER touches COM —
    /// every gesture dispatches through ExcelAsyncUtil.QueueAsMacro onto the Excel thread, except
    /// Cancel, which just sets the session's volatile flag the builder polls between phases.
    /// </summary>
    internal sealed class SwpBridgeDashboard
    {
        private Thread _thread;
        private DashboardForm _form;

        public static SwpBridgeDashboard Show()
        {
            var dashboard = new SwpBridgeDashboard();
            try
            {
                using (var shown = new ManualResetEventSlim(false))
                {
                    dashboard._thread = new Thread(() =>
                    {
                        try
                        {
                            dashboard._form = new DashboardForm();
                            // Always-on-top: the build curtain is a topmost window, and the
                            // dashboard must narrate ABOVE it, not vanish beneath.
                            dashboard._form.TopMost = true;
                            dashboard._form.Shown += (s, e) => shown.Set();
                            Application.Run(dashboard._form);
                        }
                        catch { shown.Set(); }
                    })
                    {
                        IsBackground = true,
                        Name = "PT-SwpBridge-Dashboard",
                    };
                    dashboard._thread.SetApartmentState(ApartmentState.STA);
                    dashboard._thread.Start();
                    shown.Wait(700);
                }
            }
            catch
            {
                dashboard._form = null;
                dashboard._thread = null;
            }
            return dashboard;
        }

        public void PostPhase(string phase) => Post(form => form.SetPhase(phase));

        /// <summary>Pushes the current tallies, sheet table and findings tail (immutable copies).</summary>
        public void PostModel(SwpBridgeSession session)
        {
            if (session?.Model == null) return;
            var tallies = new List<SwpSheetTally>(session.Model.Tallies);
            int pending = 0, included = 0;
            foreach (SwpAction a in session.Queue.Actions)
                if (a.Status == SwpActionStatus.Pending)
                {
                    pending++;
                    if (a.Include) included++;
                }

            var tail = new List<(string Severity, string Text, string Mnemonic)>();
            List<SwpLogEntry> log = session.BuildLog;
            for (int i = Math.Max(0, log.Count - 60); i < log.Count; i++)
                tail.Add((log[i].Severity.ToString(), log[i].Message, log[i].Mnemonic));

            int rows = session.Model.Rows.Count;
            Post(form => form.SetModel(rows, pending, included, tallies, tail));
        }

        public void PostBusy(bool busy) => Post(form => form.SetBusy(busy));

        public void BringToFront() => Post(form =>
        {
            form.TopMost = true;   // it lives always-on-top; just make sure and activate
            form.Activate();
        });

        public void CloseFromExcel()
        {
            try
            {
                DashboardForm form = _form;
                _form = null;
                if (form != null && !form.IsDisposed)
                    form.BeginInvoke(new Action(() => { try { form.AllowClose(); form.Close(); } catch { } }));
                _thread?.Join(1000);
            }
            catch { }
        }

        private void Post(Action<DashboardForm> action)
        {
            try
            {
                DashboardForm form = _form;
                if (form != null && !form.IsDisposed && form.IsHandleCreated)
                    form.BeginInvoke(new Action(() => { try { action(form); } catch { } }));
            }
            catch { }
        }

        // ---- the form (dashboard thread only) -----------------------------------------------

        private sealed class DashboardForm : Form
        {
            private readonly Label _phase = new Label { AutoSize = false, Dock = DockStyle.Fill, Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 10.5f, FontStyle.Bold) };
            private readonly ProgressBar _progress = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 0, Height = 14 };
            private readonly Label _tallies = new Label { AutoSize = false, Dock = DockStyle.Fill };
            private readonly ListView _sheets = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, MultiSelect = false };
            private readonly ListView _log = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, MultiSelect = false };
            private bool _closable;

            public DashboardForm()
            {
                Text = "SWP Bridge";
                FormBorderStyle = FormBorderStyle.Sizable;
                StartPosition = FormStartPosition.Manual;
                ShowInTaskbar = true;
                MinimizeBox = true;
                MaximizeBox = false;
                ClientSize = new Size(520, 560);
                MinimumSize = new Size(420, 420);
                Font = SystemFonts.MessageBoxFont;
                try
                {
                    Rectangle area = Screen.FromHandle(ExcelDnaUtil.WindowHandle).WorkingArea;
                    Location = new Point(area.Right - Width - 24, area.Top + 80);
                }
                catch { }

                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10), ColumnCount = 1, RowCount = 6 };
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));   // phase
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));   // progress
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));   // tallies
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 42));    // sheets
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 58));    // log
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));   // buttons

                _phase.Text = "Ready";
                layout.Controls.Add(_phase, 0, 0);
                layout.Controls.Add(_progress, 0, 1);
                layout.Controls.Add(_tallies, 0, 2);

                _sheets.Columns.Add("Sheet", 150);
                _sheets.Columns.Add("Total", 50, HorizontalAlignment.Right);
                _sheets.Columns.Add("Pending", 60, HorizontalAlignment.Right);
                _sheets.Columns.Add("Mismatch", 64, HorizontalAlignment.Right);
                _sheets.Columns.Add("Missing", 60, HorizontalAlignment.Right);
                _sheets.Columns.Add("OK", 50, HorizontalAlignment.Right);
                _sheets.ItemActivate += (s, e) => FilterSpineToSelectedSheet();
                layout.Controls.Add(_sheets, 0, 3);

                _log.Columns.Add("Sev", 46);
                _log.Columns.Add("Finding", 400);
                _log.ItemActivate += (s, e) => JumpToSelectedFinding();
                layout.Controls.Add(_log, 0, 4);

                var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = new Padding(0) };
                var apply = new Button { Text = "Apply queue", Width = 100, Height = 30 };
                var refresh = new Button { Text = "Refresh", Width = 84, Height = 30 };
                var cancel = new Button { Text = "Cancel", Width = 84, Height = 30 };
                apply.Click += (s, e) => RunOnExcel(() => SwpBridgeCommands.ApplyBatch());
                refresh.Click += (s, e) => RunOnExcel(() => SwpBridgeCommands.RefreshWorkspace());
                cancel.Click += (s, e) => { var session = SwpBridgeSession.Current; if (session != null) session.CancelRequested = true; };
                buttons.Controls.Add(apply);
                buttons.Controls.Add(refresh);
                buttons.Controls.Add(cancel);
                layout.Controls.Add(buttons, 0, 5);

                Controls.Add(layout);
                ExcelPowerTools.UI.GdiText.Apply(this);   // GDI text, like the native controls
            }

            public void AllowClose() => _closable = true;

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                // The X hides the window; the session teardown closes it for real.
                if (!_closable && e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                    return;
                }
                base.OnFormClosing(e);
            }

            public void SetPhase(string phase)
            {
                _phase.Text = phase ?? "";
                bool busy = phase != null && phase != "Ready";
                SetBusy(busy);
                if (busy && !Visible) Show();
            }

            public void SetBusy(bool busy) => _progress.MarqueeAnimationSpeed = busy ? 25 : 0;

            public void SetModel(int rows, int pending, int included,
                List<SwpSheetTally> tallies, List<(string Severity, string Text, string Mnemonic)> tail)
            {
                _tallies.Text = rows + " rows · " + pending + " pending (" + included + " included)";

                _sheets.BeginUpdate();
                _sheets.Items.Clear();
                foreach (SwpSheetTally tally in tallies)
                {
                    var item = new ListViewItem(tally.SheetName);
                    item.SubItems.Add(tally.Total.ToString());
                    item.SubItems.Add((tally.Pending + tally.LabelMismatch).ToString());
                    item.SubItems.Add(tally.LabelMismatch.ToString());
                    item.SubItems.Add((tally.RowMissing + tally.SheetMissing + tally.Invalid + tally.Duplicate).ToString());
                    item.SubItems.Add(tally.UpToDate.ToString());
                    _sheets.Items.Add(item);
                }
                _sheets.EndUpdate();

                _log.BeginUpdate();
                _log.Items.Clear();
                foreach (var line in tail)
                {
                    var item = new ListViewItem(line.Severity == "Info" ? "" : line.Severity);
                    item.SubItems.Add(line.Text);
                    item.Tag = line.Mnemonic;
                    if (line.Severity == "Error") item.ForeColor = Color.FromArgb(156, 0, 6);
                    else if (line.Severity == "Warning") item.ForeColor = Color.FromArgb(140, 100, 0);
                    _log.Items.Add(item);
                }
                if (_log.Items.Count > 0) _log.EnsureVisible(_log.Items.Count - 1);
                _log.EndUpdate();
            }

            private void FilterSpineToSelectedSheet()
            {
                if (_sheets.SelectedItems.Count == 0) return;
                string sheetName = _sheets.SelectedItems[0].Text;
                RunOnExcel(() => SwpBridgeCommands.FilterSpineToSheet(sheetName));
            }

            private void JumpToSelectedFinding()
            {
                if (_log.SelectedItems.Count == 0) return;
                string mnemonic = _log.SelectedItems[0].Tag as string;
                if (string.IsNullOrEmpty(mnemonic)) return;
                RunOnExcel(() => SwpBridgeCommands.JumpToMnemonic(mnemonic));
            }

            /// <summary>All COM work runs on the Excel thread, in macro context.</summary>
            private static void RunOnExcel(Action action)
            {
                try { ExcelAsyncUtil.QueueAsMacro(() => { try { action(); } catch { } }); }
                catch { }
            }
        }
    }
}
