using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPowerTools.UI
{
    public sealed class ComparisonBuildProgressUpdate
    {
        public ComparisonBuildProgressUpdate(
            string phase,
            string detail,
            int? percent = null,
            bool canCancel = true)
        {
            Phase = phase ?? "";
            Detail = detail ?? "";
            Percent = percent;
            CanCancel = canCancel;
        }

        public string Phase { get; }
        public string Detail { get; }

        /// <summary>Zero through one hundred, or null for indeterminate progress.</summary>
        public int? Percent { get; }

        /// <summary>
        /// False while an operation is in an atomic section that must finish before cancellation is safe.
        /// </summary>
        public bool CanCancel { get; }
    }

    public sealed class ComparisonBuildProgressDialogOptions
    {
        public string Title { get; set; } = "Build comparison project";
        public string Heading { get; set; } = "Building comparison project";
        public string InitialPhase { get; set; } = "Preparing";
        public string InitialDetail { get; set; } = "Inspecting the selected revisions.";
        public bool InitialCanCancel { get; set; } = true;
        public string NonCancelableClosingPhase { get; set; } = "Finishing atomic publication...";
        public string NonCancelableClosingDetail { get; set; } =
            "This window can close after the new review workbook is safely published.";

        /// <summary>
        /// A coarse, non-sensitive workload key, preferably created with ComparisonWorkloadBucket.Create.
        /// It must never contain a source path, filename, workbook label, or cell value.
        /// </summary>
        public string WorkloadBucket { get; set; }
    }

    /// <summary>
    /// Modal progress host for pure comparison builds. It keeps the Excel UI responsive, waits for safe
    /// cancellation acknowledgement, records only successful elapsed times, and rethrows worker failures
    /// to its caller after the dialog closes.
    /// </summary>
    public sealed class ComparisonBuildProgressDialog : Form
    {
        private static readonly Color Ink = Color.FromArgb(32, 39, 57);
        private static readonly Color Muted = Color.FromArgb(96, 106, 127);
        private static readonly Color Surface = Color.FromArgb(248, 249, 252);
        private static readonly Color Accent = Color.FromArgb(92, 75, 192);

        private readonly ComparisonBuildProgressDialogOptions _options;
        private readonly Func<IProgress<ComparisonBuildProgressUpdate>, CancellationToken, Task<object>> _operation;
        private readonly ComparisonElapsedTimeEstimateStore _estimateStore;
        private readonly CancellationTokenSource _cancellation = new CancellationTokenSource();
        private readonly Stopwatch _stopwatch = new Stopwatch();
        private readonly System.Windows.Forms.Timer _clock = new System.Windows.Forms.Timer();
        private readonly ToolTip _toolTips = new ToolTip
        {
            InitialDelay = 400,
            ReshowDelay = 100,
            AutoPopDelay = 15000,
            ShowAlways = true,
        };

        private readonly Label _phase = new Label();
        private readonly Label _detail = new Label();
        private readonly Label _timing = new Label();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly Button _cancel = new Button();

        private TimeSpan? _estimatedDuration;
        private int _estimateSamples;
        private ExceptionDispatchInfo _failure;
        private object _result;
        private bool _operationFinished;
        private bool _canCancel = true;
        private bool _cancelRequested;

        private ComparisonBuildProgressDialog(
            ComparisonBuildProgressDialogOptions options,
            Func<IProgress<ComparisonBuildProgressUpdate>, CancellationToken, Task<object>> operation,
            ComparisonElapsedTimeEstimateStore estimateStore)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _operation = operation ?? throw new ArgumentNullException(nameof(operation));
            _estimateStore = estimateStore ?? ComparisonElapsedTimeEstimateStore.Default;

            Text = string.IsNullOrWhiteSpace(options.Title) ? "Build comparison project" : options.Title;
            Name = "ComparisonBuildProgressDialog";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(620, 286);
            MinimumSize = new Size(520, 286);
            MaximumSize = new Size(1000, 286);
            BackColor = Surface;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = true;

            BuildLayout();
            _canCancel = options.InitialCanCancel;
            _cancel.Enabled = _canCancel;
            _phase.Text = options.InitialPhase ?? "";
            _detail.Text = options.InitialDetail ?? "";
            _clock.Interval = 500;
            _clock.Tick += (_, __) => UpdateTimingText();
            _cancel.Click += (_, __) => RequestCancellation();
            Shown += async (_, __) => await StartOperationAsync();
            FormClosing += OnFormClosing;
            CancelButton = _cancel;
        }

        /// <summary>
        /// Runs an asynchronous, non-COM operation on a worker thread while this modal dialog owns the UI.
        /// The return value is forwarded on success. Cancellation and all other exceptions are rethrown to
        /// the caller, preserving the original exception and stack trace.
        /// </summary>
        public static T Run<T>(
            IWin32Window owner,
            ComparisonBuildProgressDialogOptions options,
            Func<IProgress<ComparisonBuildProgressUpdate>, CancellationToken, Task<T>> operation,
            ComparisonElapsedTimeEstimateStore estimateStore = null)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));

            async Task<object> BoxedOperation(
                IProgress<ComparisonBuildProgressUpdate> progress,
                CancellationToken cancellationToken)
            {
                return await operation(progress, cancellationToken).ConfigureAwait(false);
            }

            using (var dialog = new ComparisonBuildProgressDialog(options, BoxedOperation, estimateStore))
            {
                dialog.ShowDialog(owner);
                return dialog.GetResult<T>();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _clock.Stop();
                _clock.Dispose();
                _cancellation.Dispose();
                _toolTips.Dispose();
            }
            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(20, 16, 20, 14),
                Margin = Padding.Empty,
                BackColor = Surface,
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 31));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));

            var heading = new Label
            {
                Dock = DockStyle.Fill,
                Text = string.IsNullOrWhiteSpace(_options.Heading) ? "Building comparison project" : _options.Heading,
                Font = new Font("Segoe UI Semibold", 15f, FontStyle.Bold),
                ForeColor = Ink,
                UseCompatibleTextRendering = false,
            };
            _phase.Dock = DockStyle.Fill;
            _phase.Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold);
            _phase.ForeColor = Accent;
            _phase.UseCompatibleTextRendering = false;
            _detail.Dock = DockStyle.Fill;
            _detail.ForeColor = Ink;
            _detail.AutoEllipsis = true;
            _detail.UseCompatibleTextRendering = false;
            _progress.Dock = DockStyle.Top;
            _progress.Height = 19;
            _progress.Minimum = 0;
            _progress.Maximum = 100;
            _progress.Style = ProgressBarStyle.Marquee;
            _progress.MarqueeAnimationSpeed = 24;
            _timing.Dock = DockStyle.Fill;
            _timing.ForeColor = Muted;
            _timing.UseCompatibleTextRendering = false;
            _cancel.Text = "Cancel";
            _cancel.Size = new Size(96, 29);
            _cancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            var buttonRow = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            _cancel.Location = new Point(buttonRow.Width - _cancel.Width, 5);
            buttonRow.Resize += (_, __) =>
                _cancel.Location = new Point(Math.Max(0, buttonRow.ClientSize.Width - _cancel.Width), 5);
            buttonRow.Controls.Add(_cancel);

            SetHelp(_phase, "Current comparison phase.");
            SetHelp(_detail, "Detailed progress for the current safe comparison operation.");
            SetHelp(_progress, "Progress is determinate when the current phase has a measurable total; otherwise it animates while work continues.");
            SetHelp(_timing, "Elapsed time and, when prior comparable runs exist, a locally learned non-sensitive duration estimate.");
            SetHelp(_cancel, "Request safe cancellation. This is temporarily unavailable while an atomic operation must finish.");
            SetHelp(buttonRow, "Request safe cancellation. This is temporarily unavailable while an atomic operation must finish.");

            root.Controls.Add(heading, 0, 0);
            root.Controls.Add(_phase, 0, 1);
            root.Controls.Add(_detail, 0, 2);
            root.Controls.Add(_progress, 0, 3);
            root.Controls.Add(_timing, 0, 4);
            root.Controls.Add(buttonRow, 0, 5);
            Controls.Add(root);
        }

        private void SetHelp(Control control, string helpText)
        {
            _toolTips.SetToolTip(control, helpText);
            control.AccessibleDescription = helpText;
        }

        private async Task StartOperationAsync()
        {
            if (!string.IsNullOrWhiteSpace(_options.WorkloadBucket) &&
                _estimateStore.TryGetEstimate(_options.WorkloadBucket, out TimeSpan estimate, out int samples))
            {
                _estimatedDuration = estimate;
                _estimateSamples = samples;
            }

            UpdateTimingText();
            _stopwatch.Start();
            _clock.Start();
            var progress = new Progress<ComparisonBuildProgressUpdate>(ApplyProgress);

            try
            {
                // Comparison parsing and packaging are non-COM work. Leaving Excel's UI thread here also
                // protects responsiveness when an async implementation does substantial work before yielding.
                _result = await Task.Run(
                    () => _operation(progress, _cancellation.Token),
                    _cancellation.Token);

                _stopwatch.Stop();
                if (!string.IsNullOrWhiteSpace(_options.WorkloadBucket))
                    _estimateStore.TryRecordSuccessfulRun(_options.WorkloadBucket, _stopwatch.Elapsed);
                CompleteAndClose(DialogResult.OK);
            }
            catch (Exception exception)
            {
                _stopwatch.Stop();
                _failure = ExceptionDispatchInfo.Capture(exception);
                CompleteAndClose(exception is OperationCanceledException
                    ? DialogResult.Cancel
                    : DialogResult.Abort);
            }
        }

        private void ApplyProgress(ComparisonBuildProgressUpdate update)
        {
            if (update == null || IsDisposed || _operationFinished) return;
            _phase.Text = update.Phase;
            _detail.Text = update.Detail;
            _canCancel = update.CanCancel;
            _cancel.Enabled = update.CanCancel && !_cancelRequested;

            if (update.Percent.HasValue)
            {
                if (_progress.Style != ProgressBarStyle.Continuous)
                {
                    _progress.MarqueeAnimationSpeed = 0;
                    _progress.Style = ProgressBarStyle.Continuous;
                }
                _progress.Value = Math.Max(_progress.Minimum, Math.Min(_progress.Maximum, update.Percent.Value));
            }
            else if (_progress.Style != ProgressBarStyle.Marquee)
            {
                _progress.Style = ProgressBarStyle.Marquee;
                _progress.MarqueeAnimationSpeed = 24;
            }
        }

        private void RequestCancellation()
        {
            if (_operationFinished || _cancelRequested || !_canCancel) return;
            _cancelRequested = true;
            _cancel.Enabled = false;
            _phase.Text = "Cancelling...";
            _detail.Text = "Waiting for the current operation to stop safely.";
            _cancellation.Cancel();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (_operationFinished) return;
            e.Cancel = true;
            if (_canCancel)
            {
                RequestCancellation();
            }
            else
            {
                _phase.Text = _options.NonCancelableClosingPhase ?? "Finishing the protected operation...";
                _detail.Text = _options.NonCancelableClosingDetail ??
                               "This window can close after the current operation finishes safely.";
            }
        }

        private void UpdateTimingText()
        {
            TimeSpan elapsed = _stopwatch.Elapsed;
            if (!_estimatedDuration.HasValue)
            {
                _timing.Text = "Elapsed " + FormatDuration(elapsed) +
                               "  ·  An estimate will be learned from successful runs of similar workloads.";
                return;
            }

            TimeSpan remaining = _estimatedDuration.Value - elapsed;
            if (remaining < TimeSpan.Zero) remaining = TimeSpan.Zero;
            string sampleWord = _estimateSamples == 1 ? "run" : "runs";
            _timing.Text = "Elapsed " + FormatDuration(elapsed) +
                           "  ·  Estimated about " + FormatDuration(_estimatedDuration.Value) +
                           " from " + _estimateSamples + " prior " + sampleWord +
                           "  ·  About " + FormatDuration(remaining) + " remaining";
        }

        private void CompleteAndClose(DialogResult result)
        {
            _clock.Stop();
            _operationFinished = true;
            DialogResult = result;
            Close();
        }

        private T GetResult<T>()
        {
            if (_failure != null) _failure.Throw();
            if (!_operationFinished)
                throw new OperationCanceledException("The comparison progress dialog closed before the operation completed.");
            return (T)_result;
        }

        private static string FormatDuration(TimeSpan value)
        {
            if (value < TimeSpan.Zero) value = TimeSpan.Zero;
            if (value.TotalHours >= 1)
                return ((int)value.TotalHours).ToString("0") + ":" + value.Minutes.ToString("00") + ":" +
                       value.Seconds.ToString("00");
            return ((int)value.TotalMinutes).ToString("0") + ":" + value.Seconds.ToString("00");
        }
    }
}
