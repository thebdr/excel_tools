using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Comparison.Revision;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Explicit Before/After setup and progress UI for the bounded pairwise preview. The labels and limitation
    /// banner are intentionally prominent: file-picker order is never treated as revision evidence.
    /// </summary>
    public sealed class ComparisonSetupForm : Form
    {
        private static readonly Color Ink = Color.FromArgb(32, 39, 57);
        private static readonly Color Muted = Color.FromArgb(96, 106, 127);
        private static readonly Color Border = Color.FromArgb(215, 221, 232);
        private static readonly Color Accent = Color.FromArgb(92, 75, 192);
        private static readonly Color BannerBackground = Color.FromArgb(249, 245, 255);
        private static readonly Color Surface = Color.FromArgb(248, 249, 252);

        private readonly TextBox _beforePath = new TextBox();
        private readonly TextBox _afterPath = new TextBox();
        private readonly TextBox _outputPath = new TextBox();
        private readonly Label _beforeMetadata = new Label();
        private readonly Label _afterMetadata = new Label();
        private readonly Button _browseBefore = new Button();
        private readonly Button _browseAfter = new Button();
        private readonly Button _browseOutput = new Button();
        private readonly Button _swap = new Button();
        private readonly Button _run = new Button();
        private readonly Button _cancel = new Button();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly Label _phase = new Label();
        private readonly Label _detail = new Label();

        private CancellationTokenSource _cancellation;
        private bool _running;
        private bool _canCancel = true;
        private bool _settingOutput;
        private bool _outputIsAutomatic = true;

        public ComparisonSetupForm(string beforePath = null, string afterPath = null)
        {
            Text = "Exact Compare Preview";
            Name = "ComparisonSetupForm";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(790, 590);
            MinimumSize = new Size(680, 550);
            BackColor = Surface;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            KeyPreview = true;

            BuildLayout();
            WireEvents();

            _beforePath.Text = beforePath ?? "";
            _afterPath.Text = afterPath ?? "";
            UpdateArtifactMetadata(_beforePath.Text, _beforeMetadata, "Before");
            UpdateArtifactMetadata(_afterPath.Text, _afterMetadata, "After");
            UpdateSuggestedOutput(force: true);

            AcceptButton = _run;
            CancelButton = _cancel;
            FormClosing += OnFormClosing;
        }

        public ExactPreviewResult Result { get; private set; }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_cancellation != null)
                {
                    _cancellation.Dispose();
                    _cancellation = null;
                }
            }
            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Surface,
                ColumnCount = 1,
                RowCount = 8,
                Margin = Padding.Empty,
                Padding = new Padding(18, 14, 18, 14),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 91));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 39));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 91));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 86));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildLimitationBanner(), 0, 1);
            root.Controls.Add(BuildArtifactPanel(
                "BEFORE  ·  EARLIER REVISION",
                "The baseline side of every visible change",
                _beforePath,
                _beforeMetadata,
                _browseBefore), 0, 2);
            root.Controls.Add(BuildSwapRow(), 0, 3);
            root.Controls.Add(BuildArtifactPanel(
                "AFTER  ·  LATER / MASTER CANDIDATE",
                "The candidate side of every visible change",
                _afterPath,
                _afterMetadata,
                _browseAfter), 0, 4);
            root.Controls.Add(BuildOutputPanel(), 0, 5);
            root.Controls.Add(BuildProgressPanel(), 0, 6);
            root.Controls.Add(BuildButtonRow(), 0, 7);

            Controls.Add(root);
        }

        private Control BuildHeader()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                BackColor = Surface,
            };
            var title = new Label
            {
                AutoSize = true,
                Location = new Point(0, 1),
                Text = "Exact Compare Preview",
                ForeColor = Ink,
                Font = new Font("Segoe UI Semibold", 16f, FontStyle.Bold),
                UseCompatibleTextRendering = false,
            };
            var subtitle = new Label
            {
                AutoSize = true,
                Location = new Point(2, 35),
                Text = "Generate a new, self-contained Review.xlsx without modifying either source.",
                ForeColor = Muted,
                Font = new Font("Segoe UI", 9.2f),
                UseCompatibleTextRendering = false,
            };
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            return panel;
        }

        private Control BuildLimitationBanner()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 2, 0, 7),
                Padding = new Padding(13, 9, 12, 8),
                BackColor = BannerBackground,
            };
            panel.Paint += (_, args) =>
            {
                using (var pen = new Pen(Color.FromArgb(203, 190, 237)))
                    args.Graphics.DrawRectangle(pen, 0, 0, panel.ClientSize.Width - 1, panel.ClientSize.Height - 1);
                using (var brush = new SolidBrush(Accent))
                    args.Graphics.FillRectangle(brush, 0, 0, 4, panel.ClientSize.Height);
            };

            var label = new Label
            {
                Dock = DockStyle.Fill,
                Text = "PREVIEW LIMITS\r\nExact positional comparison of two .xlsx/.xlsm workbooks or two separated-value files. " +
                       "Sheets pair only by exact name; inserted rows can cascade. No rename/row alignment, grouping, snapshots, decisions, merge, or XLSB yet.",
                ForeColor = Color.FromArgb(64, 48, 111),
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8.7f),
                UseCompatibleTextRendering = false,
            };
            panel.Controls.Add(label);
            return panel;
        }

        private Control BuildArtifactPanel(
            string heading,
            string description,
            TextBox path,
            Label metadata,
            Button browse)
        {
            var shell = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 3, 0, 3),
                Padding = new Padding(12, 8, 10, 7),
                BackColor = Color.White,
            };
            shell.Paint += (_, args) =>
            {
                using (var pen = new Pen(Border))
                    args.Graphics.DrawRectangle(pen, 0, 0, shell.ClientSize.Width - 1, shell.ClientSize.Height - 1);
            };

            var headingLabel = new Label
            {
                AutoSize = true,
                Location = new Point(12, 7),
                Text = heading,
                ForeColor = Accent,
                Font = new Font("Segoe UI Semibold", 8.2f, FontStyle.Bold),
                UseCompatibleTextRendering = false,
            };
            var descriptionLabel = new Label
            {
                AutoSize = true,
                Location = new Point(234, 7),
                Text = description,
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8.1f),
                UseCompatibleTextRendering = false,
            };

            path.Location = new Point(12, 31);
            path.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            path.Size = new Size(635, 24);
            path.Font = new Font("Segoe UI", 9f);
            path.BorderStyle = BorderStyle.FixedSingle;

            browse.Text = "Browse...";
            browse.Location = new Point(655, 29);
            browse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            browse.Size = new Size(95, 27);
            browse.UseVisualStyleBackColor = true;

            metadata.Location = new Point(13, 60);
            metadata.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            metadata.Size = new Size(735, 18);
            metadata.ForeColor = Muted;
            metadata.Font = new Font("Segoe UI", 8.1f);
            metadata.AutoEllipsis = true;
            metadata.UseCompatibleTextRendering = false;

            shell.Resize += (_, __) =>
            {
                browse.Left = Math.Max(120, shell.ClientSize.Width - browse.Width - 10);
                path.Width = Math.Max(100, browse.Left - path.Left - 8);
                metadata.Width = Math.Max(100, shell.ClientSize.Width - metadata.Left - 10);
            };
            shell.Controls.Add(headingLabel);
            shell.Controls.Add(descriptionLabel);
            shell.Controls.Add(path);
            shell.Controls.Add(browse);
            shell.Controls.Add(metadata);
            return shell;
        }

        private Control BuildSwapRow()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                BackColor = Surface,
            };
            _swap.Text = "⇅  Swap Before / After";
            _swap.AutoSize = false;
            _swap.Size = new Size(178, 29);
            _swap.Location = new Point(0, 5);
            _swap.Anchor = AnchorStyles.Top;
            _swap.FlatStyle = FlatStyle.Flat;
            _swap.FlatAppearance.BorderColor = Color.FromArgb(184, 177, 219);
            _swap.BackColor = Color.White;
            _swap.ForeColor = Color.FromArgb(72, 57, 137);
            _swap.Font = new Font("Segoe UI Semibold", 8.7f, FontStyle.Bold);
            panel.Resize += (_, __) => _swap.Left = Math.Max(0, (panel.ClientSize.Width - _swap.Width) / 2);
            panel.Controls.Add(_swap);
            return panel;
        }

        private Control BuildOutputPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 7, 0, 3),
                BackColor = Surface,
            };
            var heading = new Label
            {
                AutoSize = true,
                Location = new Point(1, 1),
                Text = "NEW REVIEW WORKBOOK",
                ForeColor = Ink,
                Font = new Font("Segoe UI Semibold", 8.3f, FontStyle.Bold),
                UseCompatibleTextRendering = false,
            };
            var note = new Label
            {
                AutoSize = true,
                Location = new Point(172, 1),
                Text = "An existing file is never overwritten",
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8.1f),
                UseCompatibleTextRendering = false,
            };

            _outputPath.Location = new Point(1, 27);
            _outputPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _outputPath.Size = new Size(646, 24);
            _outputPath.Font = new Font("Segoe UI", 9f);
            _outputPath.BorderStyle = BorderStyle.FixedSingle;

            _browseOutput.Text = "Choose...";
            _browseOutput.Location = new Point(655, 25);
            _browseOutput.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _browseOutput.Size = new Size(95, 27);
            _browseOutput.UseVisualStyleBackColor = true;

            panel.Resize += (_, __) =>
            {
                _browseOutput.Left = Math.Max(120, panel.ClientSize.Width - _browseOutput.Width);
                _outputPath.Width = Math.Max(100, _browseOutput.Left - _outputPath.Left - 8);
            };
            panel.Controls.Add(heading);
            panel.Controls.Add(note);
            panel.Controls.Add(_outputPath);
            panel.Controls.Add(_browseOutput);
            return panel;
        }

        private Control BuildProgressPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 7, 0, 4),
                BackColor = Surface,
            };
            _phase.Location = new Point(1, 2);
            _phase.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _phase.Size = new Size(748, 20);
            _phase.Text = "Ready";
            _phase.ForeColor = Ink;
            _phase.Font = new Font("Segoe UI Semibold", 8.8f, FontStyle.Bold);
            _phase.AutoEllipsis = true;
            _phase.UseCompatibleTextRendering = false;

            _progress.Location = new Point(1, 27);
            _progress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _progress.Size = new Size(748, 17);
            _progress.Style = ProgressBarStyle.Continuous;

            _detail.Location = new Point(1, 49);
            _detail.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _detail.Size = new Size(748, 20);
            _detail.Text = "Confirm the order, choose a new output name, then run the preview.";
            _detail.ForeColor = Muted;
            _detail.Font = new Font("Segoe UI", 8.1f);
            _detail.AutoEllipsis = true;
            _detail.UseCompatibleTextRendering = false;

            panel.Resize += (_, __) =>
            {
                _phase.Width = Math.Max(100, panel.ClientSize.Width - 2);
                _progress.Width = Math.Max(100, panel.ClientSize.Width - 2);
                _detail.Width = Math.Max(100, panel.ClientSize.Width - 2);
            };
            panel.Controls.Add(_phase);
            panel.Controls.Add(_progress);
            panel.Controls.Add(_detail);
            return panel;
        }

        private Control BuildButtonRow()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                BackColor = Surface,
            };
            _cancel.Text = "Close";
            _cancel.Size = new Size(92, 31);
            _cancel.Location = new Point(0, 7);
            _cancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _cancel.UseVisualStyleBackColor = true;

            _run.Text = "Generate review";
            _run.Size = new Size(142, 31);
            _run.Location = new Point(0, 7);
            _run.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _run.BackColor = Accent;
            _run.ForeColor = Color.White;
            _run.FlatStyle = FlatStyle.Flat;
            _run.FlatAppearance.BorderSize = 0;
            _run.Font = new Font("Segoe UI Semibold", 9f, FontStyle.Bold);

            panel.Resize += (_, __) =>
            {
                _run.Left = Math.Max(0, panel.ClientSize.Width - _run.Width);
                _cancel.Left = Math.Max(0, _run.Left - _cancel.Width - 9);
            };
            panel.Controls.Add(_cancel);
            panel.Controls.Add(_run);
            return panel;
        }

        private void WireEvents()
        {
            _beforePath.TextChanged += (_, __) =>
            {
                UpdateArtifactMetadata(_beforePath.Text, _beforeMetadata, "Before");
                UpdateSuggestedOutput(force: false);
            };
            _afterPath.TextChanged += (_, __) =>
            {
                UpdateArtifactMetadata(_afterPath.Text, _afterMetadata, "After");
                UpdateSuggestedOutput(force: false);
            };
            _outputPath.TextChanged += (_, __) =>
            {
                if (!_settingOutput) _outputIsAutomatic = false;
            };
            _browseBefore.Click += (_, __) => BrowseInput(_beforePath);
            _browseAfter.Click += (_, __) => BrowseInput(_afterPath);
            _browseOutput.Click += (_, __) => BrowseOutput();
            _swap.Click += (_, __) => SwapInputs();
            _run.Click += RunClicked;
            _cancel.Click += CancelClicked;
        }

        private void BrowseInput(TextBox target)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = ComparisonCommands.SupportedInputFilter,
                CheckFileExists = true,
                Multiselect = false,
                FileName = SafeFileName(target.Text),
                InitialDirectory = SafeDirectory(target.Text),
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK) target.Text = dialog.FileName;
            }
        }

        private void BrowseOutput()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = ComparisonCommands.ReviewWorkbookFilter,
                AddExtension = true,
                DefaultExt = "xlsx",
                OverwritePrompt = false,
                CheckPathExists = true,
                FileName = SafeFileName(_outputPath.Text),
                InitialDirectory = SafeDirectory(_outputPath.Text),
                Title = "Choose a new Review.xlsx filename",
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _settingOutput = true;
                try { _outputPath.Text = dialog.FileName; }
                finally { _settingOutput = false; }
                _outputIsAutomatic = false;
            }
        }

        private void SwapInputs()
        {
            string before = _beforePath.Text;
            _beforePath.Text = _afterPath.Text;
            _afterPath.Text = before;
            if (_outputIsAutomatic) UpdateSuggestedOutput(force: true);
        }

        private async void RunClicked(object sender, EventArgs e)
        {
            if (_running) return;

            string before = (_beforePath.Text ?? "").Trim();
            string after = (_afterPath.Text ?? "").Trim();
            string output = (_outputPath.Text ?? "").Trim();
            string validation = ValidateSetup(before, after, output);
            if (validation != null)
            {
                ExcelPowerTools.Interop.Dialogs.Info(validation, "Exact Compare Preview", this);
                return;
            }

            Result = null;
            _cancellation = new CancellationTokenSource();
            CancellationToken token = _cancellation.Token;
            _canCancel = true;
            SetRunning(true);
            _progress.Value = 0;
            _phase.Text = "Starting";
            _detail.Text = "No source file will be opened or modified.";

            var progress = new Progress<ComparisonPreviewProgress>(UpdateProgress);
            try
            {
                // Task.Run makes the non-COM inventory/parser/package pipeline unconditionally leave Excel's
                // UI thread, including for tiny files whose asynchronous reads might otherwise complete inline.
                Result = await Task.Run(() => ComparisonCommands.GenerateExactPreviewAsync(
                    before,
                    after,
                    output,
                    progress,
                    token));
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (OperationCanceledException)
            {
                _phase.Text = "Cancelled";
                _detail.Text = "No review workbook was published. You can change the setup and run again.";
                _progress.Value = 0;
            }
            catch (Exception exception)
            {
                _phase.Text = "Comparison stopped";
                _detail.Text = FirstLine(exception.Message);
                ComparisonFailureDialog.Show(this, exception);
            }
            finally
            {
                if (!IsDisposed && DialogResult != DialogResult.OK) SetRunning(false);
                if (_cancellation != null)
                {
                    _cancellation.Dispose();
                    _cancellation = null;
                }
            }
        }

        private void CancelClicked(object sender, EventArgs e)
        {
            if (!_running)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }
            if (!_canCancel || _cancellation == null) return;
            _cancel.Enabled = false;
            _phase.Text = "Cancelling...";
            _detail.Text = "Waiting for the current bounded operation to stop safely.";
            _cancellation.Cancel();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (!_running || DialogResult == DialogResult.OK) return;
            e.Cancel = true;
            if (_canCancel && _cancellation != null)
            {
                _cancel.Enabled = false;
                _phase.Text = "Cancelling...";
                _detail.Text = "The window will remain available until the worker stops safely.";
                _cancellation.Cancel();
            }
            else
            {
                _phase.Text = "Finishing atomic publication...";
                _detail.Text = "The window can close after the new workbook has been written and verified.";
            }
        }

        private void SetRunning(bool running)
        {
            _running = running;
            _beforePath.Enabled = !running;
            _afterPath.Enabled = !running;
            _outputPath.Enabled = !running;
            _browseBefore.Enabled = !running;
            _browseAfter.Enabled = !running;
            _browseOutput.Enabled = !running;
            _swap.Enabled = !running;
            _run.Enabled = !running;
            _cancel.Text = running ? "Cancel" : "Close";
            _cancel.Enabled = !running || _canCancel;
        }

        private void UpdateProgress(ComparisonPreviewProgress update)
        {
            if (update == null || IsDisposed) return;
            _phase.Text = update.Phase;
            _detail.Text = update.Detail;
            _progress.Value = Math.Max(_progress.Minimum, Math.Min(_progress.Maximum, update.Percent));
            _canCancel = update.CanCancel;
            if (_running) _cancel.Enabled = update.CanCancel;
        }

        private void UpdateSuggestedOutput(bool force)
        {
            if (!force && !_outputIsAutomatic && !string.IsNullOrWhiteSpace(_outputPath.Text)) return;
            string suggestion = ComparisonCommands.SuggestOutputPath(_beforePath.Text, _afterPath.Text);
            _settingOutput = true;
            try { _outputPath.Text = suggestion; }
            finally { _settingOutput = false; }
            _outputIsAutomatic = true;
        }

        private static void UpdateArtifactMetadata(string path, Label label, string role)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                label.Text = role + " file not selected";
                return;
            }
            try
            {
                var file = new FileInfo(Path.GetFullPath(path.Trim()));
                if (!file.Exists)
                {
                    label.Text = "File not found";
                    return;
                }

                RevisionLabelInferenceResult inference = RevisionLabelInference.Infer(file.Name);
                string revision;
                if (inference.HasStrongCandidate)
                {
                    revision = "filename suggests revision " + inference.StrongCandidate.Number +
                               " (" + inference.StrongCandidate.RawToken + ")";
                }
                else if (inference.Candidates.Count > 0)
                {
                    revision = "filename revision is ambiguous - confirm the order";
                }
                else
                {
                    revision = "no reliable filename revision - confirm the order";
                }

                label.Text = file.Extension.ToUpperInvariant() + "  ·  " + FormatBytes(file.Length) +
                             "  ·  modified " + file.LastWriteTime.ToString("g", CultureInfo.CurrentCulture) +
                             "  ·  " + revision;
            }
            catch (Exception exception)
            {
                label.Text = "Path cannot be inspected: " + FirstLine(exception.Message);
            }
        }

        private static string ValidateSetup(string before, string after, string output)
        {
            if (before.Length == 0) return "Choose the Before file.";
            if (after.Length == 0) return "Choose the After file.";
            if (output.Length == 0) return "Choose a new Review.xlsx destination.";

            string beforeFull;
            string afterFull;
            string outputFull;
            try { beforeFull = Path.GetFullPath(before); }
            catch { return "The Before path is invalid."; }
            try { afterFull = Path.GetFullPath(after); }
            catch { return "The After path is invalid."; }
            try { outputFull = Path.GetFullPath(output); }
            catch { return "The Review.xlsx path is invalid."; }

            if (!File.Exists(beforeFull)) return "The Before file does not exist.";
            if (!File.Exists(afterFull)) return "The After file does not exist.";
            if (string.Equals(beforeFull, afterFull, StringComparison.OrdinalIgnoreCase))
                return "Before and After are the same file. Choose two distinct files.";
            if (!string.Equals(Path.GetExtension(outputFull), ".xlsx", StringComparison.OrdinalIgnoreCase))
                return "The review destination must end in .xlsx.";
            string outputDirectory = Path.GetDirectoryName(outputFull);
            if (string.IsNullOrEmpty(outputDirectory) || !Directory.Exists(outputDirectory))
                return "The review destination folder does not exist.";
            if (File.Exists(outputFull) || Directory.Exists(outputFull))
                return "That review destination already exists. Choose a new filename; no file will be overwritten.";
            if (string.Equals(outputFull, beforeFull, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(outputFull, afterFull, StringComparison.OrdinalIgnoreCase))
                return "The review destination cannot replace an input file.";
            return null;
        }

        private static string SafeDirectory(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path)) return "";
                string directory = Path.GetDirectoryName(Path.GetFullPath(path));
                return directory != null && Directory.Exists(directory) ? directory : "";
            }
            catch { return ""; }
        }

        private static string SafeFileName(string path)
        {
            try { return string.IsNullOrWhiteSpace(path) ? "" : Path.GetFileName(path); }
            catch { return ""; }
        }

        private static string FirstLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "Unknown error.";
            int line = text.IndexOfAny(new[] { '\r', '\n' });
            return line < 0 ? text : text.Substring(0, line);
        }

        private static string FormatBytes(long value)
        {
            string[] units = { "B", "KiB", "MiB", "GiB" };
            double display = value;
            int unit = 0;
            while (display >= 1024 && unit < units.Length - 1)
            {
                display /= 1024;
                unit++;
            }
            return display.ToString(unit == 0 ? "N0" : "N1", CultureInfo.CurrentCulture) + " " + units[unit];
        }

        private sealed class ComparisonFailureDialog : Form
        {
            private ComparisonFailureDialog(Exception exception)
            {
                Text = "Exact Compare Preview - detailed error";
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(720, 430);
                MinimumSize = new Size(560, 340);
                BackColor = Color.White;
                Font = new Font("Segoe UI", 9f);
                AutoScaleMode = AutoScaleMode.Dpi;

                var heading = new Label
                {
                    Dock = DockStyle.Top,
                    Height = 58,
                    Padding = new Padding(12, 11, 12, 4),
                    Text = "Comparison stopped; neither input was modified.\r\n" + FirstLine(exception.Message),
                    ForeColor = Color.FromArgb(142, 37, 45),
                    Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold),
                    UseCompatibleTextRendering = false,
                };
                var details = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Margin = Padding.Empty,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false,
                    Font = new Font("Consolas", 8.5f),
                    Text = exception.ToString(),
                };
                var buttons = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = 48,
                    Padding = new Padding(8),
                    BackColor = Surface,
                };
                var close = new Button
                {
                    Text = "Close",
                    DialogResult = DialogResult.OK,
                    Size = new Size(88, 29),
                    Dock = DockStyle.Right,
                };
                var copy = new Button
                {
                    Text = "Copy details",
                    Size = new Size(106, 29),
                    Dock = DockStyle.Left,
                };
                copy.Click += (_, __) =>
                {
                    try { Clipboard.SetText(details.Text); } catch { }
                };
                buttons.Controls.Add(copy);
                buttons.Controls.Add(close);
                Controls.Add(details);
                Controls.Add(heading);
                Controls.Add(buttons);
                AcceptButton = close;
                CancelButton = close;
            }

            internal static void Show(IWin32Window owner, Exception exception)
            {
                using (var dialog = new ComparisonFailureDialog(exception)) dialog.ShowDialog(owner);
            }
        }
    }
}
