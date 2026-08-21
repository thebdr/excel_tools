using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelPowerTools.Core.Comparison.Lineage;
using ExcelPowerTools.Core.Comparison.Model;

namespace ExcelPowerTools.UI
{
    internal sealed class ComparisonProjectArtifactInfo
    {
        internal ComparisonProjectArtifactInfo(
            string path,
            string revisionId,
            string byteContentId,
            long byteLength,
            DateTimeOffset? createdUtc,
            DateTimeOffset? modifiedUtc,
            string formatLabel)
        {
            Path = path ?? throw new ArgumentNullException(nameof(path));
            RevisionId = revisionId ?? throw new ArgumentNullException(nameof(revisionId));
            ByteContentId = byteContentId ?? throw new ArgumentNullException(nameof(byteContentId));
            ByteLength = byteLength;
            CreatedUtc = createdUtc;
            ModifiedUtc = modifiedUtc;
            FormatLabel = formatLabel ?? "";
        }

        internal string Path { get; }
        internal string RevisionId { get; }
        internal string ByteContentId { get; }
        internal long ByteLength { get; }
        internal DateTimeOffset? CreatedUtc { get; }
        internal DateTimeOffset? ModifiedUtc { get; }
        internal string FormatLabel { get; }
        internal string BaseName => System.IO.Path.GetFileName(Path);
    }

    internal sealed class ComparisonProjectSetupResult
    {
        internal ComparisonProjectSetupResult(
            IReadOnlyList<ComparisonProjectArtifactInfo> orderedArtifacts,
            LineageMappingGeneration mapping,
            string outputPath)
        {
            OrderedArtifacts = orderedArtifacts;
            Mapping = mapping;
            OutputPath = outputPath;
        }

        internal IReadOnlyList<ComparisonProjectArtifactInfo> OrderedArtifacts { get; }
        internal LineageMappingGeneration Mapping { get; }
        internal string OutputPath { get; }
    }

    /// <summary>
    /// Explicit, user-correctable setup for a multi-revision comparison. Filename/timestamp inference remains a
    /// proposal until the confirmation checkbox is checked; every correction is materialized through the core
    /// append-only lineage generation API.
    /// </summary>
    internal sealed class ComparisonProjectSetupForm : Form
    {
        private static readonly Color Background = Color.FromArgb(247, 248, 252);
        private static readonly Color Surface = Color.White;
        private static readonly Color Ink = Color.FromArgb(28, 31, 42);
        private static readonly Color Muted = Color.FromArgb(96, 102, 119);
        private static readonly Color Accent = Color.FromArgb(100, 73, 214);

        private readonly LineageInferenceResult _inference;
        private readonly List<RowModel> _rows;
        private readonly ToolTip _toolTips = new ToolTip
        {
            InitialDelay = 400,
            ReshowDelay = 100,
            AutoPopDelay = 15000,
            ShowAlways = true,
            UseAnimation = true,
            UseFading = true,
        };
        private readonly ListView _revisions = new ListView();
        private readonly TextBox _label = new TextBox();
        private readonly ComboBox _parent = new ComboBox();
        private readonly Label _selectionEvidence = new Label();
        private readonly Label _masterLabel = new Label();
        private readonly Label _baselineLabel = new Label();
        private readonly CheckBox _confirm = new CheckBox();
        private readonly TextBox _output = new TextBox();
        private readonly Button _run = new Button();
        private string _masterRevisionId;
        private string _baselineRevisionId;
        private bool _updatingSelection;

        internal ComparisonProjectSetupForm(
            IReadOnlyList<ComparisonProjectArtifactInfo> artifacts,
            LineageInferenceResult inference,
            string suggestedOutputPath)
        {
            if (artifacts == null) throw new ArgumentNullException(nameof(artifacts));
            if (artifacts.Count < 2) throw new ArgumentException("At least two revisions are required.", nameof(artifacts));
            _inference = inference ?? throw new ArgumentNullException(nameof(inference));
            var byId = artifacts.ToDictionary(item => item.RevisionId, StringComparer.Ordinal);
            _rows = inference.Proposals.Select((proposal, index) =>
            {
                ComparisonProjectArtifactInfo artifact = byId[proposal.RevisionId];
                string parent = proposal.ProposedParentRevisionId;
                if (index > 0 && string.IsNullOrEmpty(parent))
                    parent = inference.Proposals[index - 1].RevisionId;
                return new RowModel(
                    artifact,
                    proposal.LabelInference.BaseName,
                    parent,
                    proposal.OrderConfidence,
                    proposal.ParentConfidence,
                    proposal);
            }).ToList();
            _masterRevisionId = inference.ExplicitMasterRevisionId ?? _rows[_rows.Count - 1].Artifact.RevisionId;
            _baselineRevisionId = inference.ExplicitTrustedBaselineRevisionId;

            Text = "Comparison Project — revision setup";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(920, 650);
            Size = new Size(1120, 850);
            AutoScroll = true;
            AutoScrollMinSize = new Size(880, 760);
            BackColor = Background;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9F);
            KeyPreview = true;
            BuildLayout(suggestedOutputPath ?? "");
            HelpRequested += (_, args) =>
            {
                OpenGuide();
                args.Handled = true;
            };
            RefreshRows();
            SelectRow(0);
        }

        internal ComparisonProjectSetupResult Result { get; private set; }

        private void BuildLayout(string outputPath)
        {
            var title = new Label
            {
                Text = "Confirm revision lineage",
                Font = new Font("Segoe UI Semibold", 18F),
                ForeColor = Ink,
                Location = new Point(22, 18),
                AutoSize = true,
            };
            var subtitle = new Label
            {
                Text = "Order and parentage are separate. Correct weak filename guesses, then explicitly confirm the master and lineage.",
                ForeColor = Muted,
                Location = new Point(25, 55),
                AutoSize = true,
            };
            var guide = new Button
            {
                Text = "?",
                Font = new Font("Segoe UI Semibold", 11F),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Accent,
                BackColor = Surface,
                Location = new Point(ClientSize.Width - 58, 18),
                Size = new Size(34, 32),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                AccessibleName = "Open comparison setup guide",
            };
            guide.FlatAppearance.BorderColor = Color.FromArgb(206, 199, 239);
            guide.Click += (_, __) => OpenGuide();

            _revisions.Location = new Point(22, 88);
            _revisions.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _revisions.Size = new Size(ClientSize.Width - 44, 330);
            _revisions.View = View.Details;
            _revisions.FullRowSelect = true;
            _revisions.HideSelection = false;
            _revisions.MultiSelect = false;
            _revisions.BackColor = Surface;
            _revisions.BorderStyle = BorderStyle.FixedSingle;
            _revisions.Columns.Add("#", 42);
            _revisions.Columns.Add("Label", 160);
            _revisions.Columns.Add("File", 245);
            _revisions.Columns.Add("Parent", 150);
            _revisions.Columns.Add("Role", 150);
            _revisions.Columns.Add("Order confidence", 110);
            _revisions.Columns.Add("Parent confidence", 115);
            _revisions.SelectedIndexChanged += (_, __) => LoadSelectedRow();
            _revisions.ContextMenuStrip = BuildContextMenu();

            var earlier = Button("Move earlier", 22, 431, (_, __) => MoveSelected(-1));
            var later = Button("Move later", 132, 431, (_, __) => MoveSelected(1));
            var master = Button("Set master", 242, 431, (_, __) => SetMaster());
            var baseline = Button("Set baseline", 342, 431, (_, __) => SetBaseline());
            var clearBaseline = Button("Clear baseline", 447, 431, (_, __) => ClearBaseline());

            var editor = new Panel
            {
                BackColor = Surface,
                Location = new Point(22, 472),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Size = new Size(ClientSize.Width - 44, 92),
            };
            var labelCaption = Caption("Revision label", 12, 10);
            _label.Location = new Point(12, 31);
            _label.Size = new Size(220, 26);
            _label.TextChanged += (_, __) => CommitLabel();
            var parentCaption = Caption("Parent revision", 247, 10);
            _parent.DropDownStyle = ComboBoxStyle.DropDownList;
            _parent.Location = new Point(247, 31);
            _parent.Size = new Size(250, 26);
            _parent.SelectedIndexChanged += (_, __) => CommitParent();
            _selectionEvidence.Location = new Point(512, 10);
            _selectionEvidence.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _selectionEvidence.Size = new Size(editor.Width - 524, 68);
            _selectionEvidence.ForeColor = Muted;
            _selectionEvidence.AutoEllipsis = true;
            _masterLabel.Location = new Point(12, 64);
            _masterLabel.AutoSize = true;
            _masterLabel.ForeColor = Accent;
            _baselineLabel.Location = new Point(247, 64);
            _baselineLabel.AutoSize = true;
            _baselineLabel.ForeColor = Accent;
            editor.Controls.AddRange(new Control[]
            {
                labelCaption, _label, parentCaption, _parent, _selectionEvidence, _masterLabel, _baselineLabel,
            });

            var outputCaption = Caption("New Review.xlsx", 22, 580);
            _output.Location = new Point(22, 601);
            _output.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _output.Size = new Size(ClientSize.Width - 154, 26);
            _output.Text = outputPath;
            _output.TextChanged += (_, __) => MarkUnconfirmed();
            var browse = Button("Browse…", ClientSize.Width - 120, 599, (_, __) => BrowseOutput());
            browse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            browse.Width = 98;

            _confirm.Text = "I confirm the displayed order, parent links, master, and optional trusted baseline.";
            _confirm.Location = new Point(22, 640);
            _confirm.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _confirm.AutoSize = true;
            _confirm.CheckedChanged += (_, __) => _run.Enabled = _confirm.Checked;

            var limits = new Label
            {
                Text = "ALPHA: .xlsx/.xlsm only. AGGREGATE LIMITS: 2,000,000 physically stored cells; 10,000,000 content-bounded grid cells. PER SHEET: 100,000 rows; 4,096 columns; 5,000,000 grid cells. Each sheet stops at its last formula/value row and column; stored blank/formatting-only tails are ignored and reported. Ambiguous structure opens correction review; sources stay read-only.",
                ForeColor = Color.FromArgb(130, 76, 18),
                Location = new Point(22, 669),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Size = new Size(ClientSize.Width - 44, 66),
            };

            _run.Text = "Build comparison";
            _run.BackColor = Accent;
            _run.ForeColor = Color.White;
            _run.FlatStyle = FlatStyle.Flat;
            _run.FlatAppearance.BorderSize = 0;
            _run.Dock = DockStyle.Fill;
            _run.Enabled = false;
            _run.Click += (_, __) => Complete();
            var runHost = new Panel
            {
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Size = new Size(142, 34),
                Location = new Point(ClientSize.Width - 278, ClientSize.Height - 48),
            };
            runHost.Controls.Add(_run);
            var cancel = new Button
            {
                Text = "Cancel",
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Size = new Size(110, 34),
                Location = new Point(ClientSize.Width - 126, ClientSize.Height - 48),
                DialogResult = DialogResult.Cancel,
            };

            SetHelp(guide, ComparisonProjectSetupHelpText.OpenGuide);
            SetHelp(_revisions, ComparisonProjectSetupHelpText.RevisionList);
            SetHelp(earlier, ComparisonProjectSetupHelpText.MoveEarlier);
            SetHelp(later, ComparisonProjectSetupHelpText.MoveLater);
            SetHelp(master, ComparisonProjectSetupHelpText.SetMaster);
            SetHelp(baseline, ComparisonProjectSetupHelpText.SetBaseline);
            SetHelp(clearBaseline, ComparisonProjectSetupHelpText.ClearBaseline);
            SetHelp(_label, ComparisonProjectSetupHelpText.RevisionLabel);
            SetHelp(_parent, ComparisonProjectSetupHelpText.ParentRevision);
            SetHelp(_selectionEvidence, ComparisonProjectSetupHelpText.InferenceEvidence);
            SetHelp(_masterLabel, ComparisonProjectSetupHelpText.MasterStatus);
            SetHelp(_baselineLabel, ComparisonProjectSetupHelpText.BaselineStatus);
            SetHelp(_output, ComparisonProjectSetupHelpText.OutputPath);
            SetHelp(browse, ComparisonProjectSetupHelpText.BrowseOutput);
            SetHelp(_confirm, ComparisonProjectSetupHelpText.ConfirmLineage);
            SetHelp(limits, ComparisonProjectSetupHelpText.Limits);
            SetHelp(_run, ComparisonProjectSetupHelpText.BuildComparison);
            SetHelp(runHost, ComparisonProjectSetupHelpText.BuildComparison);
            SetHelp(cancel, ComparisonProjectSetupHelpText.Cancel);

            Controls.AddRange(new Control[]
            {
                title, subtitle, guide, _revisions, earlier, later, master, baseline, clearBaseline, editor,
                outputCaption, _output, browse, _confirm, limits, runHost, cancel,
            });
            CancelButton = cancel;
            Resize += (_, __) =>
            {
                _revisions.Width = Math.Max(700, ClientSize.Width - 44);
                editor.Width = Math.Max(700, ClientSize.Width - 44);
                _selectionEvidence.Width = Math.Max(170, editor.Width - 524);
                _output.Width = Math.Max(500, ClientSize.Width - 154);
                limits.Width = Math.Max(700, ClientSize.Width - 44);
            };
        }

        private ContextMenuStrip BuildContextMenu()
        {
            var menu = new ContextMenuStrip { ShowItemToolTips = true };
            AddContextAction(
                menu,
                "Move earlier",
                ComparisonProjectSetupHelpText.MoveEarlier,
                (_, __) => MoveSelected(-1));
            AddContextAction(
                menu,
                "Move later",
                ComparisonProjectSetupHelpText.MoveLater,
                (_, __) => MoveSelected(1));
            menu.Items.Add(new ToolStripSeparator());
            AddContextAction(
                menu,
                "Set as master",
                ComparisonProjectSetupHelpText.SetMaster,
                (_, __) => SetMaster());
            AddContextAction(
                menu,
                "Set as trusted baseline",
                ComparisonProjectSetupHelpText.SetBaseline,
                (_, __) => SetBaseline());
            AddContextAction(
                menu,
                "Clear trusted baseline",
                ComparisonProjectSetupHelpText.ClearBaseline,
                (_, __) => ClearBaseline());
            menu.Items.Add(new ToolStripSeparator());
            AddContextAction(
                menu,
                "Use previous revision as parent",
                ComparisonProjectSetupHelpText.PreviousParent,
                (_, __) => SetPreviousParent());
            AddContextAction(
                menu,
                "Clear parent",
                ComparisonProjectSetupHelpText.ClearParent,
                (_, __) => ClearParent());
            return menu;
        }

        private void OpenGuide()
        {
            using (var guide = new ComparisonProjectGuideForm())
                guide.ShowDialog(this);
        }

        private void SetHelp(Control control, string helpText)
        {
            _toolTips.SetToolTip(control, helpText);
            control.AccessibleDescription = helpText.Replace("\r\n", " ");
        }

        private static void AddContextAction(
            ContextMenuStrip menu,
            string text,
            string helpText,
            EventHandler action)
        {
            var item = new ToolStripMenuItem(text)
            {
                ToolTipText = helpText,
                AccessibleDescription = helpText.Replace("\r\n", " "),
            };
            item.Click += action;
            menu.Items.Add(item);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _toolTips.Dispose();
            base.Dispose(disposing);
        }

        private static Label Caption(string text, int x, int y) => new Label
        {
            Text = text,
            ForeColor = Muted,
            Location = new Point(x, y),
            AutoSize = true,
            Font = new Font("Segoe UI Semibold", 8.5F),
        };

        private static Button Button(string text, int x, int y, EventHandler action)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(100, 29),
                FlatStyle = FlatStyle.System,
            };
            button.Click += action;
            return button;
        }

        private void RefreshRows()
        {
            string selectedId = SelectedRow?.Artifact.RevisionId;
            _revisions.BeginUpdate();
            try
            {
                _revisions.Items.Clear();
                for (int index = 0; index < _rows.Count; index++)
                {
                    RowModel row = _rows[index];
                    string parent = ParentLabel(row.ParentRevisionId);
                    string role = Role(row.Artifact.RevisionId);
                    var item = new ListViewItem((index + 1).ToString(CultureInfo.InvariantCulture))
                    {
                        Tag = row.Artifact.RevisionId,
                    };
                    item.SubItems.Add(row.Label);
                    item.SubItems.Add(row.Artifact.BaseName);
                    item.SubItems.Add(parent);
                    item.SubItems.Add(role);
                    item.SubItems.Add(Confidence(row.OrderConfidence));
                    item.SubItems.Add(Confidence(row.ParentConfidence));
                    _revisions.Items.Add(item);
                }
            }
            finally
            {
                _revisions.EndUpdate();
            }
            UpdateRoleLabels();
            if (selectedId != null)
            {
                int index = _rows.FindIndex(row => row.Artifact.RevisionId == selectedId);
                if (index >= 0) SelectRow(index);
            }
        }

        private void LoadSelectedRow()
        {
            RowModel selected = SelectedRow;
            if (selected == null) return;
            _updatingSelection = true;
            try
            {
                _label.Text = selected.Label;
                _parent.Items.Clear();
                _parent.Items.Add(new ParentChoice("", "(root — no parent)"));
                int selectedIndex = _rows.IndexOf(selected);
                for (int index = 0; index < selectedIndex; index++)
                    _parent.Items.Add(new ParentChoice(
                        _rows[index].Artifact.RevisionId,
                        (index + 1).ToString(CultureInfo.InvariantCulture) + " · " + _rows[index].Label));
                int choice = 0;
                for (int index = 0; index < _parent.Items.Count; index++)
                    if (((ParentChoice)_parent.Items[index]).RevisionId == selected.ParentRevisionId)
                        choice = index;
                _parent.SelectedIndex = choice;
                _selectionEvidence.Text = Evidence(selected);
            }
            finally
            {
                _updatingSelection = false;
            }
        }

        private void CommitLabel()
        {
            if (_updatingSelection || SelectedRow == null) return;
            string value = _label.Text.Trim();
            if (value.Length == 0) return;
            if (SelectedRow.Label == value) return;
            SelectedRow.Label = value;
            MarkUnconfirmed();
            RefreshRows();
        }

        private void CommitParent()
        {
            if (_updatingSelection || SelectedRow == null || !(_parent.SelectedItem is ParentChoice choice)) return;
            if (SelectedRow.ParentRevisionId == choice.RevisionId) return;
            SelectedRow.ParentRevisionId = choice.RevisionId;
            MarkUnconfirmed();
            RefreshRows();
        }

        private void MoveSelected(int direction)
        {
            RowModel row = SelectedRow;
            if (row == null) return;
            int current = _rows.IndexOf(row);
            int target = current + direction;
            if (target < 0 || target >= _rows.Count) return;
            _rows.RemoveAt(current);
            _rows.Insert(target, row);
            RepairParentsAfterReorder();
            MarkUnconfirmed();
            RefreshRows();
            SelectRow(target);
        }

        private void RepairParentsAfterReorder()
        {
            var positions = _rows.Select((row, index) => new { row.Artifact.RevisionId, Index = index })
                .ToDictionary(item => item.RevisionId, item => item.Index, StringComparer.Ordinal);
            for (int index = 0; index < _rows.Count; index++)
            {
                RowModel row = _rows[index];
                if (row.ParentRevisionId.Length == 0) continue;
                if (!positions.TryGetValue(row.ParentRevisionId, out int parent) || parent >= index)
                    row.ParentRevisionId = index == 0 ? "" : _rows[index - 1].Artifact.RevisionId;
            }
            if (_rows[0].ParentRevisionId.Length != 0) _rows[0].ParentRevisionId = "";
        }

        private void SetMaster()
        {
            if (SelectedRow == null) return;
            _masterRevisionId = SelectedRow.Artifact.RevisionId;
            MarkUnconfirmed();
            RefreshRows();
        }

        private void SetBaseline()
        {
            if (SelectedRow == null) return;
            _baselineRevisionId = SelectedRow.Artifact.RevisionId;
            MarkUnconfirmed();
            RefreshRows();
        }

        private void ClearBaseline()
        {
            if (_baselineRevisionId == null) return;
            _baselineRevisionId = null;
            MarkUnconfirmed();
            RefreshRows();
        }

        private void SetPreviousParent()
        {
            RowModel row = SelectedRow;
            if (row == null) return;
            int index = _rows.IndexOf(row);
            string parent = index == 0 ? "" : _rows[index - 1].Artifact.RevisionId;
            if (row.ParentRevisionId == parent) return;
            row.ParentRevisionId = parent;
            MarkUnconfirmed();
            RefreshRows();
        }

        private void ClearParent()
        {
            if (SelectedRow == null || SelectedRow.ParentRevisionId.Length == 0) return;
            SelectedRow.ParentRevisionId = "";
            MarkUnconfirmed();
            RefreshRows();
        }

        private void BrowseOutput()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "Power Tools review workbook|*.xlsx|Excel workbook|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true,
                OverwritePrompt = false,
                Title = "Choose a new comparison Review.xlsx",
                FileName = Path.GetFileName(_output.Text),
                InitialDirectory = Directory.Exists(Path.GetDirectoryName(_output.Text))
                    ? Path.GetDirectoryName(_output.Text)
                    : "",
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK) _output.Text = dialog.FileName;
            }
        }

        private void Complete()
        {
            string validation = ValidateSetup();
            if (validation != null)
            {
                ExcelPowerTools.Interop.Dialogs.Info(validation, "Comparison Project", this);
                return;
            }

            try
            {
                LineageMappingGeneration mapping = BuildMappingGeneration();
                Result = new ComparisonProjectSetupResult(
                    _rows.Select(row => row.Artifact).ToList().AsReadOnly(),
                    mapping,
                    Path.GetFullPath(_output.Text.Trim()));
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                ExcelPowerTools.Interop.Dialogs.Info(
                    "The confirmed lineage could not be materialized.\r\n\r\n" + exception.Message,
                    "Comparison Project", this);
            }
        }

        private string ValidateSetup()
        {
            if (!_confirm.Checked) return "Confirm the displayed lineage before building the comparison.";
            if (string.IsNullOrEmpty(_masterRevisionId)) return "Select and confirm the master revision.";
            if (_rows.Count(row => row.ParentRevisionId.Length == 0) != 1)
                return "The current engine requires one rooted revision tree. Choose exactly one root revision.";
            var positions = _rows.Select((row, index) => new { row.Artifact.RevisionId, Index = index })
                .ToDictionary(item => item.RevisionId, item => item.Index, StringComparer.Ordinal);
            for (int index = 0; index < _rows.Count; index++)
            {
                string parent = _rows[index].ParentRevisionId;
                if (parent.Length == 0) continue;
                if (!positions.TryGetValue(parent, out int parentIndex) || parentIndex >= index)
                    return "Every parent must precede its child. Review revision " + _rows[index].Label + ".";
            }

            string output = _output.Text.Trim();
            if (output.Length == 0) return "Choose a new Review.xlsx destination.";
            string full;
            try { full = Path.GetFullPath(output); }
            catch { return "The Review.xlsx destination is invalid."; }
            if (!string.Equals(Path.GetExtension(full), ".xlsx", StringComparison.OrdinalIgnoreCase))
                return "The review destination must end in .xlsx.";
            if (File.Exists(full) || Directory.Exists(full))
                return "The review destination already exists. Choose a new filename.";
            if (!Directory.Exists(Path.GetDirectoryName(full)))
                return "The review destination directory does not exist.";
            if (_rows.Any(row => string.Equals(Path.GetFullPath(row.Artifact.Path), full, StringComparison.OrdinalIgnoreCase)))
                return "The new review destination cannot be one of the source files.";
            return null;
        }

        private LineageMappingGeneration BuildMappingGeneration()
        {
            ReviewerSnapshot reviewer = ReviewerSnapshot.Empty;
            DateTimeOffset now = DateTimeOffset.UtcNow;
            LineageMappingGeneration mapping = LineageCorrectionEngine.FromInference(_inference);

            foreach (string revisionId in mapping.OrderedRevisionIds)
                if (!string.IsNullOrEmpty(mapping.GetParent(revisionId)))
                    mapping = LineageCorrectionEngine.Apply(
                        mapping,
                        LineageCorrectionRequest.ClearParent(revisionId, reviewer, now, "Confirmed setup rebuild"));

            string[] desiredOrder = _rows.Select(row => row.Artifact.RevisionId).ToArray();
            if (!mapping.OrderedRevisionIds.SequenceEqual(desiredOrder, StringComparer.Ordinal))
                mapping = LineageCorrectionEngine.Apply(
                    mapping,
                    LineageCorrectionRequest.ReplaceOrder(desiredOrder, reviewer, now, "User-confirmed revision order"));

            foreach (RowModel row in _rows)
                if (row.ParentRevisionId.Length != 0)
                    mapping = LineageCorrectionEngine.Apply(
                        mapping,
                        LineageCorrectionRequest.SetParent(
                            row.Artifact.RevisionId,
                            row.ParentRevisionId,
                            reviewer,
                            now,
                            "User-confirmed parent"));

            foreach (RowModel row in _rows)
                if (!string.Equals(mapping.GetLabel(row.Artifact.RevisionId), row.Label, StringComparison.Ordinal))
                    mapping = LineageCorrectionEngine.Apply(
                        mapping,
                        LineageCorrectionRequest.SetLabel(
                            row.Artifact.RevisionId,
                            row.Label,
                            reviewer,
                            now,
                            "User-confirmed label"));

            if (!string.Equals(mapping.MasterRevisionId, _masterRevisionId, StringComparison.Ordinal))
                mapping = LineageCorrectionEngine.Apply(
                    mapping,
                    LineageCorrectionRequest.SetMaster(
                        _masterRevisionId,
                        reviewer,
                        now,
                        "User-confirmed master"));
            if (_baselineRevisionId != null &&
                !string.Equals(mapping.TrustedBaselineRevisionId, _baselineRevisionId, StringComparison.Ordinal))
                mapping = LineageCorrectionEngine.Apply(
                    mapping,
                    LineageCorrectionRequest.SetTrustedBaseline(
                        _baselineRevisionId,
                        reviewer,
                        now,
                        "User-confirmed trusted baseline"));
            if (_baselineRevisionId == null && mapping.TrustedBaselineRevisionId != null)
                mapping = LineageCorrectionEngine.Apply(
                    mapping,
                    LineageCorrectionRequest.ClearTrustedBaseline(reviewer, now, "Trusted baseline cleared"));
            return mapping;
        }

        private RowModel SelectedRow
        {
            get
            {
                if (_revisions.SelectedItems.Count != 1) return null;
                string id = _revisions.SelectedItems[0].Tag as string;
                return _rows.FirstOrDefault(row => row.Artifact.RevisionId == id);
            }
        }

        private void SelectRow(int index)
        {
            if (index < 0 || index >= _revisions.Items.Count) return;
            _revisions.Items[index].Selected = true;
            _revisions.Items[index].Focused = true;
            _revisions.EnsureVisible(index);
        }

        private void MarkUnconfirmed()
        {
            if (_updatingSelection) return;
            _confirm.Checked = false;
        }

        private string ParentLabel(string revisionId)
        {
            if (string.IsNullOrEmpty(revisionId)) return "(root)";
            RowModel row = _rows.FirstOrDefault(item => item.Artifact.RevisionId == revisionId);
            return row == null ? "(missing)" : row.Label;
        }

        private string Role(string revisionId)
        {
            var roles = new List<string>();
            if (revisionId == _masterRevisionId) roles.Add("Master");
            if (revisionId == _baselineRevisionId) roles.Add("Trusted baseline");
            return string.Join(" + ", roles);
        }

        private void UpdateRoleLabels()
        {
            _masterLabel.Text = "Master: " + ParentLabel(_masterRevisionId);
            _baselineLabel.Text = "Trusted baseline: " +
                                  (_baselineRevisionId == null ? "(not set)" : ParentLabel(_baselineRevisionId));
        }

        private static string Confidence(LineageConfidence confidence) =>
            confidence.Level + " " + confidence.Score.ToString("0.00", CultureInfo.InvariantCulture);

        private static string Evidence(RowModel row)
        {
            string order = row.Proposal.OrderReasons.Count == 0
                ? "No strong order evidence."
                : string.Join(" ", row.Proposal.OrderReasons.Select(reason => reason.Message));
            string parent = row.Proposal.ParentReasons.Count == 0
                ? "No strong parent evidence."
                : string.Join(" ", row.Proposal.ParentReasons.Select(reason => reason.Message));
            return order + " " + parent;
        }

        private sealed class ParentChoice
        {
            internal ParentChoice(string revisionId, string label)
            {
                RevisionId = revisionId;
                Label = label;
            }

            internal string RevisionId { get; }
            internal string Label { get; }
            public override string ToString() => Label;
        }

        private sealed class RowModel
        {
            internal RowModel(
                ComparisonProjectArtifactInfo artifact,
                string label,
                string parentRevisionId,
                LineageConfidence orderConfidence,
                LineageConfidence parentConfidence,
                LineageRevisionProposal proposal)
            {
                Artifact = artifact;
                Label = label;
                ParentRevisionId = parentRevisionId ?? "";
                OrderConfidence = orderConfidence;
                ParentConfidence = parentConfidence;
                Proposal = proposal;
            }

            internal ComparisonProjectArtifactInfo Artifact { get; }
            internal string Label { get; set; }
            internal string ParentRevisionId { get; set; }
            internal LineageConfidence OrderConfidence { get; }
            internal LineageConfidence ParentConfidence { get; }
            internal LineageRevisionProposal Proposal { get; }
        }
    }
}
