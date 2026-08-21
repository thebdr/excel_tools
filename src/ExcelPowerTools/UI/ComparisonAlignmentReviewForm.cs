using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Comparison.Alignment;
using ExcelPowerTools.Core.Comparison.Diagnostics;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Project;
using ExcelPowerTools.Core.Comparison.Projection;

namespace ExcelPowerTools.UI
{
    internal enum ComparisonAlignmentReviewAction
    {
        Reanalyze,
        Publish,
        PublishDiagnostic,
        Cancel,
    }

    internal sealed class ComparisonAlignmentReviewResult
    {
        internal ComparisonAlignmentReviewResult(
            ComparisonAlignmentReviewAction action,
            AlignmentCorrectionGeneration corrections)
        {
            Action = action;
            Corrections = corrections ?? throw new ArgumentNullException(nameof(corrections));
        }

        internal ComparisonAlignmentReviewAction Action { get; }
        internal AlignmentCorrectionGeneration Corrections { get; }
    }

    /// <summary>
    /// A command-lifetime editor for exact alignment confirmations. Every editable scope is recovered from the
    /// structured project edge, mapping, and alignment objects; diagnostic prose is display-only and is never
    /// parsed to identify a revision, sheet, row, or column. Grids are retained only while this form is alive.
    /// </summary>
    internal sealed class ComparisonAlignmentReviewForm : Form
    {
        private static readonly Color Background = Color.FromArgb(246, 247, 251);
        private static readonly Color Surface = Color.White;
        private static readonly Color Ink = Color.FromArgb(27, 30, 41);
        private static readonly Color Muted = Color.FromArgb(94, 100, 116);
        private static readonly Color Accent = Color.FromArgb(99, 72, 211);
        private static readonly Color Warning = Color.FromArgb(153, 86, 16);
        private static readonly Color Danger = Color.FromArgb(177, 48, 62);

        private readonly ComparisonProjectAnalysisResult _analysis;
        private readonly List<IssueModel> _issues;
        private readonly ToolTip _toolTips = new ToolTip
        {
            InitialDelay = 400,
            ReshowDelay = 100,
            AutoPopDelay = 15000,
            ShowAlways = true,
            UseAnimation = true,
            UseFading = true,
        };
        private readonly Dictionary<string, CanonicalGridProjection> _grids =
            new Dictionary<string, CanonicalGridProjection>(StringComparer.Ordinal);
        private readonly ListView _issueList = new ListView();
        private readonly Panel _editorHost = new Panel();
        private readonly TextBox _reviewer = new TextBox();
        private readonly TextBox _note = new TextBox();
        private readonly Label _workingStatus = new Label();
        private readonly Button _reanalyze = new Button();
        private AlignmentCorrectionGeneration _working;
        private bool _closingWithResult;

        internal ComparisonAlignmentReviewForm(ComparisonProjectAnalysisResult analysis)
        {
            _analysis = analysis ?? throw new ArgumentNullException(nameof(analysis));
            _working = analysis.AlignmentCorrections;
            _issues = BuildIssues(analysis);

            Text = "Comparison Project — alignment review";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1040, 700);
            Size = new Size(1380, 900);
            BackColor = Background;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9F);
            KeyPreview = true;
            BuildLayout();
            PopulateIssues();
            if (_issueList.Items.Count > 0)
            {
                ListViewItem initial = _issueList.Items.Cast<ListViewItem>().FirstOrDefault(item =>
                    item.Tag is IssueModel issue && issue.EditorKind != IssueEditorKind.ReadOnly &&
                    issue.Finding.Severity >= DiagnosticSeverity.Correctable)
                    ?? _issueList.Items.Cast<ListViewItem>().FirstOrDefault(item =>
                        item.Tag is IssueModel issue && issue.EditorKind != IssueEditorKind.ReadOnly)
                    ?? _issueList.Items[0];
                initial.Selected = true;
                initial.Focused = true;
                initial.EnsureVisible();
            }
            else
            {
                ShowEmptyEditor();
            }
            RefreshWorkingStatus();
        }

        internal ComparisonAlignmentReviewResult Result { get; private set; }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_closingWithResult && Result == null)
                Result = new ComparisonAlignmentReviewResult(
                    ComparisonAlignmentReviewAction.Cancel,
                    _analysis.AlignmentCorrections);
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _grids.Clear();
            base.OnFormClosed(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _toolTips.Dispose();
            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Background,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(18, 14, 18, 12),
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 82F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F));

            var heading = new Panel { Dock = DockStyle.Fill, BackColor = Background };
            heading.Controls.Add(new Label
            {
                Text = "Resolve alignment evidence",
                Font = new Font("Segoe UI Semibold", 18F),
                ForeColor = Ink,
                Location = new Point(2, 0),
                AutoSize = true,
            });
            heading.Controls.Add(new Label
            {
                Text = "All findings stay visible. Queue only exact sheet, header, or key confirmations, then rerun analysis over the retained input snapshots.",
                ForeColor = Muted,
                Location = new Point(5, 38),
                AutoSize = true,
            });

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Size = new Size(1120, 620),
                SplitterDistance = 470,
                SplitterWidth = 7,
                BackColor = Background,
                Panel1MinSize = 360,
                Panel2MinSize = 500,
            };
            BuildIssuePane(split.Panel1);
            _editorHost.Dock = DockStyle.Fill;
            _editorHost.BackColor = Surface;
            _editorHost.BorderStyle = BorderStyle.FixedSingle;
            _editorHost.AutoScroll = true;
            SetHelp(_editorHost, ComparisonAlignmentReviewHelpText.Editor);
            split.Panel2.Padding = new Padding(1, 0, 0, 0);
            split.Panel2.Controls.Add(_editorHost);

            var audit = BuildAuditPanel();
            var footer = BuildFooter();

            root.Controls.Add(heading, 0, 0);
            root.Controls.Add(split, 0, 1);
            root.Controls.Add(audit, 0, 2);
            root.Controls.Add(footer, 0, 3);
            Controls.Add(root);
        }

        private void BuildIssuePane(Control parent)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Surface,
                Padding = new Padding(10),
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            panel.Controls.Add(new Label
            {
                Text = "Findings  ·  Ctrl/Shift-click for multiple",
                Font = new Font("Segoe UI Semibold", 10F),
                ForeColor = Ink,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
            }, 0, 0);

            _issueList.Dock = DockStyle.Fill;
            _issueList.View = View.Details;
            _issueList.FullRowSelect = true;
            _issueList.HideSelection = false;
            _issueList.MultiSelect = true;
            _issueList.ShowGroups = true;
            _issueList.ShowItemToolTips = true;
            _issueList.BorderStyle = BorderStyle.FixedSingle;
            _issueList.BackColor = Surface;
            _issueList.Columns.Add("Status", 92);
            _issueList.Columns.Add("Finding", 210);
            _issueList.Columns.Add("Severity", 90);
            _issueList.SelectedIndexChanged += (_, __) => ShowSelectedEditor();
            _issueList.ContextMenuStrip = BuildIssueContextMenu();
            SetHelp(_issueList, ComparisonAlignmentReviewHelpText.Findings);
            panel.Controls.Add(_issueList, 0, 1);
            parent.Controls.Add(panel);
        }

        private Control BuildAuditPanel()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2,
                BackColor = Background,
                Padding = new Padding(2, 10, 2, 4),
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240F));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62F));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38F));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 0F));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 29F));
            panel.Controls.Add(Caption("Reviewer name (optional)"), 0, 0);
            panel.Controls.Add(Caption("Audit note (optional; applied to the next queued batch)"), 1, 0);
            _workingStatus.Dock = DockStyle.Fill;
            _workingStatus.ForeColor = Muted;
            _workingStatus.TextAlign = ContentAlignment.BottomRight;
            panel.Controls.Add(_workingStatus, 2, 0);
            _reviewer.Dock = DockStyle.Fill;
            _note.Dock = DockStyle.Fill;
            SetHelp(_reviewer, ComparisonAlignmentReviewHelpText.Reviewer);
            SetHelp(_note, ComparisonAlignmentReviewHelpText.AuditNote);
            SetHelp(_workingStatus, ComparisonAlignmentReviewHelpText.WorkingStatus);
            panel.Controls.Add(_reviewer, 0, 1);
            panel.Controls.Add(_note, 1, 1);
            panel.SetColumnSpan(_note, 2);
            return panel;
        }

        private Control BuildFooter()
        {
            var footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Background,
                Padding = new Padding(0, 7, 0, 0),
            };
            var cancel = FooterButton("Cancel", 106, (_, __) => Finish(ComparisonAlignmentReviewAction.Cancel));
            var publish = FooterButton(
                _analysis.ProjectResult.IsComplete ? "Continue to Review" : "Build diagnostic Review",
                178,
                (_, __) =>
                {
                    if (_analysis.ProjectResult.IsComplete)
                        Finish(ComparisonAlignmentReviewAction.Publish);
                    else
                        ConfirmDiagnostic();
                });
            _reanalyze.Text = "Rerun analysis";
            _reanalyze.Size = new Size(142, 34);
            _reanalyze.BackColor = Accent;
            _reanalyze.ForeColor = Color.White;
            _reanalyze.FlatStyle = FlatStyle.Flat;
            _reanalyze.FlatAppearance.BorderSize = 0;
            _reanalyze.Click += (_, __) => Finish(ComparisonAlignmentReviewAction.Reanalyze);
            var reanalyzeHost = new Panel
            {
                Size = _reanalyze.Size,
                Margin = new Padding(8, 0, 0, 0),
            };
            _reanalyze.Dock = DockStyle.Fill;
            _reanalyze.Margin = Padding.Empty;
            reanalyzeHost.Controls.Add(_reanalyze);
            SetHelp(cancel, ComparisonAlignmentReviewHelpText.Cancel);
            SetHelp(
                publish,
                _analysis.ProjectResult.IsComplete
                    ? ComparisonAlignmentReviewHelpText.ContinueToReview
                    : ComparisonAlignmentReviewHelpText.BuildDiagnosticReview);
            SetHelp(_reanalyze, ComparisonAlignmentReviewHelpText.RerunAnalysis);
            SetHelp(reanalyzeHost, ComparisonAlignmentReviewHelpText.RerunAnalysis);
            footer.Controls.Add(cancel);
            footer.Controls.Add(publish);
            footer.Controls.Add(reanalyzeHost);
            CancelButton = cancel;
            return footer;
        }

        private ContextMenuStrip BuildIssueContextMenu()
        {
            var menu = new ContextMenuStrip { ShowItemToolTips = true };
            var queue = new ToolStripMenuItem("Queue safe recommendation for selected")
            {
                ToolTipText = ComparisonAlignmentReviewHelpText.QueueSafeRecommendations,
                AccessibleDescription = ComparisonAlignmentReviewHelpText.QueueSafeRecommendations,
            };
            queue.Click += (_, __) => QueueRecommendationsForSelection();
            menu.Items.Add(queue);
            var open = new ToolStripMenuItem("Open primary selection")
            {
                ToolTipText = ComparisonAlignmentReviewHelpText.OpenPrimarySelection,
                AccessibleDescription = ComparisonAlignmentReviewHelpText.OpenPrimarySelection,
            };
            open.Click += (_, __) => ShowSelectedEditor();
            menu.Items.Add(open);
            menu.Opening += (_, e) =>
            {
                menu.Items[0].Enabled = SelectedIssues().Any(issue => issue.HasSafeRecommendation);
                menu.Items[1].Enabled = _issueList.SelectedItems.Count > 0;
                if (_issueList.SelectedItems.Count == 0) e.Cancel = true;
            };
            return menu;
        }

        private void PopulateIssues()
        {
            _issueList.BeginUpdate();
            try
            {
                _issueList.Items.Clear();
                _issueList.Groups.Clear();
                var groups = new Dictionary<string, ListViewGroup>(StringComparer.Ordinal);
                foreach (IssueModel issue in _issues)
                {
                    string groupKey = issue.GroupKey;
                    if (!groups.TryGetValue(groupKey, out ListViewGroup group))
                    {
                        group = new ListViewGroup(groupKey, issue.GroupLabel);
                        groups.Add(groupKey, group);
                        _issueList.Groups.Add(group);
                    }
                    var item = new ListViewItem(Status(issue), group)
                    {
                        Tag = issue,
                        ForeColor = SeverityColor(issue.Finding.Severity),
                    };
                    item.SubItems.Add(FindingLabel(issue));
                    item.SubItems.Add(issue.Finding.Severity.ToString());
                    item.ToolTipText = issue.Finding.Message;
                    _issueList.Items.Add(item);
                }
            }
            finally
            {
                _issueList.EndUpdate();
            }
        }

        private void RefreshIssueStatuses()
        {
            foreach (ListViewItem item in _issueList.Items)
                if (item.Tag is IssueModel issue)
                    item.SubItems[0].Text = Status(issue);
            RefreshWorkingStatus();
        }

        private void ShowSelectedEditor()
        {
            IssueModel issue = SelectedIssues().FirstOrDefault();
            if (issue == null)
            {
                ShowEmptyEditor();
                return;
            }
            try
            {
                switch (issue.EditorKind)
                {
                    case IssueEditorKind.SheetPair:
                        ShowSheetPairEditor(issue);
                        break;
                    case IssueEditorKind.Header:
                        ShowHeaderEditor(issue);
                        break;
                    case IssueEditorKind.SingleHeader:
                        ShowSingleHeaderEditor(issue);
                        break;
                    case IssueEditorKind.Key:
                        ShowKeyEditor(issue);
                        break;
                    default:
                        ShowReadOnlyEditor(issue);
                        break;
                }
            }
            catch (Exception exception)
            {
                ShowReadOnlyEditor(
                    issue,
                    "The structured editor could not be prepared safely: " + exception.Message);
            }
        }

        private void ShowEmptyEditor()
        {
            ReplaceEditor(new Label
            {
                Text = "Select a finding to inspect its exact scope and available correction.",
                ForeColor = Muted,
                AutoSize = true,
                Location = new Point(24, 24),
            });
        }

        private void ShowSheetPairEditor(IssueModel issue)
        {
            UnresolvedSheetMapping unresolved = issue.UnresolvedSheet;
            WorkbookSheetFingerprint source = issue.Edge?.SheetMapping.Before.Fingerprints.SingleOrDefault(item =>
                string.Equals(item.SheetKey, issue.BeforeSheetKey, StringComparison.Ordinal));
            if (source == null || issue.Edge == null)
            {
                ShowReadOnlyEditor(issue, "No exact source sheet occurrence is attached to this confirmation.");
                return;
            }

            var page = EditorPage(issue, "Resolve this sheet occurrence");
            page.Controls.Add(Paragraph(
                "Source occurrence: " + source.Name + "  [" + ShortId(source.SheetKey) + "]\r\n" +
                "Choose its target in " + RevisionLabel(issue.Edge.ChildRevisionId) +
                ", or choose deleted + new when it has no successor.\r\n" +
                "Both deletion and insertion remain visible in structural history."));

            var target = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Top,
                IntegralHeight = false,
                MaxDropDownItems = 14,
                Height = 28,
            };
            SetHelp(target, ComparisonAlignmentReviewHelpText.TargetSheet);
            IReadOnlyList<TargetChoice> targets = SheetTargets(issue);
            foreach (TargetChoice choice in targets) target.Items.Add(choice);

            AlignmentSheetPairConfirmation active = ActiveSheetPair(_working, issue);
            int selected = -1;
            if (active != null)
                selected = targets.ToList().FindIndex(choice =>
                    string.Equals(choice.Fingerprint.SheetKey, active.AfterSheetKey, StringComparison.Ordinal));
            if (selected < 0 && unresolved?.Candidates.Count == 1)
                selected = targets.ToList().FindIndex(choice =>
                    string.Equals(
                        choice.Fingerprint.SheetKey,
                        unresolved.Candidates[0].After.SheetKey,
                        StringComparison.Ordinal));
            if (selected < 0 && targets.Count > 0) selected = 0;
            target.SelectedIndex = selected;

            var selectionBox = new Panel { Height = 71, Dock = DockStyle.Top, Padding = new Padding(0, 4, 0, 7) };
            selectionBox.Controls.Add(target);
            selectionBox.Controls.Add(new Label
            {
                Text = "Target sheet occurrence",
                Dock = DockStyle.Top,
                Height = 23,
                Font = new Font("Segoe UI Semibold", 9F),
                ForeColor = Ink,
            });
            page.Controls.Add(selectionBox);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 42,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
            };
            var queue = ActionButton("Queue sheet pair", 142);
            queue.Click += (_, __) =>
            {
                if (!(target.SelectedItem is TargetChoice choice))
                {
                    Inform("No eligible target sheet is available. Use the deleted + new option when the source has no successor.");
                    return;
                }
                ApplySheetPair(issue, choice.Fingerprint.SheetKey);
            };
            SetHelp(queue, ComparisonAlignmentReviewHelpText.QueueSheetPair);
            var unmatched = ActionButton("Treat as deleted + new", 184);
            unmatched.Click += (_, __) => ApplySheetUnmatched(issue);
            SetHelp(unmatched, ComparisonAlignmentReviewHelpText.KeepSheetsUnmatched);
            var clear = ActionButton("Use automatic mapping", 164);
            clear.Click += (_, __) => ClearSheetResolution(issue);
            SetHelp(clear, ComparisonAlignmentReviewHelpText.AutomaticSheetMapping);
            buttons.Controls.Add(queue);
            buttons.Controls.Add(unmatched);
            buttons.Controls.Add(clear);
            page.Controls.Add(buttons);
            page.Controls.Add(Details(issue));
            ReplaceEditor(page);
        }

        private void ShowHeaderEditor(IssueModel issue)
        {
            if (issue.Alignment == null || issue.Edge == null ||
                !TryGetAlignmentEvidence(issue, out OpenXmlSheetFingerprintEvidence beforeEvidence,
                    out OpenXmlSheetFingerprintEvidence afterEvidence))
            {
                ShowReadOnlyEditor(issue, "No exact mapped sheet pair is attached to this header finding.");
                return;
            }

            CanonicalGridProjection beforeGrid = Grid(issue.Edge.ParentRevisionId, beforeEvidence);
            CanonicalGridProjection afterGrid = Grid(issue.Edge.ChildRevisionId, afterEvidence);
            var page = EditorPage(issue, "Confirm paired header decisions");
            page.Controls.Add(Paragraph(
                "Coordinates below are physical Excel row numbers. Auto displays the current detector decision but " +
                "does not turn it into a confirmation. For separated subheader/unit rows, set the optional first data row."));

            var pair = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 386,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(0, 5, 0, 4),
            };
            pair.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            pair.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            var before = new HeaderSideEditor(
                "Before · " + RevisionLabel(issue.Edge.ParentRevisionId) + " · " + beforeEvidence.Fingerprint.Name,
                beforeGrid,
                issue.Alignment.BeforeHeader,
                ActiveHeader(_working, issue.Edge.ParentRevisionId, issue.BeforeSheetKey),
                _toolTips);
            var after = new HeaderSideEditor(
                "After · " + RevisionLabel(issue.Edge.ChildRevisionId) + " · " + afterEvidence.Fingerprint.Name,
                afterGrid,
                issue.Alignment.AfterHeader,
                ActiveHeader(_working, issue.Edge.ChildRevisionId, issue.AfterSheetKey),
                _toolTips);
            pair.Controls.Add(before.Control, 0, 0);
            pair.Controls.Add(after.Control, 1, 0);
            page.Controls.Add(pair);

            var apply = ActionButton("Queue paired decisions", 184);
            apply.Dock = DockStyle.Top;
            apply.Click += (_, __) => ApplyHeaders(issue, before, after);
            SetHelp(apply, ComparisonAlignmentReviewHelpText.QueuePairedHeaders);
            page.Controls.Add(apply);
            page.Controls.Add(Details(issue));
            ReplaceEditor(page);
        }

        private void ShowSingleHeaderEditor(IssueModel issue)
        {
            OpenXmlSheetFingerprintEvidence evidence = FindEvidence(
                issue.SingleRevisionId,
                issue.SingleSheetKey);
            if (evidence == null)
            {
                ShowReadOnlyEditor(issue, "The exact sheet occurrence for this active header confirmation is unavailable.");
                return;
            }

            CanonicalGridProjection grid = Grid(issue.SingleRevisionId, evidence);
            var page = EditorPage(issue, "Edit an active header confirmation");
            page.Controls.Add(Paragraph(
                "This confirmation is not attached to one uniquely editable comparison edge. Coordinates are still " +
                "the exact physical Excel rows for this revision occurrence. Auto clears only this active override; " +
                "the next analysis may require a paired confirmation on a related revision."));
            var editor = new HeaderSideEditor(
                RevisionLabel(issue.SingleRevisionId) + " · " + evidence.Fingerprint.Name,
                grid,
                issue.SingleHeaderDetection,
                ActiveHeader(_working, issue.SingleRevisionId, issue.SingleSheetKey),
                _toolTips);
            var host = new Panel { Dock = DockStyle.Top, Height = 386, Padding = new Padding(0, 5, 0, 4) };
            host.Controls.Add(editor.Control);
            page.Controls.Add(host);
            var apply = ActionButton("Queue header decision", 176);
            apply.Dock = DockStyle.Top;
            apply.Click += (_, __) =>
            {
                try
                {
                    if (editor.Choice == HeaderEditorChoice.Header && grid.Table.RowCount == 0)
                        throw new InvalidOperationException(
                            "An empty physical grid cannot carry a confirmed header row. Choose No header or Auto.");
                    var requests = new List<AlignmentCorrectionRequest>();
                    AddHeaderRequest(
                        requests,
                        issue.SingleRevisionId,
                        issue.SingleSheetKey,
                        editor);
                    if (requests.Count == 0)
                    {
                        Inform("No header confirmation changed.");
                        return;
                    }
                    TryApplyBatch(requests);
                }
                catch (Exception exception)
                {
                    Error("The header decision is not valid.", exception);
                }
            };
            SetHelp(apply, ComparisonAlignmentReviewHelpText.QueueSingleHeader);
            page.Controls.Add(apply);
            page.Controls.Add(Details(issue));
            ReplaceEditor(page);
        }

        private void ShowKeyEditor(IssueModel issue)
        {
            if (issue.Edge == null ||
                !TryGetAlignmentEvidence(issue, out OpenXmlSheetFingerprintEvidence beforeEvidence,
                    out OpenXmlSheetFingerprintEvidence afterEvidence))
            {
                ShowReadOnlyEditor(issue, "No exact mapped sheet pair is attached to this row-key finding.");
                return;
            }

            CanonicalGridProjection beforeGrid = Grid(issue.Edge.ParentRevisionId, beforeEvidence);
            CanonicalGridProjection afterGrid = Grid(issue.Edge.ChildRevisionId, afterEvidence);
            IReadOnlyList<ColumnChoice> beforeColumns = ColumnChoices(beforeGrid, issue.Alignment?.BeforeHeader);
            IReadOnlyList<ColumnChoice> afterColumns = ColumnChoices(afterGrid, issue.Alignment?.AfterHeader);
            AlignmentKeyColumnsConfirmation active = ActiveKey(_working, issue);
            KeyCandidate suggested = active == null ? SuggestedKey(issue.Alignment?.KeyInference) : null;

            var page = EditorPage(issue, "Confirm ordered key columns");
            page.Controls.Add(Paragraph(
                "Use Auto to keep inference in control, or confirm one to four ordered column pairs. Column letters and " +
                "numbers are physical Excel coordinates; statistics are exact for up to the first 5,000 data rows and labelled when sampled."));

            var auto = new RadioButton
            {
                Text = "Auto / no explicit key (current inference remains read-only)",
                AutoSize = true,
                Checked = active == null,
                Margin = new Padding(2, 5, 2, 3),
            };
            var explicitKey = new RadioButton
            {
                Text = "Confirm these ordered Before → After column pairs",
                AutoSize = true,
                Checked = active != null,
                Margin = new Padding(2, 2, 2, 4),
            };
            SetHelp(auto, ComparisonAlignmentReviewHelpText.AutomaticKey);
            SetHelp(explicitKey, ComparisonAlignmentReviewHelpText.ExplicitKey);
            page.Controls.Add(auto);
            page.Controls.Add(explicitKey);

            var pairs = new List<KeyPairEditor>();
            var pairTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 184,
                ColumnCount = 3,
                RowCount = 5,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                BackColor = Color.FromArgb(236, 238, 245),
                Padding = new Padding(1),
            };
            pairTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 42F));
            pairTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            pairTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            pairTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 31F));
            for (int row = 1; row <= 4; row++) pairTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 37F));
            pairTable.Controls.Add(GridHeader("#"), 0, 0);
            pairTable.Controls.Add(GridHeader("Before physical column"), 1, 0);
            pairTable.Controls.Add(GridHeader("After physical column"), 2, 0);

            for (int index = 0; index < 4; index++)
            {
                var editor = new KeyPairEditor(index + 1, beforeColumns, afterColumns, _toolTips);
                pairs.Add(editor);
                pairTable.Controls.Add(GridHeader((index + 1).ToString(CultureInfo.InvariantCulture)), 0, index + 1);
                pairTable.Controls.Add(editor.Before, 1, index + 1);
                pairTable.Controls.Add(editor.After, 2, index + 1);
            }
            InitializeKeyPairs(pairs, active, suggested);
            foreach (KeyPairEditor editor in pairs)
            {
                editor.Before.SelectedIndexChanged += (_, __) => explicitKey.Checked = true;
                editor.After.SelectedIndexChanged += (_, __) => explicitKey.Checked = true;
            }
            page.Controls.Add(pairTable);

            string inference = KeyInferenceSummary(issue.Alignment?.KeyInference, beforeGrid, afterGrid);
            page.Controls.Add(Paragraph(inference, Muted));
            var queue = ActionButton("Queue key decision", 166);
            queue.Dock = DockStyle.Top;
            queue.Click += (_, __) => ApplyKey(issue, auto.Checked, pairs);
            SetHelp(queue, ComparisonAlignmentReviewHelpText.QueueKey);
            page.Controls.Add(queue);
            page.Controls.Add(Details(issue));
            ReplaceEditor(page);
        }

        private void ShowReadOnlyEditor(IssueModel issue, string extra = null)
        {
            var page = EditorPage(issue, "Read-only finding");
            page.Controls.Add(Paragraph(
                string.IsNullOrEmpty(extra)
                    ? "This first correction slice does not offer a safe structured edit for this finding. It remains fully visible and will be preserved in a diagnostic Review."
                    : extra,
                Warning));
            page.Controls.Add(Details(issue));
            ReplaceEditor(page);
        }

        private FlowLayoutPanel EditorPage(IssueModel issue, string title)
        {
            var page = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Surface,
                Padding = new Padding(18, 16, 18, 22),
                Width = Math.Max(520, _editorHost.ClientSize.Width - 22),
            };
            page.Controls.Add(new Label
            {
                Text = title,
                Font = new Font("Segoe UI Semibold", 15F),
                ForeColor = Ink,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 2),
            });
            page.Controls.Add(new Label
            {
                Text = issue.Finding.Code + "  ·  " + ScopeLabel(issue),
                ForeColor = SeverityColor(issue.Finding.Severity),
                AutoSize = true,
                Margin = new Padding(1, 0, 0, 9),
            });
            return page;
        }

        private void ReplaceEditor(Control control)
        {
            _editorHost.SuspendLayout();
            try
            {
                foreach (Control old in _editorHost.Controls.Cast<Control>().ToArray()) old.Dispose();
                _editorHost.Controls.Clear();
                _editorHost.Controls.Add(control);
            }
            finally
            {
                _editorHost.ResumeLayout(true);
            }
        }

        private void ApplySheetPair(IssueModel issue, string afterSheetKey)
        {
            AlignmentSheetPairConfirmation active = ActiveSheetPair(_working, issue);
            if (active != null && string.Equals(active.AfterSheetKey, afterSheetKey, StringComparison.Ordinal))
            {
                Inform("That exact sheet pair is already queued.");
                return;
            }
            TryApplyBatch(new[]
            {
                AlignmentCorrectionRequest.ConfirmSheetPair(
                    issue.Edge.ParentRevisionId,
                    issue.Edge.ChildRevisionId,
                    issue.BeforeSheetKey,
                    afterSheetKey,
                    Reviewer(),
                    DateTimeOffset.UtcNow,
                    AuditNote()),
            });
        }

        private void ApplySheetUnmatched(IssueModel issue)
        {
            AlignmentSheetUnmatchedConfirmation active = ActiveSheetUnmatched(_working, issue);
            if (active != null)
            {
                Inform("This exact sheet is already queued as deleted, with unclaimed targets treated as new.");
                return;
            }
            TryApplyBatch(new[]
            {
                AlignmentCorrectionRequest.ConfirmSheetUnmatched(
                    issue.Edge.ParentRevisionId,
                    issue.Edge.ChildRevisionId,
                    issue.BeforeSheetKey,
                    Reviewer(),
                    DateTimeOffset.UtcNow,
                    AuditNote()),
            });
        }

        private void ClearSheetResolution(IssueModel issue)
        {
            AlignmentSheetPairConfirmation active = ActiveSheetPair(_working, issue);
            AlignmentSheetUnmatchedConfirmation activeUnmatched = ActiveSheetUnmatched(_working, issue);
            if (active == null && activeUnmatched == null)
            {
                Inform("There is no explicit sheet resolution to clear.");
                return;
            }
            TryApplyBatch(new[]
            {
                active != null
                    ? AlignmentCorrectionRequest.ClearSheetPair(
                        issue.Edge.ParentRevisionId,
                        issue.Edge.ChildRevisionId,
                        active.BeforeSheetKey,
                        Reviewer(),
                        DateTimeOffset.UtcNow,
                        AuditNote())
                    : AlignmentCorrectionRequest.ClearSheetUnmatched(
                        issue.Edge.ParentRevisionId,
                        issue.Edge.ChildRevisionId,
                        activeUnmatched.BeforeSheetKey,
                        Reviewer(),
                        DateTimeOffset.UtcNow,
                        AuditNote()),
            });
        }

        private void ApplyHeaders(IssueModel issue, HeaderSideEditor before, HeaderSideEditor after)
        {
            try
            {
                ValidateHeaderPair(before, after);
                var requests = new List<AlignmentCorrectionRequest>();
                AddHeaderRequest(
                    requests,
                    issue.Edge.ParentRevisionId,
                    issue.BeforeSheetKey,
                    before);
                AddHeaderRequest(
                    requests,
                    issue.Edge.ChildRevisionId,
                    issue.AfterSheetKey,
                    after);
                if (requests.Count == 0)
                {
                    Inform("No header confirmation changed.");
                    return;
                }
                TryApplyBatch(requests);
            }
            catch (Exception exception)
            {
                Error("The paired header decision is not valid.", exception);
            }
        }

        private static void ValidateHeaderPair(HeaderSideEditor before, HeaderSideEditor after)
        {
            if (before.Choice != after.Choice)
                throw new InvalidOperationException(
                    "Choose a paired decision: Auto on both sides, No header on both sides, or Header on both " +
                    "sides. Mixing confirmed and automatic/headerless shapes is guaranteed to remain blocked.");
            if (before.Choice != HeaderEditorChoice.Header) return;
            if (before.Grid.Table.RowCount == 0 || after.Grid.Table.RowCount == 0)
                throw new InvalidOperationException(
                    "An empty physical grid cannot carry a confirmed header row. Choose No header on both sides.");

            int beforeHeader = before.HeaderPhysicalRow - before.Grid.FirstPhysicalRow;
            int afterHeader = after.HeaderPhysicalRow - after.Grid.FirstPhysicalRow;
            int beforeData = before.HasFirstDataRow
                ? before.FirstDataPhysicalRow - before.Grid.FirstPhysicalRow
                : beforeHeader + 1;
            int afterData = after.HasFirstDataRow
                ? after.FirstDataPhysicalRow - after.Grid.FirstPhysicalRow
                : afterHeader + 1;
            if (beforeHeader != afterHeader || beforeData != afterData)
                throw new InvalidOperationException(
                    "Confirmed header shapes must have equal table-relative positions on both revisions. " +
                    "Current choices are header/data " +
                    beforeHeader.ToString(CultureInfo.InvariantCulture) + "/" +
                    beforeData.ToString(CultureInfo.InvariantCulture) + " Before and " +
                    afterHeader.ToString(CultureInfo.InvariantCulture) + "/" +
                    afterData.ToString(CultureInfo.InvariantCulture) + " After. Adjust the physical rows so " +
                    "each sheet retains the same number of pre-header and pre-data rows.");
        }

        private void AddHeaderRequest(
            ICollection<AlignmentCorrectionRequest> requests,
            string revisionId,
            string sheetKey,
            HeaderSideEditor editor)
        {
            AlignmentHeaderRowConfirmation active = ActiveHeader(_working, revisionId, sheetKey);
            ReviewerSnapshot reviewer = Reviewer();
            DateTimeOffset now = DateTimeOffset.UtcNow;
            string note = AuditNote();
            switch (editor.Choice)
            {
                case HeaderEditorChoice.Auto:
                    if (active != null)
                        requests.Add(AlignmentCorrectionRequest.ClearHeaderRow(
                            revisionId, sheetKey, reviewer, now, note));
                    return;

                case HeaderEditorChoice.NoHeader:
                    if (active != null && active.Decision == AlignmentHeaderConfirmationDecision.NoHeader) return;
                    requests.Add(AlignmentCorrectionRequest.ConfirmNoHeader(
                        revisionId, sheetKey, reviewer, now, note));
                    return;

                case HeaderEditorChoice.Header:
                    int tableRow = editor.HeaderPhysicalRow - editor.Grid.FirstPhysicalRow;
                    int? firstDataTable = editor.HasFirstDataRow
                        ? editor.FirstDataPhysicalRow - editor.Grid.FirstPhysicalRow
                        : (int?)null;
                    if (active != null && HeaderMatches(
                            active,
                            tableRow,
                            editor.HeaderPhysicalRow,
                            firstDataTable,
                            editor.HasFirstDataRow ? editor.FirstDataPhysicalRow : (int?)null))
                        return;
                    if (firstDataTable.HasValue)
                        requests.Add(AlignmentCorrectionRequest.ConfirmHeaderRow(
                            revisionId,
                            sheetKey,
                            tableRow,
                            editor.HeaderPhysicalRow,
                            firstDataTable.Value,
                            editor.FirstDataPhysicalRow,
                            reviewer,
                            now,
                            note));
                    else
                        requests.Add(AlignmentCorrectionRequest.ConfirmHeaderRow(
                            revisionId,
                            sheetKey,
                            tableRow,
                            editor.HeaderPhysicalRow,
                            reviewer,
                            now,
                            note));
                    return;

                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private void ApplyKey(IssueModel issue, bool auto, IReadOnlyList<KeyPairEditor> pairs)
        {
            try
            {
                AlignmentKeyColumnsConfirmation active = ActiveKey(_working, issue);
                if (auto)
                {
                    if (active == null)
                    {
                        Inform("Automatic key inference is already active.");
                        return;
                    }
                    TryApplyBatch(new[]
                    {
                        AlignmentCorrectionRequest.ClearKeyColumns(
                            issue.Edge.ParentRevisionId,
                            issue.Edge.ChildRevisionId,
                            issue.BeforeSheetKey,
                            issue.AfterSheetKey,
                            Reviewer(),
                            DateTimeOffset.UtcNow,
                            AuditNote()),
                    });
                    return;
                }

                var selected = pairs.Where(pair => pair.HasEitherSelection).ToList();
                if (selected.Count == 0)
                    throw new InvalidOperationException("Choose at least one complete Before → After column pair.");
                if (selected.Any(pair => !pair.HasCompleteSelection))
                    throw new InvalidOperationException("Each used key row needs both a Before and an After column.");
                int last = selected.Max(pair => pair.Order);
                if (pairs.Take(last).Any(pair => !pair.HasCompleteSelection))
                    throw new InvalidOperationException("Key pairs must be contiguous and ordered from row 1.");

                List<AlignmentGridIndexReference> before = selected
                    .OrderBy(pair => pair.Order)
                    .Select(pair => pair.BeforeChoice.Reference)
                    .ToList();
                List<AlignmentGridIndexReference> after = selected
                    .OrderBy(pair => pair.Order)
                    .Select(pair => pair.AfterChoice.Reference)
                    .ToList();
                if (active != null && KeyMatches(active, before, after))
                {
                    Inform("That exact key confirmation is already queued.");
                    return;
                }
                TryApplyBatch(new[]
                {
                    AlignmentCorrectionRequest.ConfirmKeyColumns(
                        issue.Edge.ParentRevisionId,
                        issue.Edge.ChildRevisionId,
                        issue.BeforeSheetKey,
                        issue.AfterSheetKey,
                        before,
                        after,
                        Reviewer(),
                        DateTimeOffset.UtcNow,
                        AuditNote()),
                });
            }
            catch (Exception exception)
            {
                Error("The key-column decision is not valid.", exception);
            }
        }

        private void TryApplyBatch(IEnumerable<AlignmentCorrectionRequest> requests)
        {
            try
            {
                AlignmentCorrectionGeneration next = _working;
                foreach (AlignmentCorrectionRequest request in requests)
                    next = AlignmentCorrectionEngine.Apply(next, request);
                _working = next;
                RefreshIssueStatuses();
                ShowSelectedEditor();
            }
            catch (Exception exception)
            {
                Error("The correction could not be queued safely.", exception);
            }
        }

        private void QueueRecommendationsForSelection()
        {
            List<IssueModel> selected = SelectedIssues();
            if (selected.Count == 0) return;
            AlignmentCorrectionGeneration next = _working;
            int queued = 0;
            int unavailable = 0;
            var failures = new List<string>();
            foreach (IssueModel issue in selected)
            {
                try
                {
                    if (!TryQueueRecommendation(issue, ref next)) unavailable++;
                    else queued++;
                }
                catch (Exception exception)
                {
                    failures.Add(issue.Finding.FindingId + ": " + exception.Message);
                }
            }
            if (queued > 0)
            {
                _working = next;
                RefreshIssueStatuses();
                ShowSelectedEditor();
            }
            string summary = queued.ToString(CultureInfo.InvariantCulture) + " recommendation(s) queued.";
            if (unavailable > 0)
                summary += " " + unavailable.ToString(CultureInfo.InvariantCulture) +
                           " selected finding(s) had no safe recommendation.";
            if (failures.Count > 0)
                summary += "\r\n\r\n" + string.Join("\r\n", failures);
            Inform(summary);
        }

        private bool TryQueueRecommendation(IssueModel issue, ref AlignmentCorrectionGeneration generation)
        {
            ReviewerSnapshot reviewer = Reviewer();
            DateTimeOffset now = DateTimeOffset.UtcNow;
            string note = AuditNote();
            if (issue.EditorKind == IssueEditorKind.SheetPair &&
                issue.UnresolvedSheet != null &&
                issue.UnresolvedSheet.Candidates.Count == 1)
            {
                SheetMappingCandidate candidate = issue.UnresolvedSheet.Candidates[0];
                AlignmentSheetPairConfirmation active = ActiveSheetPair(generation, issue);
                if (active != null && string.Equals(
                        active.AfterSheetKey,
                        candidate.After.SheetKey,
                        StringComparison.Ordinal))
                    return false;
                generation = AlignmentCorrectionEngine.Apply(
                    generation,
                    AlignmentCorrectionRequest.ConfirmSheetPair(
                        issue.Edge.ParentRevisionId,
                        issue.Edge.ChildRevisionId,
                        issue.BeforeSheetKey,
                        candidate.After.SheetKey,
                        reviewer,
                        now,
                        note));
                return true;
            }

            if (issue.EditorKind == IssueEditorKind.Header &&
                TryHeaderRecommendation(issue, out HeaderRecommendation before, out HeaderRecommendation after))
            {
                var requests = new List<AlignmentCorrectionRequest>();
                AddRecommendedHeaderRequest(requests, generation, before, reviewer, now, note);
                AddRecommendedHeaderRequest(requests, generation, after, reviewer, now, note);
                if (requests.Count == 0) return false;
                AlignmentCorrectionGeneration paired = generation;
                foreach (AlignmentCorrectionRequest request in requests)
                    paired = AlignmentCorrectionEngine.Apply(paired, request);
                generation = paired;
                return true;
            }

            if (issue.EditorKind == IssueEditorKind.Key && issue.Alignment != null)
            {
                KeyCandidate candidate = SuggestedKey(issue.Alignment.KeyInference);
                if (candidate == null || candidate.HasDuplicates) return false;
                if (!TryGetAlignmentEvidence(issue, out OpenXmlSheetFingerprintEvidence beforeEvidence,
                        out OpenXmlSheetFingerprintEvidence afterEvidence))
                    return false;
                CanonicalGridProjection beforeGrid = Grid(issue.Edge.ParentRevisionId, beforeEvidence);
                CanonicalGridProjection afterGrid = Grid(issue.Edge.ChildRevisionId, afterEvidence);
                var beforeColumns = candidate.BeforeColumns.Select(index => new AlignmentGridIndexReference(
                    index, beforeGrid.PhysicalColumnForTableIndex(index))).ToList();
                var afterColumns = candidate.AfterColumns.Select(index => new AlignmentGridIndexReference(
                    index, afterGrid.PhysicalColumnForTableIndex(index))).ToList();
                AlignmentKeyColumnsConfirmation active = ActiveKey(generation, issue);
                if (active != null && KeyMatches(active, beforeColumns, afterColumns)) return false;
                generation = AlignmentCorrectionEngine.Apply(
                    generation,
                    AlignmentCorrectionRequest.ConfirmKeyColumns(
                        issue.Edge.ParentRevisionId,
                        issue.Edge.ChildRevisionId,
                        issue.BeforeSheetKey,
                        issue.AfterSheetKey,
                        beforeColumns,
                        afterColumns,
                        reviewer,
                        now,
                        note));
                return true;
            }
            return false;
        }

        private static void AddRecommendedHeaderRequest(
            ICollection<AlignmentCorrectionRequest> requests,
            AlignmentCorrectionGeneration generation,
            HeaderRecommendation recommendation,
            ReviewerSnapshot reviewer,
            DateTimeOffset now,
            string note)
        {
            AlignmentHeaderRowConfirmation active = ActiveHeader(
                generation,
                recommendation.RevisionId,
                recommendation.SheetKey);
            if (active != null && HeaderMatches(
                    active,
                    recommendation.HeaderTableRow,
                    recommendation.HeaderPhysicalRow,
                    recommendation.FirstDataTableRow,
                    recommendation.FirstDataPhysicalRow))
                return;
            if (recommendation.FirstDataTableRow.HasValue)
                requests.Add(AlignmentCorrectionRequest.ConfirmHeaderRow(
                    recommendation.RevisionId,
                    recommendation.SheetKey,
                    recommendation.HeaderTableRow,
                    recommendation.HeaderPhysicalRow,
                    recommendation.FirstDataTableRow.Value,
                    recommendation.FirstDataPhysicalRow.Value,
                    reviewer,
                    now,
                    note));
            else
                requests.Add(AlignmentCorrectionRequest.ConfirmHeaderRow(
                    recommendation.RevisionId,
                    recommendation.SheetKey,
                    recommendation.HeaderTableRow,
                    recommendation.HeaderPhysicalRow,
                    reviewer,
                    now,
                    note));
        }

        private bool TryHeaderRecommendation(
            IssueModel issue,
            out HeaderRecommendation before,
            out HeaderRecommendation after)
        {
            before = null;
            after = null;
            if (issue.Alignment == null || issue.Edge == null ||
                !TryGetAlignmentEvidence(issue, out OpenXmlSheetFingerprintEvidence beforeEvidence,
                    out OpenXmlSheetFingerprintEvidence afterEvidence))
                return false;
            CanonicalGridProjection beforeGrid = Grid(issue.Edge.ParentRevisionId, beforeEvidence);
            CanonicalGridProjection afterGrid = Grid(issue.Edge.ChildRevisionId, afterEvidence);
            int? beforeRow = PreferredHeaderRow(issue.Alignment.BeforeHeader);
            int? afterRow = PreferredHeaderRow(issue.Alignment.AfterHeader);
            if (!beforeRow.HasValue || !afterRow.HasValue || beforeRow.Value != afterRow.Value) return false;
            int beforeData = PreferredDataStart(issue.Alignment.BeforeHeader, beforeRow.Value);
            int afterData = PreferredDataStart(issue.Alignment.AfterHeader, afterRow.Value);
            if (beforeData != afterData ||
                beforeRow.Value < 0 || beforeRow.Value >= beforeGrid.Table.RowCount ||
                afterRow.Value < 0 || afterRow.Value >= afterGrid.Table.RowCount ||
                beforeData > beforeGrid.Table.RowCount || afterData > afterGrid.Table.RowCount)
                return false;
            int? dataTable = beforeData == beforeRow.Value + 1 ? (int?)null : beforeData;
            before = new HeaderRecommendation(
                issue.Edge.ParentRevisionId,
                issue.BeforeSheetKey,
                beforeRow.Value,
                beforeGrid.FirstPhysicalRow + beforeRow.Value,
                dataTable,
                dataTable.HasValue ? beforeGrid.FirstPhysicalRow + dataTable.Value : (int?)null);
            after = new HeaderRecommendation(
                issue.Edge.ChildRevisionId,
                issue.AfterSheetKey,
                afterRow.Value,
                afterGrid.FirstPhysicalRow + afterRow.Value,
                dataTable,
                dataTable.HasValue ? afterGrid.FirstPhysicalRow + dataTable.Value : (int?)null);
            return true;
        }

        private void ConfirmDiagnostic()
        {
            DialogResult answer = ExcelPowerTools.Interop.Dialogs.Message(
                "Build a diagnostic Review.xlsx with the current findings?\r\n\r\n" +
                "This is an explicit incomplete artifact. It preserves evidence and queued correction audit events, " +
                "but it does not claim that logical history is complete.",
                MessageBoxButtons.YesNo, "Build diagnostic Review", this, defaultToNegative: true);
            if (answer == DialogResult.Yes) Finish(ComparisonAlignmentReviewAction.PublishDiagnostic);
        }

        private void Finish(ComparisonAlignmentReviewAction action)
        {
            AlignmentCorrectionGeneration resultCorrections = action == ComparisonAlignmentReviewAction.Cancel
                ? _analysis.AlignmentCorrections
                : _working;
            Result = new ComparisonAlignmentReviewResult(action, resultCorrections);
            _closingWithResult = true;
            DialogResult = action == ComparisonAlignmentReviewAction.Cancel
                ? DialogResult.Cancel
                : DialogResult.OK;
            Close();
        }

        private ReviewerSnapshot Reviewer()
        {
            string name = _reviewer.Text.Trim();
            return name.Length == 0
                ? ReviewerSnapshot.Empty
                : new ReviewerSnapshot(name, "", new[] { "comparison-alignment-review" });
        }

        private string AuditNote() => _note.Text.Trim();

        private void RefreshWorkingStatus()
        {
            int eventsAdded = _working.Events.Count - _analysis.AlignmentCorrections.Events.Count;
            int activeCount = _working.SheetPairs.Count + _working.UnmatchedSheets.Count +
                              _working.HeaderRows.Count + _working.KeyColumns.Count;
            bool dirty = !string.Equals(
                _working.GenerationId,
                _analysis.AlignmentCorrections.GenerationId,
                StringComparison.Ordinal);
            _workingStatus.Text = dirty
                ? eventsAdded.ToString(CultureInfo.InvariantCulture) + " new audit event(s) · " +
                  activeCount.ToString(CultureInfo.InvariantCulture) + " active confirmation(s) · rerun required"
                : activeCount > 0
                    ? activeCount.ToString(CultureInfo.InvariantCulture) +
                      " active confirmation(s) · no new edits"
                    : "Automatic inference · no queued changes";
            _workingStatus.ForeColor = dirty ? Accent : Muted;
            _reanalyze.Enabled = dirty;
        }

        private string Status(IssueModel issue)
        {
            switch (issue.EditorKind)
            {
                case IssueEditorKind.SheetPair:
                    if (ActiveSheetUnmatched(_working, issue) != null) return "Queued deleted/new";
                    return ActiveSheetPair(_working, issue) == null ? "Needs review" : "Queued pair";
                case IssueEditorKind.Header:
                    if (issue.Alignment == null || issue.Edge == null) return "Read-only";
                    int headers = 0;
                    if (ActiveHeader(_working, issue.Edge.ParentRevisionId, issue.BeforeSheetKey) != null)
                        headers++;
                    if (ActiveHeader(_working, issue.Edge.ChildRevisionId, issue.AfterSheetKey) != null)
                        headers++;
                    return headers == 0 ? "Needs review" : "Queued " + headers + "/2";
                case IssueEditorKind.SingleHeader:
                    return ActiveHeader(_working, issue.SingleRevisionId, issue.SingleSheetKey) == null
                        ? "Needs review"
                        : "Queued header";
                case IssueEditorKind.Key:
                    return ActiveKey(_working, issue) == null ? "Needs review" : "Queued key";
                default:
                    return "Read-only";
            }
        }

        private static string FindingLabel(IssueModel issue)
        {
            switch (issue.EditorKind)
            {
                case IssueEditorKind.SheetPair: return "Sheet mapping · " + issue.Finding.SheetName;
                case IssueEditorKind.Header: return "Header · " + issue.Finding.SheetName;
                case IssueEditorKind.SingleHeader: return "Header · " + issue.Finding.SheetName;
                case IssueEditorKind.Key: return "Row key · " + issue.Finding.SheetName;
                default: return issue.Finding.Code +
                                (string.IsNullOrEmpty(issue.Finding.SheetName)
                                    ? ""
                                    : " · " + issue.Finding.SheetName);
            }
        }

        private List<IssueModel> SelectedIssues() => _issueList.SelectedItems
            .Cast<ListViewItem>()
            .Select(item => item.Tag as IssueModel)
            .Where(item => item != null)
            .OrderBy(item => item.SortOrder)
            .ToList();

        private static List<IssueModel> BuildIssues(ComparisonProjectAnalysisResult analysis)
        {
            OpenXmlComparisonProjectResult result = analysis.ProjectResult;
            AlignmentCorrectionGeneration corrections = analysis.AlignmentCorrections;
            var issues = new List<IssueModel>(
                result.Findings.Count + corrections.SheetPairs.Count + corrections.UnmatchedSheets.Count +
                corrections.HeaderRows.Count + corrections.KeyColumns.Count);
            int order = 0;
            foreach (OpenXmlProjectFinding finding in result.Findings)
            {
                OpenXmlProjectEdgeResult edge = result.Edges.SingleOrDefault(item =>
                    string.Equals(item.ParentRevisionId, finding.BeforeRevisionId, StringComparison.Ordinal) &&
                    string.Equals(item.ChildRevisionId, finding.AfterRevisionId, StringComparison.Ordinal));
                OpenXmlProjectSheetAlignment alignment = ResolveAlignment(edge, finding);
                UnresolvedSheetMapping unresolved = ResolveUnresolved(result, edge, finding);
                IssueEditorKind kind = Route(finding, edge, alignment, unresolved);
                string before = RevisionLabel(analysis, finding.BeforeRevisionId);
                string after = RevisionLabel(analysis, finding.AfterRevisionId);
                string edgeLabel = edge == null
                    ? "Project / coverage"
                    : before + "  →  " + after;
                string sheet = string.IsNullOrEmpty(finding.SheetName) ? "Workbook" : finding.SheetName;
                string groupKey = (edge == null ? "project" : edge.ParentRevisionId + "\u001f" + edge.ChildRevisionId) +
                                  "\u001f" + finding.LogicalSheetId + "\u001f" + sheet;
                issues.Add(new IssueModel(
                    finding,
                    edge,
                    alignment,
                    unresolved,
                    kind,
                    groupKey,
                    edgeLabel + "  ·  " + sheet,
                    order++));
            }

            // A successful correction normally removes its original diagnostic. Keep the exact active scope in
            // this command-lifetime editor so the user can inspect, replace, or clear it without reconstructing
            // identity from stale finding prose.
            foreach (AlignmentSheetPairConfirmation confirmation in corrections.SheetPairs)
            {
                OpenXmlProjectEdgeResult edge = FindEdge(
                    result,
                    confirmation.ParentRevisionId,
                    confirmation.ChildRevisionId);
                if (edge == null || issues.Any(issue =>
                        issue.EditorKind == IssueEditorKind.SheetPair &&
                        SameEdge(issue.Edge, confirmation.ParentRevisionId, confirmation.ChildRevisionId) &&
                        string.Equals(issue.BeforeSheetKey, confirmation.BeforeSheetKey, StringComparison.Ordinal)))
                    continue;
                WorkbookSheetFingerprint before = edge.SheetMapping.Before.Fingerprints.SingleOrDefault(item =>
                    string.Equals(item.SheetKey, confirmation.BeforeSheetKey, StringComparison.Ordinal));
                WorkbookSheetFingerprint after = edge.SheetMapping.After.Fingerprints.SingleOrDefault(item =>
                    string.Equals(item.SheetKey, confirmation.AfterSheetKey, StringComparison.Ordinal));
                if (before == null || after == null) continue;
                OpenXmlProjectSheetAlignment alignment = ResolveAlignment(
                    edge,
                    confirmation.BeforeSheetKey,
                    confirmation.AfterSheetKey);
                OpenXmlProjectFinding finding = SyntheticFinding(
                    "CMP-UI-ACTIVE-SHEET-CONFIRMATION",
                    confirmation.ConfirmationId,
                    confirmation.ParentRevisionId,
                    confirmation.ChildRevisionId,
                    alignment?.LogicalSheetId ?? "",
                    before.Name,
                    "An explicit sheet-pair confirmation is active: '" + before.Name + "' → '" + after.Name + "'.",
                    "Keep, replace, or clear this exact confirmation before continuing.");
                issues.Add(SyntheticIssue(
                    analysis,
                    finding,
                    edge,
                    alignment,
                    IssueEditorKind.SheetPair,
                    order++,
                    confirmation.BeforeSheetKey,
                    confirmation.AfterSheetKey));
            }

            foreach (AlignmentSheetUnmatchedConfirmation confirmation in corrections.UnmatchedSheets)
            {
                OpenXmlProjectEdgeResult edge = FindEdge(
                    result,
                    confirmation.ParentRevisionId,
                    confirmation.ChildRevisionId);
                if (edge == null || issues.Any(issue =>
                        issue.EditorKind == IssueEditorKind.SheetPair &&
                        SameEdge(issue.Edge, confirmation.ParentRevisionId, confirmation.ChildRevisionId) &&
                        string.Equals(issue.BeforeSheetKey, confirmation.BeforeSheetKey, StringComparison.Ordinal)))
                    continue;
                WorkbookSheetFingerprint before = edge.SheetMapping.Before.Fingerprints.SingleOrDefault(item =>
                    string.Equals(item.SheetKey, confirmation.BeforeSheetKey, StringComparison.Ordinal));
                if (before == null) continue;
                OpenXmlProjectFinding finding = SyntheticFinding(
                    "CMP-UI-ACTIVE-SHEET-UNMATCHED-CONFIRMATION",
                    confirmation.ConfirmationId,
                    confirmation.ParentRevisionId,
                    confirmation.ChildRevisionId,
                    "",
                    before.Name,
                    "An explicit deleted/new confirmation is active for source sheet '" + before.Name + "'.",
                    "Keep it, replace it with an exact pair, or clear it to return this source to automatic mapping.");
                issues.Add(SyntheticIssue(
                    analysis,
                    finding,
                    edge,
                    null,
                    IssueEditorKind.SheetPair,
                    order++,
                    confirmation.BeforeSheetKey,
                    ""));
            }

            var representedHeaders = new HashSet<string>(StringComparer.Ordinal);
            foreach (IssueModel issue in issues.Where(item => item.EditorKind == IssueEditorKind.Header))
            {
                representedHeaders.Add(HeaderScope(issue.Edge.ParentRevisionId, issue.BeforeSheetKey));
                representedHeaders.Add(HeaderScope(issue.Edge.ChildRevisionId, issue.AfterSheetKey));
            }
            foreach (AlignmentHeaderRowConfirmation confirmation in corrections.HeaderRows)
            {
                string scope = HeaderScope(confirmation.RevisionId, confirmation.PhysicalSheetKey);
                if (representedHeaders.Contains(scope)) continue;

                OpenXmlProjectEdgeResult chosenEdge = null;
                OpenXmlProjectSheetAlignment chosenAlignment = null;
                foreach (OpenXmlProjectEdgeResult edge in result.Edges
                             .OrderBy(item => item.ChildRevisionId, StringComparer.Ordinal))
                {
                    foreach (OpenXmlProjectSheetAlignment alignment in edge.SheetAlignments
                                 .OrderBy(item => item.BeforeSheetKey, StringComparer.Ordinal))
                    {
                        bool matches =
                            (string.Equals(edge.ParentRevisionId, confirmation.RevisionId, StringComparison.Ordinal) &&
                             string.Equals(alignment.BeforeSheetKey, confirmation.PhysicalSheetKey, StringComparison.Ordinal)) ||
                            (string.Equals(edge.ChildRevisionId, confirmation.RevisionId, StringComparison.Ordinal) &&
                             string.Equals(alignment.AfterSheetKey, confirmation.PhysicalSheetKey, StringComparison.Ordinal));
                        if (!matches) continue;
                        chosenEdge = edge;
                        chosenAlignment = alignment;
                        bool counterpartActive = corrections.TryGetHeaderRow(
                            edge.ParentRevisionId,
                            alignment.BeforeSheetKey,
                            out _) && corrections.TryGetHeaderRow(
                            edge.ChildRevisionId,
                            alignment.AfterSheetKey,
                            out _);
                        if (counterpartActive) break;
                    }
                    if (chosenAlignment != null &&
                        corrections.TryGetHeaderRow(
                            chosenEdge.ParentRevisionId,
                            chosenAlignment.BeforeSheetKey,
                            out _) && corrections.TryGetHeaderRow(
                            chosenEdge.ChildRevisionId,
                            chosenAlignment.AfterSheetKey,
                            out _))
                        break;
                }

                if (chosenEdge != null && chosenAlignment != null)
                {
                    corrections.TryGetHeaderRow(
                        chosenEdge.ParentRevisionId,
                        chosenAlignment.BeforeSheetKey,
                        out AlignmentHeaderRowConfirmation beforeActive);
                    corrections.TryGetHeaderRow(
                        chosenEdge.ChildRevisionId,
                        chosenAlignment.AfterSheetKey,
                        out AlignmentHeaderRowConfirmation afterActive);
                    string id = beforeActive?.ConfirmationId ?? afterActive?.ConfirmationId ?? confirmation.ConfirmationId;
                    OpenXmlProjectFinding finding = SyntheticFinding(
                        "CMP-UI-ACTIVE-HEADER-CONFIRMATION",
                        id,
                        chosenEdge.ParentRevisionId,
                        chosenEdge.ChildRevisionId,
                        chosenAlignment.LogicalSheetId,
                        chosenAlignment.AfterSheetName,
                        "One or both mapped sheet occurrences have an active explicit header decision.",
                        "Keep paired shapes consistent, replace either exact decision, or choose Auto on both sides to clear them.");
                    issues.Add(SyntheticIssue(
                        analysis,
                        finding,
                        chosenEdge,
                        chosenAlignment,
                        IssueEditorKind.Header,
                        order++,
                        chosenAlignment.BeforeSheetKey,
                        chosenAlignment.AfterSheetKey));
                    if (beforeActive != null)
                        representedHeaders.Add(HeaderScope(
                            chosenEdge.ParentRevisionId,
                            chosenAlignment.BeforeSheetKey));
                    if (afterActive != null)
                        representedHeaders.Add(HeaderScope(
                            chosenEdge.ChildRevisionId,
                            chosenAlignment.AfterSheetKey));
                    continue;
                }

                OpenXmlSheetFingerprintEvidence evidence = FindEvidence(
                    result,
                    confirmation.RevisionId,
                    confirmation.PhysicalSheetKey);
                string sheetName = evidence?.Fingerprint.Name ?? "Sheet occurrence";
                OpenXmlProjectFinding singleFinding = SyntheticFinding(
                    "CMP-UI-ACTIVE-HEADER-CONFIRMATION",
                    confirmation.ConfirmationId,
                    confirmation.RevisionId,
                    "",
                    "",
                    sheetName,
                    "An explicit header decision is active for this exact sheet occurrence.",
                    "Keep, replace, or clear this exact decision before continuing.");
                issues.Add(SyntheticIssue(
                    analysis,
                    singleFinding,
                    null,
                    null,
                    IssueEditorKind.SingleHeader,
                    order++,
                    "",
                    "",
                    confirmation.RevisionId,
                    confirmation.PhysicalSheetKey,
                    FindHeaderDetection(result, confirmation.RevisionId, confirmation.PhysicalSheetKey)));
                representedHeaders.Add(scope);
            }

            foreach (AlignmentKeyColumnsConfirmation confirmation in corrections.KeyColumns)
            {
                OpenXmlProjectEdgeResult edge = FindEdge(
                    result,
                    confirmation.ParentRevisionId,
                    confirmation.ChildRevisionId);
                if (edge == null || issues.Any(issue =>
                        issue.EditorKind == IssueEditorKind.Key &&
                        SameEdge(issue.Edge, confirmation.ParentRevisionId, confirmation.ChildRevisionId) &&
                        string.Equals(issue.BeforeSheetKey, confirmation.BeforeSheetKey, StringComparison.Ordinal) &&
                        string.Equals(issue.AfterSheetKey, confirmation.AfterSheetKey, StringComparison.Ordinal)))
                    continue;
                OpenXmlProjectSheetAlignment alignment = ResolveAlignment(
                    edge,
                    confirmation.BeforeSheetKey,
                    confirmation.AfterSheetKey);
                WorkbookSheetFingerprint after = edge.SheetMapping.After.Fingerprints.SingleOrDefault(item =>
                    string.Equals(item.SheetKey, confirmation.AfterSheetKey, StringComparison.Ordinal));
                if (after == null) continue;
                OpenXmlProjectFinding finding = SyntheticFinding(
                    "CMP-UI-ACTIVE-KEY-CONFIRMATION",
                    confirmation.ConfirmationId,
                    confirmation.ParentRevisionId,
                    confirmation.ChildRevisionId,
                    alignment?.LogicalSheetId ?? "",
                    after.Name,
                    "An explicit ordered key-column confirmation is active for this mapped sheet pair.",
                    "Keep, replace, or choose Auto to clear this exact key decision.");
                issues.Add(SyntheticIssue(
                    analysis,
                    finding,
                    edge,
                    alignment,
                    IssueEditorKind.Key,
                    order++,
                    confirmation.BeforeSheetKey,
                    confirmation.AfterSheetKey));
            }

            return issues
                .OrderBy(issue => issue.Edge == null ? 1 : 0)
                .ThenBy(issue => issue.Edge?.ChildRevisionId ?? "", StringComparer.Ordinal)
                .ThenBy(issue => issue.Finding.SheetName, StringComparer.OrdinalIgnoreCase)
                .ThenByDescending(issue => issue.Finding.Severity)
                .ThenBy(issue => issue.Finding.Code, StringComparer.Ordinal)
                .ThenBy(issue => issue.SortOrder)
                .ToList();
        }

        private static IssueModel SyntheticIssue(
            ComparisonProjectAnalysisResult analysis,
            OpenXmlProjectFinding finding,
            OpenXmlProjectEdgeResult edge,
            OpenXmlProjectSheetAlignment alignment,
            IssueEditorKind kind,
            int order,
            string beforeSheetKey,
            string afterSheetKey,
            string singleRevisionId = "",
            string singleSheetKey = "",
            HeaderDetectionResult singleHeaderDetection = null)
        {
            string edgeLabel = edge == null
                ? RevisionLabel(analysis, singleRevisionId)
                : RevisionLabel(analysis, edge.ParentRevisionId) + "  →  " +
                  RevisionLabel(analysis, edge.ChildRevisionId);
            string sheet = string.IsNullOrEmpty(finding.SheetName) ? "Sheet occurrence" : finding.SheetName;
            string scopeKey = edge == null
                ? "active\u001f" + singleRevisionId + "\u001f" + singleSheetKey
                : edge.ParentRevisionId + "\u001f" + edge.ChildRevisionId + "\u001f" +
                  beforeSheetKey + "\u001f" + afterSheetKey;
            return new IssueModel(
                finding,
                edge,
                alignment,
                null,
                kind,
                scopeKey,
                edgeLabel + "  ·  " + sheet + "  ·  Active confirmation",
                order,
                beforeSheetKey,
                afterSheetKey,
                singleRevisionId,
                singleSheetKey,
                true,
                singleHeaderDetection);
        }

        private static OpenXmlProjectFinding SyntheticFinding(
            string code,
            string confirmationId,
            string beforeRevisionId,
            string afterRevisionId,
            string logicalSheetId,
            string sheetName,
            string message,
            string suggestedAction) => new OpenXmlProjectFinding(
            "active-confirmation:" + confirmationId,
            code,
            DiagnosticSeverity.Information,
            beforeRevisionId,
            afterRevisionId,
            logicalSheetId,
            sheetName,
            message,
            suggestedAction,
            new[] { "alignment.confirmation.active" },
            Array.Empty<string>());

        private static OpenXmlProjectEdgeResult FindEdge(
            OpenXmlComparisonProjectResult project,
            string parentRevisionId,
            string childRevisionId) => project.Edges.SingleOrDefault(item =>
            string.Equals(item.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
            string.Equals(item.ChildRevisionId, childRevisionId, StringComparison.Ordinal));

        private static bool SameEdge(
            OpenXmlProjectEdgeResult edge,
            string parentRevisionId,
            string childRevisionId) => edge != null &&
            string.Equals(edge.ParentRevisionId, parentRevisionId, StringComparison.Ordinal) &&
            string.Equals(edge.ChildRevisionId, childRevisionId, StringComparison.Ordinal);

        private static string HeaderScope(string revisionId, string sheetKey) =>
            (revisionId ?? "") + "\u001f" + (sheetKey ?? "");

        private static OpenXmlProjectSheetAlignment ResolveAlignment(
            OpenXmlProjectEdgeResult edge,
            string beforeSheetKey,
            string afterSheetKey) => edge?.SheetAlignments.SingleOrDefault(item =>
            string.Equals(item.BeforeSheetKey, beforeSheetKey, StringComparison.Ordinal) &&
            string.Equals(item.AfterSheetKey, afterSheetKey, StringComparison.Ordinal));

        private static OpenXmlSheetFingerprintEvidence FindEvidence(
            OpenXmlComparisonProjectResult project,
            string revisionId,
            string sheetKey)
        {
            foreach (OpenXmlProjectEdgeResult edge in project.Edges)
            {
                IEnumerable<OpenXmlSheetFingerprintEvidence> candidates = Array.Empty<OpenXmlSheetFingerprintEvidence>();
                if (string.Equals(edge.ParentRevisionId, revisionId, StringComparison.Ordinal))
                    candidates = edge.SheetMapping.Before.Sheets;
                else if (string.Equals(edge.ChildRevisionId, revisionId, StringComparison.Ordinal))
                    candidates = edge.SheetMapping.After.Sheets;
                OpenXmlSheetFingerprintEvidence exact = candidates.FirstOrDefault(item =>
                    string.Equals(item.Fingerprint.SheetKey, sheetKey, StringComparison.Ordinal));
                if (exact != null) return exact;
            }
            return null;
        }

        private static HeaderDetectionResult FindHeaderDetection(
            OpenXmlComparisonProjectResult project,
            string revisionId,
            string sheetKey)
        {
            foreach (OpenXmlProjectEdgeResult edge in project.Edges)
            foreach (OpenXmlProjectSheetAlignment alignment in edge.SheetAlignments)
            {
                if (string.Equals(edge.ParentRevisionId, revisionId, StringComparison.Ordinal) &&
                    string.Equals(alignment.BeforeSheetKey, sheetKey, StringComparison.Ordinal))
                    return alignment.BeforeHeader;
                if (string.Equals(edge.ChildRevisionId, revisionId, StringComparison.Ordinal) &&
                    string.Equals(alignment.AfterSheetKey, sheetKey, StringComparison.Ordinal))
                    return alignment.AfterHeader;
            }
            return null;
        }

        private static IssueEditorKind Route(
            OpenXmlProjectFinding finding,
            OpenXmlProjectEdgeResult edge,
            OpenXmlProjectSheetAlignment alignment,
            UnresolvedSheetMapping unresolved)
        {
            if (finding.Code == OpenXmlProjectFindingCodes.SheetMappingAmbiguous &&
                edge != null && unresolved != null)
                return IssueEditorKind.SheetPair;
            if ((finding.Code == OpenXmlProjectFindingCodes.HeaderMismatch ||
                 finding.Code == OpenXmlProjectFindingCodes.NonLeadingHeader) &&
                edge != null && alignment != null)
                return IssueEditorKind.Header;
            if ((finding.Code == OpenXmlProjectFindingCodes.RowAlignmentAmbiguous ||
                 finding.Code == OpenXmlProjectFindingCodes.WeakRowAlignment) &&
                edge != null && alignment != null && alignment.KeyInference != null)
                return IssueEditorKind.Key;
            return IssueEditorKind.ReadOnly;
        }

        private static OpenXmlProjectSheetAlignment ResolveAlignment(
            OpenXmlProjectEdgeResult edge,
            OpenXmlProjectFinding finding)
        {
            if (edge == null) return null;
            IEnumerable<OpenXmlProjectSheetAlignment> candidates = edge.SheetAlignments;
            if (!string.IsNullOrEmpty(finding.LogicalSheetId))
                candidates = candidates.Where(item =>
                    string.Equals(item.LogicalSheetId, finding.LogicalSheetId, StringComparison.Ordinal));
            if (!string.IsNullOrEmpty(finding.SheetName))
                candidates = candidates.Where(item =>
                    string.Equals(item.AfterSheetName, finding.SheetName, StringComparison.Ordinal) ||
                    string.Equals(item.BeforeSheetName, finding.SheetName, StringComparison.Ordinal));
            List<OpenXmlProjectSheetAlignment> exact = candidates.ToList();
            return exact.Count == 1 ? exact[0] : null;
        }

        private static UnresolvedSheetMapping ResolveUnresolved(
            OpenXmlComparisonProjectResult project,
            OpenXmlProjectEdgeResult edge,
            OpenXmlProjectFinding finding)
        {
            if (edge == null || finding.Code != OpenXmlProjectFindingCodes.SheetMappingAmbiguous) return null;
            string beforeKey = null;
            OpenXmlAlignedProjectRevision parent = project.Revisions.SingleOrDefault(item =>
                string.Equals(item.Source.RevisionId, edge.ParentRevisionId, StringComparison.Ordinal));
            if (parent != null && !string.IsNullOrEmpty(finding.LogicalSheetId))
            {
                List<OpenXmlAlignedSheetRevision> sheets = parent.Sheets.Where(item =>
                    string.Equals(item.LogicalSheetId, finding.LogicalSheetId, StringComparison.Ordinal) &&
                    string.Equals(item.PhysicalSheetName, finding.SheetName, StringComparison.Ordinal)).ToList();
                if (sheets.Count == 1) beforeKey = sheets[0].PhysicalSheetKey;
            }
            IEnumerable<UnresolvedSheetMapping> unresolved = edge.SheetMapping.Mapping.UnresolvedBeforeSheets;
            if (beforeKey != null)
                unresolved = unresolved.Where(item =>
                    string.Equals(item.Before.SheetKey, beforeKey, StringComparison.Ordinal));
            else
                unresolved = unresolved.Where(item =>
                    string.Equals(item.Before.Name, finding.SheetName, StringComparison.Ordinal));
            List<UnresolvedSheetMapping> exact = unresolved.ToList();
            return exact.Count == 1 ? exact[0] : null;
        }

        private bool TryGetAlignmentEvidence(
            IssueModel issue,
            out OpenXmlSheetFingerprintEvidence before,
            out OpenXmlSheetFingerprintEvidence after)
        {
            before = null;
            after = null;
            if (issue.Edge == null || string.IsNullOrEmpty(issue.BeforeSheetKey) ||
                string.IsNullOrEmpty(issue.AfterSheetKey)) return false;
            before = issue.Edge.SheetMapping.Before.Sheets.SingleOrDefault(item =>
                string.Equals(item.Fingerprint.SheetKey, issue.BeforeSheetKey, StringComparison.Ordinal));
            after = issue.Edge.SheetMapping.After.Sheets.SingleOrDefault(item =>
                string.Equals(item.Fingerprint.SheetKey, issue.AfterSheetKey, StringComparison.Ordinal));
            return before != null && after != null;
        }

        private OpenXmlSheetFingerprintEvidence FindEvidence(string revisionId, string sheetKey)
        {
            foreach (OpenXmlProjectEdgeResult edge in _analysis.ProjectResult.Edges)
            {
                IEnumerable<OpenXmlSheetFingerprintEvidence> candidates = Array.Empty<OpenXmlSheetFingerprintEvidence>();
                if (string.Equals(edge.ParentRevisionId, revisionId, StringComparison.Ordinal))
                    candidates = edge.SheetMapping.Before.Sheets;
                else if (string.Equals(edge.ChildRevisionId, revisionId, StringComparison.Ordinal))
                    candidates = edge.SheetMapping.After.Sheets;
                OpenXmlSheetFingerprintEvidence exact = candidates.FirstOrDefault(item =>
                    string.Equals(item.Fingerprint.SheetKey, sheetKey, StringComparison.Ordinal));
                if (exact != null) return exact;
            }
            return null;
        }

        private CanonicalGridProjection Grid(string revisionId, OpenXmlSheetFingerprintEvidence evidence)
        {
            string key = revisionId.Length.ToString(CultureInfo.InvariantCulture) + ":" + revisionId +
                         evidence.Fingerprint.SheetKey;
            if (!_grids.TryGetValue(key, out CanonicalGridProjection grid))
            {
                grid = CanonicalGridProjectionBuilder.Build(evidence.Source.Snapshot);
                _grids.Add(key, grid);
            }
            return grid;
        }

        private IReadOnlyList<TargetChoice> SheetTargets(IssueModel issue)
        {
            var claimed = new HashSet<string>(
                _working.GetSheetPairs(issue.Edge.ParentRevisionId, issue.Edge.ChildRevisionId)
                    .Where(item => !string.Equals(
                        item.BeforeSheetKey,
                        issue.BeforeSheetKey,
                        StringComparison.Ordinal))
                    .Select(item => item.AfterSheetKey),
                StringComparer.Ordinal);
            var candidates = (issue.UnresolvedSheet?.Candidates ?? Array.Empty<SheetMappingCandidate>())
                .ToDictionary(item => item.After.SheetKey, StringComparer.Ordinal);
            var inferredByAfter = issue.Edge.SheetMapping.Mapping.Mappings.ToDictionary(
                item => item.After.SheetKey,
                item => item.Before.Name,
                StringComparer.Ordinal);
            var result = new List<TargetChoice>();
            foreach (OpenXmlSheetFingerprintEvidence evidence in issue.Edge.SheetMapping.After.Sheets)
            {
                WorkbookSheetFingerprint sheet = evidence.Fingerprint;
                if (sheet.Kind != WorkbookSheetKind.Worksheet || claimed.Contains(sheet.SheetKey)) continue;
                candidates.TryGetValue(sheet.SheetKey, out SheetMappingCandidate candidate);
                inferredByAfter.TryGetValue(sheet.SheetKey, out string inferredSource);
                result.Add(new TargetChoice(sheet, candidate, inferredSource));
            }
            return new ReadOnlyCollection<TargetChoice>(result
                .OrderByDescending(item => item.Candidate != null)
                .ThenByDescending(item => item.Candidate?.Confidence ?? -1D)
                .ThenBy(item => item.Fingerprint.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Fingerprint.SheetKey, StringComparer.Ordinal)
                .ToList());
        }

        private static IReadOnlyList<ColumnChoice> ColumnChoices(
            CanonicalGridProjection grid,
            HeaderDetectionResult header)
        {
            if (grid.Table.ColumnCount == 0) return Array.Empty<ColumnChoice>();
            int? labelRow = header?.AcceptedHeaderRowIndex ?? header?.CandidateRowIndex;
            int dataStart = header?.HasAcceptedHeader == true
                ? header.DataStartRowIndex
                : labelRow.HasValue ? labelRow.Value + 1 : 0;
            dataStart = Math.Max(0, Math.Min(dataStart, grid.Table.RowCount));
            int sampleEnd = Math.Min(grid.Table.RowCount, dataStart + 5000);
            bool sampled = sampleEnd < grid.Table.RowCount;
            var result = new List<ColumnChoice>(grid.Table.ColumnCount + 1)
            {
                ColumnChoice.None,
            };
            for (int column = 0; column < grid.Table.ColumnCount; column++)
            {
                string label = labelRow.HasValue && labelRow.Value >= 0 && labelRow.Value < grid.Table.RowCount
                    ? grid.Table.GetCellOrNull(labelRow.Value, column)
                    : null;
                int nonBlank = 0;
                var distinct = new HashSet<string>(StringComparer.Ordinal);
                for (int row = dataStart; row < sampleEnd; row++)
                {
                    string value = grid.Table.GetCellOrNull(row, column);
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    nonBlank++;
                    distinct.Add(value);
                }
                int physical = grid.PhysicalColumnForTableIndex(column);
                result.Add(new ColumnChoice(
                    new AlignmentGridIndexReference(column, physical),
                    label,
                    nonBlank,
                    distinct.Count,
                    Math.Max(0, nonBlank - distinct.Count),
                    Math.Max(0, sampleEnd - dataStart),
                    sampled));
            }
            return new ReadOnlyCollection<ColumnChoice>(result);
        }

        private static void InitializeKeyPairs(
            IReadOnlyList<KeyPairEditor> editors,
            AlignmentKeyColumnsConfirmation active,
            KeyCandidate suggested)
        {
            IReadOnlyList<int> before = active != null
                ? active.BeforeColumns.Select(item => item.TableIndex).ToList()
                : suggested?.BeforeColumns ?? Array.Empty<int>();
            IReadOnlyList<int> after = active != null
                ? active.AfterColumns.Select(item => item.TableIndex).ToList()
                : suggested?.AfterColumns ?? Array.Empty<int>();
            for (int index = 0; index < editors.Count; index++)
            {
                int beforeTable = index < before.Count ? before[index] : -1;
                int afterTable = index < after.Count ? after[index] : -1;
                editors[index].Select(beforeTable, afterTable);
            }
        }

        private static KeyCandidate SuggestedKey(KeyInferenceResult inference)
        {
            if (inference == null) return null;
            if (inference.SelectedCandidate != null) return inference.SelectedCandidate;
            return inference.Candidates.Count == 1 ? inference.Candidates[0] : null;
        }

        private static string KeyInferenceSummary(
            KeyInferenceResult inference,
            CanonicalGridProjection beforeGrid,
            CanonicalGridProjection afterGrid)
        {
            if (inference == null) return "No structured key inference result is available.";
            var text = new StringBuilder();
            text.Append("Current inference: ").Append(inference.Decision);
            KeyCandidate candidate = SuggestedKey(inference);
            if (candidate != null)
            {
                text.Append(" · Before ")
                    .Append(string.Join(", ", candidate.BeforeColumns.Select(index =>
                        ExcelColumn(beforeGrid.PhysicalColumnForTableIndex(index)))))
                    .Append(" → After ")
                    .Append(string.Join(", ", candidate.AfterColumns.Select(index =>
                        ExcelColumn(afterGrid.PhysicalColumnForTableIndex(index)))))
                    .Append(" · coverage ")
                    .Append(candidate.BeforeCoverage.ToString("P0", CultureInfo.CurrentCulture))
                    .Append(" / ")
                    .Append(candidate.AfterCoverage.ToString("P0", CultureInfo.CurrentCulture))
                    .Append(" · uniqueness ")
                    .Append(candidate.BeforeUniqueness.ToString("P0", CultureInfo.CurrentCulture))
                    .Append(" / ")
                    .Append(candidate.AfterUniqueness.ToString("P0", CultureInfo.CurrentCulture))
                    .Append(" · overlap ")
                    .Append(candidate.CrossRevisionOverlap.ToString("P0", CultureInfo.CurrentCulture));
                if (candidate.HasDuplicates)
                    text.Append(" · duplicate key values remain explicit");
            }
            return text.ToString();
        }

        private static bool HeaderMatches(
            AlignmentHeaderRowConfirmation active,
            int headerTable,
            int headerPhysical,
            int? dataTable,
            int? dataPhysical)
        {
            if (active.Decision != AlignmentHeaderConfirmationDecision.HeaderRow || active.Row == null) return false;
            return active.Row.TableIndex == headerTable &&
                   active.Row.PhysicalIndex == headerPhysical &&
                   (!dataTable.HasValue
                       ? active.FirstDataRow == null
                       : active.FirstDataRow != null &&
                         active.FirstDataRow.TableIndex == dataTable.Value &&
                         active.FirstDataRow.PhysicalIndex == dataPhysical.Value);
        }

        private static bool KeyMatches(
            AlignmentKeyColumnsConfirmation active,
            IReadOnlyList<AlignmentGridIndexReference> before,
            IReadOnlyList<AlignmentGridIndexReference> after) =>
            ReferenceSequenceEqual(active.BeforeColumns, before) &&
            ReferenceSequenceEqual(active.AfterColumns, after);

        private static bool ReferenceSequenceEqual(
            IReadOnlyList<AlignmentGridIndexReference> left,
            IReadOnlyList<AlignmentGridIndexReference> right)
        {
            if (left.Count != right.Count) return false;
            for (int index = 0; index < left.Count; index++)
                if (!left[index].Equals(right[index])) return false;
            return true;
        }

        private static AlignmentSheetPairConfirmation ActiveSheetPair(
            AlignmentCorrectionGeneration generation,
            IssueModel issue) => issue.Edge == null || string.IsNullOrEmpty(issue.BeforeSheetKey)
            ? null
            : generation.GetSheetPairs(issue.Edge.ParentRevisionId, issue.Edge.ChildRevisionId)
                .SingleOrDefault(item => string.Equals(
                    item.BeforeSheetKey,
                    issue.BeforeSheetKey,
                    StringComparison.Ordinal));

        private static AlignmentSheetUnmatchedConfirmation ActiveSheetUnmatched(
            AlignmentCorrectionGeneration generation,
            IssueModel issue) => issue.Edge == null || string.IsNullOrEmpty(issue.BeforeSheetKey)
            ? null
            : generation.GetUnmatchedSheets(issue.Edge.ParentRevisionId, issue.Edge.ChildRevisionId)
                .SingleOrDefault(item => string.Equals(
                    item.BeforeSheetKey,
                    issue.BeforeSheetKey,
                    StringComparison.Ordinal));

        private static AlignmentHeaderRowConfirmation ActiveHeader(
            AlignmentCorrectionGeneration generation,
            string revisionId,
            string sheetKey)
        {
            generation.TryGetHeaderRow(revisionId, sheetKey, out AlignmentHeaderRowConfirmation active);
            return active;
        }

        private static AlignmentKeyColumnsConfirmation ActiveKey(
            AlignmentCorrectionGeneration generation,
            IssueModel issue)
        {
            if (issue.Edge == null || string.IsNullOrEmpty(issue.BeforeSheetKey) ||
                string.IsNullOrEmpty(issue.AfterSheetKey)) return null;
            generation.TryGetKeyColumns(
                issue.Edge.ParentRevisionId,
                issue.Edge.ChildRevisionId,
                issue.BeforeSheetKey,
                issue.AfterSheetKey,
                out AlignmentKeyColumnsConfirmation active);
            return active;
        }

        private static int? PreferredHeaderRow(HeaderDetectionResult result) =>
            result?.AcceptedHeaderRowIndex ?? result?.CandidateRowIndex;

        private static int PreferredDataStart(HeaderDetectionResult result, int headerRow) =>
            result?.HasAcceptedHeader == true ? result.DataStartRowIndex : headerRow + 1;

        private static string RevisionLabel(ComparisonProjectAnalysisResult analysis, string revisionId)
        {
            if (string.IsNullOrEmpty(revisionId)) return "Project";
            string label = analysis.Session.Setup.Mapping.GetLabel(revisionId);
            return string.IsNullOrEmpty(label) ? ShortId(revisionId) : label;
        }

        private string RevisionLabel(string revisionId) => RevisionLabel(_analysis, revisionId);

        private static string ScopeLabel(IssueModel issue)
        {
            if (issue.EditorKind == IssueEditorKind.SingleHeader)
                return ShortId(issue.SingleRevisionId) + " · " + issue.Finding.SheetName;
            if (issue.Edge == null) return "project scope";
            return issue.Edge.ParentRevisionId == null
                ? issue.Finding.SheetName
                : issue.Finding.SheetName + " · " + ShortId(issue.Finding.FindingId);
        }

        private Control Details(IssueModel issue)
        {
            var box = new GroupBox
            {
                Text = "Finding evidence (read-only)",
                Dock = DockStyle.Top,
                Height = 216,
                Padding = new Padding(10, 22, 10, 8),
                ForeColor = Ink,
            };
            var text = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(250, 250, 252),
                BorderStyle = BorderStyle.FixedSingle,
                Text = FindingDetails(issue),
            };
            SetHelp(text, ComparisonAlignmentReviewHelpText.FindingEvidence);
            SetHelp(box, ComparisonAlignmentReviewHelpText.FindingEvidence);
            box.Controls.Add(text);
            return box;
        }

        private static string FindingDetails(IssueModel issue)
        {
            OpenXmlProjectFinding finding = issue.Finding;
            var text = new StringBuilder();
            text.AppendLine(finding.Message);
            if (!string.IsNullOrWhiteSpace(finding.SuggestedAction))
                text.AppendLine().AppendLine("Suggested action:").AppendLine(finding.SuggestedAction);
            text.AppendLine().Append("Finding ID: ").AppendLine(finding.FindingId);
            text.Append("Before revision: ").AppendLine(finding.BeforeRevisionId);
            text.Append("After revision: ").AppendLine(finding.AfterRevisionId);
            text.Append("Logical sheet ID: ").AppendLine(finding.LogicalSheetId);
            if (finding.ReasonCodes.Count > 0)
                text.Append("Reason codes: ").AppendLine(string.Join(", ", finding.ReasonCodes));
            if (finding.CandidateSheetNames.Count > 0)
                text.Append("Candidate sheets: ").AppendLine(string.Join(", ", finding.CandidateSheetNames));
            return text.ToString();
        }

        private static Control Paragraph(string text, Color? color = null)
        {
            return new Label
            {
                Text = text,
                ForeColor = color ?? Muted,
                AutoSize = false,
                Width = 760,
                Height = Math.Max(48, TextRenderer.MeasureText(
                    text,
                    new Font("Segoe UI", 9F),
                    new Size(760, 1000),
                    TextFormatFlags.WordBreak).Height + 12),
                Margin = new Padding(0, 3, 0, 7),
            };
        }

        private static Label Caption(string text) => new Label
        {
            Text = text,
            ForeColor = Muted,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI Semibold", 8.5F),
            TextAlign = ContentAlignment.BottomLeft,
        };

        private static Label GridHeader(string text) => new Label
        {
            Text = text,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI Semibold", 8.5F),
            ForeColor = Ink,
            BackColor = Surface,
            Padding = new Padding(6, 0, 3, 0),
        };

        private void SetHelp(Control control, string helpText) =>
            SetHelp(_toolTips, control, helpText);

        private static void SetHelp(ToolTip toolTips, Control control, string helpText)
        {
            if (toolTips == null || control == null || string.IsNullOrWhiteSpace(helpText)) return;
            toolTips.SetToolTip(control, helpText);
            control.AccessibleDescription = helpText.Replace("\r\n", " ");
        }

        private static Button ActionButton(string text, int width) => new Button
        {
            Text = text,
            Width = width,
            Height = 32,
            FlatStyle = FlatStyle.System,
            Margin = new Padding(0, 4, 8, 5),
        };

        private static Button FooterButton(string text, int width, EventHandler action)
        {
            var button = new Button
            {
                Text = text,
                Size = new Size(width, 34),
                FlatStyle = FlatStyle.System,
                Margin = new Padding(8, 0, 0, 0),
            };
            button.Click += action;
            return button;
        }

        private static Color SeverityColor(DiagnosticSeverity severity)
        {
            if (severity >= DiagnosticSeverity.ProjectBlocking) return Danger;
            if (severity >= DiagnosticSeverity.Correctable) return Warning;
            return Muted;
        }

        private static string ShortId(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Length <= 12 ? value : value.Substring(0, 12) + "…";
        }

        private static string ExcelColumn(int oneBasedColumn)
        {
            if (oneBasedColumn <= 0) return "?";
            var result = new StringBuilder();
            int value = oneBasedColumn;
            while (value > 0)
            {
                value--;
                result.Insert(0, (char)('A' + value % 26));
                value /= 26;
            }
            return result.ToString();
        }

        private void Inform(string message) =>
            ExcelPowerTools.Interop.Dialogs.Info(message, "Alignment review", this);

        private void Error(string heading, Exception exception) =>
            ExcelPowerTools.Interop.Dialogs.Info(heading + "\r\n\r\n" + exception.Message, "Alignment review", this);

        private enum IssueEditorKind
        {
            ReadOnly,
            SheetPair,
            Header,
            SingleHeader,
            Key,
        }

        private enum HeaderEditorChoice
        {
            Auto,
            NoHeader,
            Header,
        }

        private sealed class IssueModel
        {
            internal IssueModel(
                OpenXmlProjectFinding finding,
                OpenXmlProjectEdgeResult edge,
                OpenXmlProjectSheetAlignment alignment,
                UnresolvedSheetMapping unresolvedSheet,
                IssueEditorKind editorKind,
                string groupKey,
                string groupLabel,
                int sortOrder,
                string beforeSheetKey = "",
                string afterSheetKey = "",
                string singleRevisionId = "",
                string singleSheetKey = "",
                bool isSynthetic = false,
                HeaderDetectionResult singleHeaderDetection = null)
            {
                Finding = finding;
                Edge = edge;
                Alignment = alignment;
                UnresolvedSheet = unresolvedSheet;
                EditorKind = editorKind;
                GroupKey = groupKey;
                GroupLabel = groupLabel;
                SortOrder = sortOrder;
                BeforeSheetKey = string.IsNullOrEmpty(beforeSheetKey)
                    ? alignment?.BeforeSheetKey ?? unresolvedSheet?.Before.SheetKey ?? ""
                    : beforeSheetKey;
                AfterSheetKey = string.IsNullOrEmpty(afterSheetKey)
                    ? alignment?.AfterSheetKey ?? ""
                    : afterSheetKey;
                SingleRevisionId = singleRevisionId ?? "";
                SingleSheetKey = singleSheetKey ?? "";
                IsSynthetic = isSynthetic;
                SingleHeaderDetection = singleHeaderDetection;
            }

            internal OpenXmlProjectFinding Finding { get; }
            internal OpenXmlProjectEdgeResult Edge { get; }
            internal OpenXmlProjectSheetAlignment Alignment { get; }
            internal UnresolvedSheetMapping UnresolvedSheet { get; }
            internal IssueEditorKind EditorKind { get; }
            internal string GroupKey { get; }
            internal string GroupLabel { get; }
            internal int SortOrder { get; }
            internal string BeforeSheetKey { get; }
            internal string AfterSheetKey { get; }
            internal string SingleRevisionId { get; }
            internal string SingleSheetKey { get; }
            internal bool IsSynthetic { get; }
            internal HeaderDetectionResult SingleHeaderDetection { get; }

            internal bool HasSafeRecommendation
            {
                get
                {
                    if (EditorKind == IssueEditorKind.SheetPair)
                        return UnresolvedSheet?.Candidates.Count == 1;
                    if (EditorKind == IssueEditorKind.Header)
                    {
                        int? before = PreferredHeaderRow(Alignment?.BeforeHeader);
                        int? after = PreferredHeaderRow(Alignment?.AfterHeader);
                        return before.HasValue && after.HasValue && before.Value == after.Value &&
                               PreferredDataStart(Alignment.BeforeHeader, before.Value) ==
                               PreferredDataStart(Alignment.AfterHeader, after.Value);
                    }
                    if (EditorKind == IssueEditorKind.Key)
                    {
                        KeyCandidate key = SuggestedKey(Alignment?.KeyInference);
                        return key != null && !key.HasDuplicates;
                    }
                    return false;
                }
            }
        }

        private sealed class TargetChoice
        {
            internal TargetChoice(
                WorkbookSheetFingerprint fingerprint,
                SheetMappingCandidate candidate,
                string inferredSource)
            {
                Fingerprint = fingerprint;
                Candidate = candidate;
                InferredSource = inferredSource;
            }

            internal WorkbookSheetFingerprint Fingerprint { get; }
            internal SheetMappingCandidate Candidate { get; }
            internal string InferredSource { get; }

            public override string ToString()
            {
                string result = Fingerprint.Name + "  [" + ShortId(Fingerprint.SheetKey) + "]";
                if (Candidate != null)
                    result += "  · candidate " + Candidate.Confidence.ToString("P0", CultureInfo.CurrentCulture);
                if (!string.IsNullOrEmpty(InferredSource))
                    result += "  · currently paired with " + InferredSource;
                return result;
            }
        }

        private sealed class HeaderRecommendation
        {
            internal HeaderRecommendation(
                string revisionId,
                string sheetKey,
                int headerTableRow,
                int headerPhysicalRow,
                int? firstDataTableRow,
                int? firstDataPhysicalRow)
            {
                RevisionId = revisionId;
                SheetKey = sheetKey;
                HeaderTableRow = headerTableRow;
                HeaderPhysicalRow = headerPhysicalRow;
                FirstDataTableRow = firstDataTableRow;
                FirstDataPhysicalRow = firstDataPhysicalRow;
            }

            internal string RevisionId { get; }
            internal string SheetKey { get; }
            internal int HeaderTableRow { get; }
            internal int HeaderPhysicalRow { get; }
            internal int? FirstDataTableRow { get; }
            internal int? FirstDataPhysicalRow { get; }
        }

        private sealed class HeaderSideEditor
        {
            private readonly RadioButton _auto;
            private readonly RadioButton _none;
            private readonly RadioButton _header;
            private readonly NumericUpDown _headerRow;
            private readonly CheckBox _separateData;
            private readonly NumericUpDown _firstData;
            private readonly TextBox _preview;

            internal HeaderSideEditor(
                string title,
                CanonicalGridProjection grid,
                HeaderDetectionResult detection,
                AlignmentHeaderRowConfirmation active,
                ToolTip toolTips)
            {
                Grid = grid;
                var group = new GroupBox
                {
                    Text = title,
                    Dock = DockStyle.Fill,
                    Padding = new Padding(10, 21, 10, 9),
                    ForeColor = Ink,
                    Margin = new Padding(0, 0, 7, 0),
                };
                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 3,
                    RowCount = 7,
                };
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112F));
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 106F));
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27F));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 27F));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 31F));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 31F));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

                layout.Controls.Add(new Label
                {
                    Text = DetectionSummary(detection, grid),
                    Dock = DockStyle.Fill,
                    ForeColor = Muted,
                    AutoEllipsis = true,
                }, 0, 0);
                layout.SetColumnSpan(layout.GetControlFromPosition(0, 0), 3);
                _auto = new RadioButton
                {
                    Text = "Auto (read-only current)",
                    AutoSize = true,
                    Checked = active == null,
                    Dock = DockStyle.Fill,
                };
                _none = new RadioButton
                {
                    Text = "No header",
                    AutoSize = true,
                    Checked = active?.Decision == AlignmentHeaderConfirmationDecision.NoHeader,
                    Dock = DockStyle.Fill,
                };
                _header = new RadioButton
                {
                    Text = "Header at Excel row",
                    AutoSize = true,
                    Checked = active?.Decision == AlignmentHeaderConfirmationDecision.HeaderRow,
                    Enabled = grid.Table.RowCount > 0,
                    Dock = DockStyle.Fill,
                };
                layout.Controls.Add(_auto, 0, 1);
                layout.SetColumnSpan(_auto, 3);
                layout.Controls.Add(_none, 0, 2);
                layout.SetColumnSpan(_none, 3);
                layout.Controls.Add(_header, 0, 3);
                layout.SetColumnSpan(_header, 2);

                int firstPhysical = grid.Table.RowCount == 0 ? 1 : grid.FirstPhysicalRow;
                int lastPhysical = grid.Table.RowCount == 0
                    ? firstPhysical
                    : grid.FirstPhysicalRow + grid.Table.RowCount - 1;
                int preferredTable = active?.Row?.TableIndex ??
                                     detection?.AcceptedHeaderRowIndex ??
                                     detection?.CandidateRowIndex ?? 0;
                preferredTable = Math.Max(0, Math.Min(Math.Max(0, grid.Table.RowCount - 1), preferredTable));
                _headerRow = new NumericUpDown
                {
                    Minimum = firstPhysical,
                    Maximum = Math.Max(firstPhysical, lastPhysical),
                    Value = grid.Table.RowCount == 0 ? firstPhysical : grid.FirstPhysicalRow + preferredTable,
                    Dock = DockStyle.Fill,
                    Enabled = grid.Table.RowCount > 0,
                };
                layout.Controls.Add(_headerRow, 2, 3);

                _separateData = new CheckBox
                {
                    Text = "First data row (optional)",
                    AutoSize = true,
                    Checked = active?.FirstDataRow != null,
                    Dock = DockStyle.Fill,
                };
                layout.Controls.Add(_separateData, 0, 4);
                layout.SetColumnSpan(_separateData, 2);
                int maxData = grid.Table.RowCount == 0
                    ? firstPhysical + 1
                    : grid.FirstPhysicalRow + grid.Table.RowCount;
                int preferredData = active?.FirstDataRow?.PhysicalIndex ??
                                    (detection?.HasAcceptedHeader == true
                                        ? grid.FirstPhysicalRow + detection.DataStartRowIndex
                                        : (int)_headerRow.Value + 1);
                preferredData = Math.Max((int)_headerRow.Value + 1, Math.Min(maxData, preferredData));
                _firstData = new NumericUpDown
                {
                    Minimum = Math.Min(maxData, (int)_headerRow.Value + 1),
                    Maximum = Math.Max(maxData, (int)_headerRow.Value + 1),
                    Value = preferredData,
                    Dock = DockStyle.Fill,
                };
                layout.Controls.Add(_firstData, 2, 4);
                layout.Controls.Add(new Label
                {
                    Text = "±3 row preview (alignment text only)",
                    Dock = DockStyle.Fill,
                    ForeColor = Muted,
                    Font = new Font("Segoe UI Semibold", 8F),
                }, 0, 5);
                layout.SetColumnSpan(layout.GetControlFromPosition(0, 5), 3);
                _preview = new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false,
                    Font = new Font("Consolas", 8.25F),
                    BackColor = Color.FromArgb(250, 250, 252),
                };
                SetHelp(toolTips, group, ComparisonAlignmentReviewHelpText.HeaderEditor);
                SetHelp(toolTips, _auto, ComparisonAlignmentReviewHelpText.AutomaticHeader);
                SetHelp(toolTips, _none, ComparisonAlignmentReviewHelpText.NoHeader);
                SetHelp(toolTips, _header, ComparisonAlignmentReviewHelpText.HeaderRow);
                SetHelp(toolTips, _headerRow, ComparisonAlignmentReviewHelpText.HeaderRowNumber);
                SetHelp(toolTips, _separateData, ComparisonAlignmentReviewHelpText.SeparateFirstDataRow);
                SetHelp(toolTips, _firstData, ComparisonAlignmentReviewHelpText.FirstDataRowNumber);
                SetHelp(toolTips, _preview, ComparisonAlignmentReviewHelpText.HeaderPreview);
                layout.Controls.Add(_preview, 0, 6);
                layout.SetColumnSpan(_preview, 3);

                _headerRow.ValueChanged += (_, __) =>
                {
                    int minimum = Math.Min(maxData, (int)_headerRow.Value + 1);
                    _firstData.Minimum = minimum;
                    if (_firstData.Value < minimum) _firstData.Value = minimum;
                    _header.Checked = true;
                    RefreshPreview();
                };
                _separateData.CheckedChanged += (_, __) =>
                {
                    _firstData.Enabled = _separateData.Checked && _header.Checked;
                    if (_separateData.Checked) _header.Checked = true;
                };
                _firstData.ValueChanged += (_, __) => _header.Checked = true;
                EventHandler updateEnabled = (_, __) =>
                {
                    _headerRow.Enabled = _header.Checked && Grid.Table.RowCount > 0;
                    _separateData.Enabled = _header.Checked;
                    _firstData.Enabled = _header.Checked && _separateData.Checked;
                };
                _auto.CheckedChanged += updateEnabled;
                _none.CheckedChanged += updateEnabled;
                _header.CheckedChanged += updateEnabled;
                updateEnabled(null, EventArgs.Empty);
                RefreshPreview();
                group.Controls.Add(layout);
                Control = group;
            }

            internal Control Control { get; }
            internal CanonicalGridProjection Grid { get; }
            internal HeaderEditorChoice Choice => _header.Checked
                ? HeaderEditorChoice.Header
                : _none.Checked ? HeaderEditorChoice.NoHeader : HeaderEditorChoice.Auto;
            internal int HeaderPhysicalRow => (int)_headerRow.Value;
            internal bool HasFirstDataRow => _separateData.Checked;
            internal int FirstDataPhysicalRow => (int)_firstData.Value;

            private void RefreshPreview()
            {
                if (Grid.Table.RowCount == 0)
                {
                    _preview.Text = "(empty physical grid)";
                    return;
                }
                int center = (int)_headerRow.Value - Grid.FirstPhysicalRow;
                int start = Math.Max(0, center - 3);
                int end = Math.Min(Grid.Table.RowCount - 1, center + 3);
                var text = new StringBuilder();
                for (int row = start; row <= end; row++)
                {
                    text.Append(row == center ? "▶ " : "  ")
                        .Append("R")
                        .Append(Grid.FirstPhysicalRow + row)
                        .Append(" | ");
                    int columns = Math.Min(8, Grid.Table.ColumnCount);
                    for (int column = 0; column < columns; column++)
                    {
                        if (column > 0) text.Append(" | ");
                        text.Append(PreviewValue(Grid.Table.GetCellOrNull(row, column)));
                    }
                    if (Grid.Table.ColumnCount > columns) text.Append(" | …");
                    if (row < end) text.AppendLine();
                }
                _preview.Text = text.ToString();
            }

            private static string DetectionSummary(HeaderDetectionResult result, CanonicalGridProjection grid)
            {
                if (result == null) return "Current detector result is unavailable.";
                string candidate = result.CandidateRowIndex.HasValue && grid.Table.RowCount > 0
                    ? " · candidate Excel row " +
                      (grid.FirstPhysicalRow + result.CandidateRowIndex.Value).ToString(CultureInfo.InvariantCulture)
                    : "";
                string accepted = result.AcceptedHeaderRowIndex.HasValue && grid.Table.RowCount > 0
                    ? " · accepted Excel row " +
                      (grid.FirstPhysicalRow + result.AcceptedHeaderRowIndex.Value).ToString(CultureInfo.InvariantCulture)
                    : "";
                return "Auto: " + result.Decision + accepted + candidate +
                       " · confidence " + result.Confidence.ToString("P0", CultureInfo.CurrentCulture);
            }

            internal static string PreviewValue(string value)
            {
                if (value == null) return "∅";
                string safe = value.Replace("\r\n", "↵").Replace("\r", "↵").Replace("\n", "↵")
                    .Replace("\t", "⇥");
                return safe.Length <= 32 ? safe : safe.Substring(0, 31) + "…";
            }
        }

        private sealed class ColumnChoice
        {
            internal static readonly ColumnChoice None = new ColumnChoice();

            private ColumnChoice()
            {
                Reference = null;
                Label = "(not used)";
            }

            internal ColumnChoice(
                AlignmentGridIndexReference reference,
                string label,
                int nonBlank,
                int distinct,
                int duplicates,
                int sampledRows,
                bool sampled)
            {
                Reference = reference;
                Label = label;
                NonBlank = nonBlank;
                Distinct = distinct;
                Duplicates = duplicates;
                SampledRows = sampledRows;
                Sampled = sampled;
            }

            internal AlignmentGridIndexReference Reference { get; }
            internal string Label { get; }
            internal int NonBlank { get; }
            internal int Distinct { get; }
            internal int Duplicates { get; }
            internal int SampledRows { get; }
            internal bool Sampled { get; }

            public override string ToString()
            {
                if (Reference == null) return Label;
                string header = string.IsNullOrWhiteSpace(Label) ? "(blank header)" : HeaderSideEditor.PreviewValue(Label);
                return ExcelColumn(Reference.PhysicalIndex) + " (" +
                       Reference.PhysicalIndex.ToString(CultureInfo.InvariantCulture) + ") · " + header +
                       " · " + NonBlank.ToString("N0", CultureInfo.CurrentCulture) + "/" +
                       SampledRows.ToString("N0", CultureInfo.CurrentCulture) + " nonblank · " +
                       Distinct.ToString("N0", CultureInfo.CurrentCulture) + " distinct · " +
                       Duplicates.ToString("N0", CultureInfo.CurrentCulture) + " duplicate" +
                       (Sampled ? " · sample" : "");
            }
        }

        private sealed class KeyPairEditor
        {
            internal KeyPairEditor(
                int order,
                IEnumerable<ColumnChoice> before,
                IEnumerable<ColumnChoice> after,
                ToolTip toolTips)
            {
                Order = order;
                Before = Combo(before);
                After = Combo(after);
                SetHelp(toolTips, Before, ComparisonAlignmentReviewHelpText.BeforeKeyColumn);
                SetHelp(toolTips, After, ComparisonAlignmentReviewHelpText.AfterKeyColumn);
            }

            internal int Order { get; }
            internal ComboBox Before { get; }
            internal ComboBox After { get; }
            internal ColumnChoice BeforeChoice => Before.SelectedItem as ColumnChoice;
            internal ColumnChoice AfterChoice => After.SelectedItem as ColumnChoice;
            internal bool HasEitherSelection => BeforeChoice?.Reference != null || AfterChoice?.Reference != null;
            internal bool HasCompleteSelection => BeforeChoice?.Reference != null && AfterChoice?.Reference != null;

            internal void Select(int beforeTableIndex, int afterTableIndex)
            {
                Select(Before, beforeTableIndex);
                Select(After, afterTableIndex);
            }

            private static ComboBox Combo(IEnumerable<ColumnChoice> choices)
            {
                var combo = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    IntegralHeight = false,
                    MaxDropDownItems = 16,
                    Margin = new Padding(3, 4, 3, 3),
                    DropDownWidth = 560,
                };
                foreach (ColumnChoice choice in choices) combo.Items.Add(choice);
                if (combo.Items.Count > 0) combo.SelectedIndex = 0;
                return combo;
            }

            private static void Select(ComboBox combo, int tableIndex)
            {
                for (int index = 0; index < combo.Items.Count; index++)
                {
                    var choice = combo.Items[index] as ColumnChoice;
                    if ((choice?.Reference?.TableIndex ?? -1) != tableIndex) continue;
                    combo.SelectedIndex = index;
                    return;
                }
                combo.SelectedIndex = combo.Items.Count > 0 ? 0 : -1;
            }
        }
    }

    internal static class ComparisonAlignmentReviewHelpText
    {
        internal const string Findings =
            "Select a finding to inspect its exact evidence and available correction. Ctrl-click or Shift-click selects several findings for the right-click batch action.";
        internal const string Editor =
            "Shows the structured correction editor for the primary selected finding. Displayed diagnostic prose is never parsed as identity.";
        internal const string Reviewer =
            "Optional reviewer name stored with the next queued correction event. Leave blank to use an anonymous reviewer snapshot.";
        internal const string AuditNote =
            "Optional note copied into every correction event queued by the next action or batch action.";
        internal const string WorkingStatus =
            "Summarizes active confirmations and new audit events. Any queued change must be applied by rerunning analysis.";
        internal const string Cancel =
            "Close alignment review and discard corrections queued during this window.";
        internal const string ContinueToReview =
            "Continue to the final Review workbook. If corrections were changed, analysis reruns first against the retained immutable inputs.";
        internal const string BuildDiagnosticReview =
            "Create an explicitly incomplete diagnostic Review workbook that preserves current findings and correction audit evidence.";
        internal const string RerunAnalysis =
            "Apply queued corrections by rebuilding alignment and logical history from the retained immutable input snapshots. Enabled after a correction changes.";
        internal const string QueueSafeRecommendations =
            "Queue only unambiguous recommendations for every selected finding. Findings without one safe recommendation are left unchanged.";
        internal const string OpenPrimarySelection =
            "Open the structured editor for the first selected finding.";
        internal const string TargetSheet =
            "Choose the exact target-sheet occurrence that continues the selected source sheet. Candidate confidence and any displaced inferred source are shown in the list.";
        internal const string QueueSheetPair =
            "Queue an exact source-to-target sheet identity. This can replace an inferred pair or an explicit deleted/new choice for the same source.";
        internal const string KeepSheetsUnmatched =
            "Confirm that the source sheet was deleted and is not a renamed target. Unclaimed target sheets remain new additions; their contents are not compared as the same sheet.";
        internal const string AutomaticSheetMapping =
            "Clear the active explicit pair or deleted/new choice for this source and return it to automatic mapping on the next analysis.";
        internal const string QueuePairedHeaders =
            "Queue the Before and After header choices together. Confirmed header and first-data positions must have matching table-relative shapes.";
        internal const string QueueSingleHeader =
            "Queue, replace, or clear the explicit header decision for this exact revision and sheet occurrence.";
        internal const string HeaderEditor =
            "Choose automatic detection, no header, or an exact physical Excel header row for this sheet occurrence.";
        internal const string AutomaticHeader =
            "Keep automatic header detection in control. Queuing this choice clears any active explicit header override for this sheet.";
        internal const string NoHeader =
            "Explicitly confirm that this sheet has no header row.";
        internal const string HeaderRow =
            "Explicitly confirm that this sheet has a header at the selected physical Excel row.";
        internal const string HeaderRowNumber =
            "The one-based physical Excel row containing the header. Changing it selects the explicit Header option.";
        internal const string SeparateFirstDataRow =
            "Enable this when units, subheaders, or separator rows occur between the header and the first data row.";
        internal const string FirstDataRowNumber =
            "The one-based physical Excel row where keyed data starts. It must follow the selected header row.";
        internal const string HeaderPreview =
            "Read-only text preview of three neighboring physical rows on either side of the selected header row.";
        internal const string AutomaticKey =
            "Keep row-key inference automatic. Queuing this choice clears the active explicit key for this exact sheet pair.";
        internal const string ExplicitKey =
            "Confirm one to four ordered Before-to-After physical column pairs as the row identity key.";
        internal const string BeforeKeyColumn =
            "Choose the next ordered key column from the Before sheet, or choose not used for this row.";
        internal const string AfterKeyColumn =
            "Choose the corresponding ordered key column from the After sheet, or choose not used for this row.";
        internal const string QueueKey =
            "Queue the automatic or explicit ordered key decision for this exact mapped sheet pair.";
        internal const string FindingEvidence =
            "Read-only verbose finding evidence, including the stable finding ID, exact revision scope, reason codes, candidate sheets, and suggested action.";
    }
}
