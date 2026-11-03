using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.Windows.Forms;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Honest in-pane home for the comparison workflow. A selected Review.xlsx is retained as a complete,
    /// validated hydration session while summary controls render only privacy-safe metadata. An attached decision
    /// ledger can be reviewed in memory and published only through a newly named, copy-on-write Review revision.
    /// </summary>
    internal sealed class ComparisonWorkspacePage : UserControl
    {
        private static readonly Color Background = Color.FromArgb(246, 248, 252);
        private static readonly Color Surface = Color.White;
        private static readonly Color Ink = Color.FromArgb(28, 35, 56);
        private static readonly Color Muted = Color.FromArgb(101, 112, 136);
        private static readonly Color Border = Color.FromArgb(221, 226, 236);
        private static readonly Color Accent = Color.FromArgb(75, 78, 225);
        private static readonly Color ComparisonAccent = Color.FromArgb(104, 70, 198);
        private static readonly Color Ready = Color.FromArgb(31, 139, 91);
        private static readonly Color DisabledInk = Color.FromArgb(143, 151, 169);

        private readonly FlowLayoutPanel _content = new FlowLayoutPanel();
        private readonly GradientSurface _header = new GradientSurface();
        private readonly RoundedSurface _actions = new RoundedSurface();
        private readonly RoundedSurface _loadedReview = new RoundedSurface();
        private readonly RoundedSurface _workflow = new RoundedSurface();
        private readonly RoundedSurface _alpha = new RoundedSurface();
        private readonly ComparisonDecisionLedgerControl _decisionReview =
            new ComparisonDecisionLedgerControl();
        private readonly ActionCard _newComparison;
        private readonly ActionCard _openReview;
        private readonly List<WorkflowStageRow> _stages = new List<WorkflowStageRow>();
        private readonly List<AlphaInfoRow> _alphaRows = new List<AlphaInfoRow>();
        private readonly SmoothLabel _headerTitle;
        private readonly SmoothLabel _headerSubtitle;
        private readonly PillLabel _alphaBadge;
        private readonly NavigationButton _back;
        private readonly PillLabel _integrityBadge;
        private readonly SmoothLabel _loadedFileName;
        private readonly SmoothLabel _loadedPath;
        private readonly SmoothLabel _loadedProject;
        private readonly SmoothLabel _loadedFileMetadata;
        private readonly SmoothLabel _loadedStateMetadata;
        private readonly SmoothLabel _loadedProjectionMetadata;
        private readonly RoundedSurface _loadedDecisionLedger = new RoundedSurface();
        private readonly SmoothLabel _loadedDecisionLedgerTitle;
        private readonly SmoothLabel _loadedDecisionLedgerMetadata;
        private readonly SmoothLabel _loadedLimitations;
        private readonly WorkspaceButton _openLoadedReview;
        private readonly WorkspaceButton _chooseAnotherReview;
        private readonly ToolTip _toolTip = new ToolTip();
        private ReviewWorkbookHydrationSession _loadedSession;

        internal ComparisonWorkspacePage()
        {
            Name = "ComparisonWorkspacePage";
            Dock = DockStyle.Fill;
            MinimumSize = new Size(240, 320);
            BackColor = Background;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            DoubleBuffered = true;

            _content.Dock = DockStyle.Fill;
            _content.Margin = Padding.Empty;
            _content.Padding = new Padding(12, 12, 12, 18);
            _content.BackColor = Background;
            _content.AutoScroll = true;
            _content.WrapContents = false;
            _content.FlowDirection = FlowDirection.TopDown;
            Controls.Add(_content);

            _header.StartColor = Color.FromArgb(25, 34, 67);
            _header.EndColor = Color.FromArgb(91, 66, 202);
            _header.Height = 168;
            _header.Margin = new Padding(0, 0, 0, 10);

            _back = new NavigationButton
            {
                Text = "‹  Back to tools",
                Location = new Point(14, 14),
                Size = new Size(126, 30),
                TabIndex = 0,
                AccessibleName = "Back to all Power Tools",
            };
            _back.Click += (_, __) =>
            {
                if (ConfirmAbandonPending()) BackRequested?.Invoke(this, EventArgs.Empty);
            };
            _header.Controls.Add(_back);

            _alphaBadge = new PillLabel
            {
                Text = "ALPHA",
                Size = new Size(55, 22),
                Location = new Point(0, 17),
                FillColor = Color.FromArgb(55, 255, 255, 255),
                BorderColor = Color.FromArgb(80, 255, 255, 255),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 7f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
            };
            _header.Controls.Add(_alphaBadge);

            _headerTitle = new SmoothLabel
            {
                Text = "Comparison workspace",
                Location = new Point(16, 57),
                Height = 34,
                BackColor = Color.Transparent,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 14f, FontStyle.Bold),
                AutoEllipsis = true,
            };
            _header.Controls.Add(_headerTitle);

            _headerSubtitle = new SmoothLabel
            {
                Text = "Compare stored values, saved calculated values, and formulas across aligned revisions, then publish Review.xlsx.",
                Location = new Point(17, 96),
                Height = 53,
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(220, 225, 249),
                Font = new Font("Segoe UI", 9f),
            };
            _header.Controls.Add(_headerSubtitle);
            _header.Resize += (_, __) => ArrangeHeader();
            _content.Controls.Add(_header);

            _actions.FillColor = Surface;
            _actions.BorderColor = Border;
            _actions.Height = 228;
            _actions.Margin = new Padding(0, 0, 0, 10);
            AddSectionHeading(
                _actions,
                "START HERE",
                "Create a new comparison or validate and open an existing Review.xlsx.");

            _newComparison = new ActionCard(
                "NEW",
                "New comparison",
                "Select two or more revision files and begin setup.",
                primary: true)
            {
                TabIndex = 1,
                AccessibleName = "New comparison",
                AccessibleDescription = "Start a new multi-revision comparison project.",
            };
            _newComparison.Activated += (_, __) =>
            {
                if (ConfirmAbandonPending()) NewComparisonRequested?.Invoke(this, EventArgs.Empty);
            };
            _actions.Controls.Add(_newComparison);

            _openReview = new ActionCard(
                "OPEN",
                "Open validated Review",
                "Verify an existing Review.xlsx before Excel opens it read-only.",
                primary: false)
            {
                TabIndex = 2,
                AccessibleName = "Open validated Review workbook",
                AccessibleDescription = "Validate and open an existing comparison Review.xlsx workbook.",
            };
            _openReview.Activated += (_, __) =>
            {
                if (ConfirmAbandonPending()) ChooseReviewRequested?.Invoke(this, EventArgs.Empty);
            };
            _actions.Controls.Add(_openReview);
            _actions.Resize += (_, __) => ArrangeActions();
            _content.Controls.Add(_actions);

            _loadedReview.FillColor = Color.FromArgb(252, 255, 253);
            _loadedReview.BorderColor = Color.FromArgb(177, 223, 200);
            _loadedReview.Height = 336;
            _loadedReview.Margin = new Padding(0, 0, 0, 10);
            _loadedReview.Visible = false;
            AddSectionHeading(
                _loadedReview,
                "PROJECT LOADED · COPY-ON-WRITE",
                "The source remains unchanged; reviewed decisions publish only into a newly named Review revision.");

            _integrityBadge = new PillLabel
            {
                Text = "LOADED · SAFE",
                Size = new Size(132, 22),
                FillColor = Color.FromArgb(226, 247, 237),
                BorderColor = Color.FromArgb(159, 219, 189),
                ForeColor = Ready,
                Font = new Font("Segoe UI Semibold", 6.8f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
            };
            _loadedReview.Controls.Add(_integrityBadge);

            _loadedFileName = LoadedSummaryLabel(
                new Font("Segoe UI Semibold", 11f, FontStyle.Bold),
                Ink,
                25);
            _loadedReview.Controls.Add(_loadedFileName);
            _loadedPath = LoadedSummaryLabel(new Font("Segoe UI", 7.7f), Muted, 31);
            _loadedReview.Controls.Add(_loadedPath);
            _loadedProject = LoadedSummaryLabel(new Font("Segoe UI Semibold", 8f), Ink, 21);
            _loadedReview.Controls.Add(_loadedProject);
            _loadedFileMetadata = LoadedSummaryLabel(new Font("Segoe UI", 7.8f), Muted, 21);
            _loadedReview.Controls.Add(_loadedFileMetadata);
            _loadedStateMetadata = LoadedSummaryLabel(new Font("Segoe UI", 7.8f), Muted, 21);
            _loadedReview.Controls.Add(_loadedStateMetadata);
            _loadedProjectionMetadata = LoadedSummaryLabel(new Font("Segoe UI", 7.8f), Muted, 21);
            _loadedReview.Controls.Add(_loadedProjectionMetadata);

            _loadedDecisionLedger.FillColor = Color.FromArgb(249, 248, 255);
            _loadedDecisionLedger.BorderColor = Color.FromArgb(211, 204, 239);
            _loadedDecisionLedger.Visible = false;
            _loadedDecisionLedger.AccessibleRole = AccessibleRole.StaticText;
            _loadedDecisionLedger.AccessibleName = "Decision ledger initialized";
            _loadedDecisionLedgerTitle = LoadedSummaryLabel(
                new Font("Segoe UI Semibold", 8.2f, FontStyle.Bold),
                ComparisonAccent,
                19);
            _loadedDecisionLedgerTitle.Text = "✓  Decision ledger initialized";
            _loadedDecisionLedger.Controls.Add(_loadedDecisionLedgerTitle);
            _loadedDecisionLedgerMetadata = LoadedSummaryLabel(
                new Font("Segoe UI", 7.4f),
                Muted,
                30);
            _loadedDecisionLedgerMetadata.AutoEllipsis = false;
            _loadedDecisionLedger.Controls.Add(_loadedDecisionLedgerMetadata);
            _loadedReview.Controls.Add(_loadedDecisionLedger);

            _loadedLimitations = LoadedSummaryLabel(new Font("Segoe UI", 7.7f), WarningColor(), 43);
            _loadedLimitations.AutoEllipsis = false;
            _loadedLimitations.Text =
                "The loaded source stays immutable. Pending task-pane decisions are saved only to a new Review.xlsx. " +
                "Opening a protected Excel copy still revalidates the exact selected occurrence.";
            _loadedReview.Controls.Add(_loadedLimitations);

            _openLoadedReview = new WorkspaceButton(primary: true)
            {
                Text = "Open protected copy",
                TabIndex = 3,
                AccessibleName = "Open protected read-only Review workbook copy",
                AccessibleDescription = "Revalidate the selected Review workbook and open a protected read-only copy in Excel.",
            };
            _openLoadedReview.Click += (_, __) =>
            {
                if (_decisionReview.HasPendingDecisions)
                {
                    Dialogs.Info(
                        "Save or discard the pending review decisions before opening the protected Excel copy.");
                    return;
                }
                ReviewWorkbookHydrationSession current = _loadedSession;
                if (current != null)
                    OpenValidatedReviewRequested?.Invoke(
                        this,
                        new ReviewWorkbookOpenRequestedEventArgs(current));
            };
            _loadedReview.Controls.Add(_openLoadedReview);

            _chooseAnotherReview = new WorkspaceButton(primary: false)
            {
                Text = "Choose another",
                TabIndex = 4,
                AccessibleName = "Choose another Review workbook",
            };
            _chooseAnotherReview.Click += (_, __) =>
            {
                if (ConfirmAbandonPending()) ChooseReviewRequested?.Invoke(this, EventArgs.Empty);
            };
            _loadedReview.Controls.Add(_chooseAnotherReview);
            _loadedReview.Resize += (_, __) => ArrangeLoadedReview();
            _content.Controls.Add(_loadedReview);

            _decisionReview.Visible = false;
            _decisionReview.CommitRequested += (_, e) =>
                DecisionCommitRequested?.Invoke(this, e);
            _content.Controls.Add(_decisionReview);

            _workflow.FillColor = Surface;
            _workflow.BorderColor = Border;
            _workflow.Height = 410;
            _workflow.Margin = new Padding(0, 0, 0, 10);
            AddSectionHeading(
                _workflow,
                "WORKFLOW STATUS",
                "What the alpha can do today, and what remains intentionally unavailable.");

            AddStage("Setup", "Revision order, parent links, labels, baseline and master.", true);
            AddStage("Alignment", "Sheets, headers, data starts and key columns; uncertain cases require confirmation.", true);
            AddStage("Review.xlsx", "History of stored values, saved calculated values, and formulas, with integrity-bound project state.", true);
            AddStage("Review decisions", "Accept, reject and annotate individual or grouped changes.", true);
            AddStage("Snapshots", "Neighboring-row evidence and expandable revision context.", false);
            AddStage("Merge", "Copy selected older-revision values into a new master revision with an audit log.", false);
            AddStage("Archive", "Export inputs, report and snapshots as one portable project package.", false);
            _workflow.Resize += (_, __) => ArrangeWorkflow();
            _content.Controls.Add(_workflow);

            _alpha.FillColor = Surface;
            _alpha.BorderColor = Border;
            _alpha.Height = 300;
            _alpha.Margin = new Padding(0, 0, 0, 8);
            AddSectionHeading(
                _alpha,
                "CURRENT ALPHA ENVELOPE",
                "Format coverage and fail-closed limits in this build.");
            AddAlphaInfo(
                "FORMATS",
                "Multi-revision input: .xlsx and .xlsm. Output: one newly created Review.xlsx.");
            AddAlphaInfo(
                "OPENING",
                "Review.xlsx is integrity-validated, copied to a protected launch location and opened read-only.");
            AddAlphaInfo(
                "SAFETY",
                "768 MiB aggregate input; 2M stored cells per revision / 10M audited; 2M retained useful-range cells; 10M bounded grid cells.");
            AddAlphaInfo(
                "BEHAVIOR",
                "Ambiguous identity blocks history and asks for correction; the engine never silently guesses.");
            _alpha.Resize += (_, __) => ArrangeAlphaSummary();
            _content.Controls.Add(_alpha);

            _content.Resize += (_, __) => UpdateResponsiveLayout();
            Resize += (_, __) => UpdateResponsiveLayout();
            UpdateResponsiveLayout();
        }

        internal event EventHandler BackRequested;
        internal event EventHandler NewComparisonRequested;
        internal event EventHandler ChooseReviewRequested;
        internal event EventHandler<ReviewWorkbookOpenRequestedEventArgs> OpenValidatedReviewRequested;
        internal event EventHandler<ComparisonDecisionCommitRequestedEventArgs> DecisionCommitRequested;

        internal ReviewWorkbookHydrationSession LoadedSession => _loadedSession;
        internal ReviewWorkbookValidationSummary LoadedReview => _loadedSession?.Summary;

        internal void ShowHydratedReview(ReviewWorkbookHydrationSession session)
        {
            _loadedSession = session ?? throw new ArgumentNullException(nameof(session));
            ReviewWorkbookValidationSummary summary = session.Summary;
            _loadedFileName.Text = summary.FileName;
            _loadedPath.Text = summary.FullPath;
            _loadedProject.Text = "Project  ·  " + summary.ProjectId;
            _loadedFileMetadata.Text = "File  ·  " + FormatBytes(summary.FileByteLength) + "  ·  modified " +
                                       summary.ModifiedUtc.UtcDateTime.ToString(
                                           "yyyy-MM-dd HH:mm:ss 'UTC'",
                                           CultureInfo.InvariantCulture);
            _loadedStateMetadata.Text = "State  ·  " + FormatBytes(summary.StateByteCount) + "  ·  " +
                                        summary.StateChunkCount.ToString("N0", CultureInfo.CurrentCulture) +
                                        (summary.StateChunkCount == 1 ? " chunk" : " chunks") + "  ·  SHA " +
                                        summary.StateSha256Prefix + "…";
            _loadedProjectionMetadata.Text = "Visible projection  ·  SHA " +
                                             summary.ProjectionSha256Prefix + "…";
            _loadedDecisionLedger.Visible = summary.HasDecisionSidecar;
            if (summary.HasDecisionSidecar)
            {
                _loadedDecisionLedgerMetadata.Text =
                    FormatCount(summary.DecisionEvidenceRevisionCount, "evidence revision") +
                    "  ·  " +
                    FormatCount(summary.DecisionSubjectCount, "subject") +
                    "  ·  " +
                    FormatCount(summary.DecisionObservationCount, "observation") +
                    "  ·  " +
                    FormatCount(summary.ActiveDecisionCount, "active decision") +
                    Environment.NewLine +
                    "Catalog " + summary.DecisionCatalogIdPrefix + "…  ·  Generation " +
                    summary.DecisionGenerationIdPrefix + "…";
                _loadedDecisionLedger.AccessibleDescription =
                    "Validated immutable decision evidence is present. " +
                    _loadedDecisionLedgerMetadata.Text.Replace(Environment.NewLine, ". ");
                _toolTip.SetToolTip(
                    _loadedDecisionLedgerMetadata,
                    _loadedDecisionLedgerMetadata.Text.Replace(Environment.NewLine, "  "));
                _decisionReview.LoadSession(session);
                _decisionReview.Visible = true;
            }
            else
            {
                _loadedDecisionLedgerMetadata.Text = "";
                _loadedDecisionLedger.AccessibleDescription = "";
                _toolTip.SetToolTip(_loadedDecisionLedgerMetadata, "");
                _decisionReview.ClearSession();
                _decisionReview.Visible = false;
            }
            _toolTip.SetToolTip(_loadedFileName, summary.FileName);
            _toolTip.SetToolTip(_loadedPath, summary.FullPath);
            _toolTip.SetToolTip(_loadedProject, summary.ProjectId);
            _loadedReview.Visible = true;
            UpdateResponsiveLayout();
            ArrangeLoadedReview();
            if (summary.HasDecisionSidecar)
            {
                _content.ScrollControlIntoView(_decisionReview);
                _decisionReview.FocusPageStart();
            }
            else
            {
                _content.ScrollControlIntoView(_loadedReview);
                if (_openLoadedReview.CanFocus) _openLoadedReview.Focus();
            }
        }

        /// <summary>
        /// Discards the complete hydrated state and every rendered identity. When an expected session is supplied,
        /// a later selection wins and cannot be cleared by a stale queued interaction.
        /// </summary>
        internal bool ClearLoadedReview(ReviewWorkbookHydrationSession expectedSession = null)
        {
            if (expectedSession != null && !ReferenceEquals(_loadedSession, expectedSession)) return false;
            _loadedSession = null;
            _loadedReview.Visible = false;
            _loadedFileName.Text = "";
            _loadedPath.Text = "";
            _loadedProject.Text = "";
            _loadedFileMetadata.Text = "";
            _loadedStateMetadata.Text = "";
            _loadedProjectionMetadata.Text = "";
            _loadedDecisionLedger.Visible = false;
            _loadedDecisionLedgerMetadata.Text = "";
            _loadedDecisionLedger.AccessibleDescription = "";
            _decisionReview.ClearSession();
            _decisionReview.Visible = false;
            _toolTip.SetToolTip(_loadedFileName, "");
            _toolTip.SetToolTip(_loadedPath, "");
            _toolTip.SetToolTip(_loadedProject, "");
            _toolTip.SetToolTip(_loadedDecisionLedgerMetadata, "");
            UpdateResponsiveLayout();
            return true;
        }

        internal void FocusPageStart()
        {
            if (!Visible) return;
            _content.AutoScrollPosition = Point.Empty;
            if (_back.CanFocus) _back.Focus();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _loadedSession = null;
                _decisionReview.ClearSession();
                _toolTip.Dispose();
            }
            base.Dispose(disposing);
        }

        private static SmoothLabel LoadedSummaryLabel(Font font, Color color, int height) => new SmoothLabel
        {
            Height = height,
            BackColor = Color.Transparent,
            ForeColor = color,
            Font = font,
            AutoEllipsis = true,
        };

        private static Color WarningColor() => Color.FromArgb(139, 82, 16);

        private static string FormatBytes(long value)
        {
            string[] units = { "B", "KiB", "MiB", "GiB" };
            double display = Math.Max(0, value);
            int unit = 0;
            while (display >= 1024D && unit < units.Length - 1)
            {
                display /= 1024D;
                unit++;
            }
            return display.ToString(unit == 0 ? "N0" : "N1", CultureInfo.CurrentCulture) + " " + units[unit];
        }

        private static string FormatCount(int value, string singular) =>
            value.ToString("N0", CultureInfo.CurrentCulture) + " " + singular + (value == 1 ? "" : "s");

        private static int MeasureWrappedText(string text, Font font, int width)
        {
            string normalized = (text ?? "").Replace("\r\n", "\n").Replace('\r', '\n');
            int measuredHeight = 0;
            using (var bitmap = new Bitmap(1, 1))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                foreach (string line in normalized.Split('\n'))
                {
                    SizeF measured = graphics.MeasureString(
                        string.IsNullOrEmpty(line) ? " " : line,
                        font,
                        new SizeF(Math.Max(1, width), 32767f));
                    measuredHeight += Math.Max(font.Height, (int)Math.Ceiling(measured.Height));
                }
            }
            return Math.Max(font.Height, measuredHeight);
        }

        private static void AddSectionHeading(Control parent, string heading, string description)
        {
            parent.Controls.Add(new SmoothLabel
            {
                Name = "SectionHeading",
                Text = heading,
                Location = new Point(16, 13),
                Height = 19,
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(98, 107, 132),
                Font = new Font("Segoe UI Semibold", 7.4f, FontStyle.Bold),
                AutoEllipsis = true,
            });
            parent.Controls.Add(new SmoothLabel
            {
                Name = "SectionDescription",
                Text = description,
                Location = new Point(16, 34),
                Height = 27,
                BackColor = Color.Transparent,
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8.2f),
                AutoEllipsis = true,
            });
        }

        private void AddStage(string title, string description, bool implemented)
        {
            var row = new WorkflowStageRow(title, description, implemented);
            _stages.Add(row);
            _workflow.Controls.Add(row);
        }

        private void AddAlphaInfo(string label, string value)
        {
            var row = new AlphaInfoRow(label, value);
            _alphaRows.Add(row);
            _alpha.Controls.Add(row);
        }

        private void UpdateResponsiveLayout()
        {
            if (_content.IsDisposed || _content.ClientSize.Width <= 0) return;
            int width = Math.Max(
                214,
                _content.ClientSize.Width - _content.Padding.Horizontal -
                SystemInformation.VerticalScrollBarWidth - 5);
            foreach (Control section in new Control[]
                     {
                         _header,
                         _actions,
                         _loadedReview,
                         _decisionReview,
                         _workflow,
                         _alpha,
                     })
                section.Width = width;

            _actions.Height = width >= 500 ? 139 : 228;
            ArrangeHeader();
            ArrangeActions();
            ArrangeLoadedReview();
            ArrangeWorkflow();
            ArrangeAlphaSummary();
        }

        private bool ConfirmAbandonPending()
        {
            if (!_decisionReview.HasPendingDecisions) return true;
            bool abandon = Dialogs.Confirm(
                "Discard " +
                _decisionReview.PendingDecisionCount.ToString("N0", CultureInfo.CurrentCulture) +
                " pending review decision(s)?\n\nThey have not been saved to a new Review workbook.");
            if (abandon) _decisionReview.DiscardPending();
            return abandon;
        }

        private void ArrangeHeader()
        {
            int width = _header.ClientSize.Width;
            if (width <= 0) return;
            _alphaBadge.Left = Math.Max(148, width - _alphaBadge.Width - 14);
            _headerTitle.Text = width < 230 ? "Compare revisions" : "Comparison workspace";
            _headerTitle.Width = Math.Max(120, width - 32);
            _headerSubtitle.Width = Math.Max(120, width - 34);
        }

        private void ArrangeActions()
        {
            int width = _actions.ClientSize.Width;
            if (width <= 0) return;
            ArrangeSectionHeading(_actions, width);
            int innerWidth = Math.Max(120, width - 28);
            if (width >= 500)
            {
                int cardWidth = Math.Max(170, (innerWidth - 8) / 2);
                _newComparison.SetBounds(14, 48, cardWidth, 78);
                _openReview.SetBounds(22 + cardWidth, 48, innerWidth - cardWidth - 8, 78);
            }
            else
            {
                _newComparison.SetBounds(14, 48, innerWidth, 78);
                _openReview.SetBounds(14, 136, innerWidth, 78);
            }
        }

        private void ArrangeLoadedReview()
        {
            int width = _loadedReview.ClientSize.Width;
            if (width <= 0) return;
            ArrangeSectionHeading(_loadedReview, width);
            int innerWidth = Math.Max(120, width - 28);
            _integrityBadge.SetBounds(14, 66, Math.Min(132, innerWidth), 22);
            _loadedFileName.SetBounds(14, 94, innerWidth, 25);
            _loadedPath.SetBounds(14, 119, innerWidth, 31);

            if (width >= 440)
            {
                int gap = 12;
                int leftWidth = Math.Max(120, (innerWidth - gap) / 2);
                int rightLeft = 14 + leftWidth + gap;
                int rightWidth = Math.Max(100, innerWidth - leftWidth - gap);
                _loadedProject.SetBounds(14, 157, leftWidth, 21);
                _loadedFileMetadata.SetBounds(14, 180, leftWidth, 21);
                _loadedStateMetadata.SetBounds(rightLeft, 157, rightWidth, 21);
                _loadedProjectionMetadata.SetBounds(rightLeft, 180, rightWidth, 21);
                int contentY = 208;
                if (_loadedDecisionLedger.Visible)
                {
                    int decisionHeight = ArrangeDecisionLedger(innerWidth);
                    _loadedDecisionLedger.SetBounds(14, contentY, innerWidth, decisionHeight);
                    contentY += decisionHeight + 7;
                }
                int limitationHeight = Math.Max(
                    43,
                    MeasureWrappedText(_loadedLimitations.Text, _loadedLimitations.Font, innerWidth) + 2);
                _loadedLimitations.SetBounds(14, contentY, innerWidth, limitationHeight);
                int buttonY = contentY + limitationHeight + 9;
                int chooseWidth = 119;
                int openWidth = Math.Max(145, Math.Min(190, innerWidth - chooseWidth - 8));
                _openLoadedReview.SetBounds(14, buttonY, openWidth, 34);
                _chooseAnotherReview.SetBounds(22 + openWidth, buttonY, chooseWidth, 34);
                _loadedReview.Height = buttonY + 48;
            }
            else
            {
                _loadedProject.SetBounds(14, 157, innerWidth, 21);
                _loadedFileMetadata.SetBounds(14, 179, innerWidth, 21);
                _loadedStateMetadata.SetBounds(14, 201, innerWidth, 21);
                _loadedProjectionMetadata.SetBounds(14, 223, innerWidth, 21);
                int contentY = 250;
                if (_loadedDecisionLedger.Visible)
                {
                    int decisionHeight = ArrangeDecisionLedger(innerWidth);
                    _loadedDecisionLedger.SetBounds(14, contentY, innerWidth, decisionHeight);
                    contentY += decisionHeight + 7;
                }
                int limitationHeight = Math.Max(
                    47,
                    MeasureWrappedText(_loadedLimitations.Text, _loadedLimitations.Font, innerWidth) + 2);
                _loadedLimitations.SetBounds(14, contentY, innerWidth, limitationHeight);
                int buttonY = contentY + limitationHeight + 8;
                if (width >= 330)
                {
                    int gap = 8;
                    int buttonWidth = Math.Max(110, (innerWidth - gap) / 2);
                    _openLoadedReview.SetBounds(14, buttonY, buttonWidth, 34);
                    _chooseAnotherReview.SetBounds(
                        14 + buttonWidth + gap,
                        buttonY,
                        Math.Max(100, innerWidth - buttonWidth - gap),
                        34);
                    _loadedReview.Height = buttonY + 48;
                }
                else
                {
                    _openLoadedReview.SetBounds(14, buttonY, innerWidth, 34);
                    _chooseAnotherReview.SetBounds(14, buttonY + 40, innerWidth, 34);
                    _loadedReview.Height = buttonY + 88;
                }
            }
        }

        private int ArrangeDecisionLedger(int width)
        {
            int textWidth = Math.Max(80, width - 20);
            _loadedDecisionLedgerTitle.SetBounds(10, 7, textWidth, 19);
            int metadataHeight = MeasureWrappedText(
                _loadedDecisionLedgerMetadata.Text,
                _loadedDecisionLedgerMetadata.Font,
                textWidth);
            _loadedDecisionLedgerMetadata.SetBounds(10, 27, textWidth, metadataHeight + 2);
            return Math.Max(51, metadataHeight + 35);
        }

        private void ArrangeWorkflow()
        {
            int width = _workflow.ClientSize.Width;
            if (width <= 0) return;
            ArrangeSectionHeading(_workflow, width);
            int innerWidth = Math.Max(120, width - 28);
            int y = 66;
            int rowHeight = width < 360 ? 52 : 43;
            foreach (WorkflowStageRow row in _stages)
            {
                row.SetBounds(14, y, innerWidth, rowHeight);
                y += rowHeight + 4;
            }
            _workflow.Height = y + 9;
        }

        private void ArrangeAlphaSummary()
        {
            int width = _alpha.ClientSize.Width;
            if (width <= 0) return;
            ArrangeSectionHeading(_alpha, width);
            int innerWidth = Math.Max(120, width - 28);
            int y = 66;
            foreach (AlphaInfoRow row in _alphaRows)
            {
                row.Width = innerWidth;
                int height = row.ArrangeAndMeasure(innerWidth);
                row.SetBounds(14, y, innerWidth, height);
                y += height + 3;
            }
            _alpha.Height = y + 10;
        }

        private static void ArrangeSectionHeading(Control parent, int width)
        {
            Control heading = parent.Controls["SectionHeading"];
            Control description = parent.Controls["SectionDescription"];
            if (heading != null) heading.Width = Math.Max(100, width - 32);
            if (description != null) description.Width = Math.Max(100, width - 32);
        }

        private class SmoothLabel : Label
        {
            internal SmoothLabel()
            {
                UseCompatibleTextRendering = true;
                UseMnemonic = false;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                base.OnPaint(e);
            }
        }

        private sealed class RoundedSurface : Panel
        {
            internal Color FillColor { get; set; } = Color.White;
            internal Color BorderColor { get; set; } = Border;

            internal RoundedSurface()
            {
                BackColor = Color.Transparent;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = Rounded(bounds, 11))
                using (var fill = new SolidBrush(FillColor))
                using (var border = new Pen(BorderColor))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }
            }
        }

        private sealed class GradientSurface : Panel
        {
            internal Color StartColor { get; set; }
            internal Color EndColor { get; set; }

            internal GradientSurface()
            {
                BackColor = Color.Transparent;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = Rounded(bounds, 12))
                using (var fill = new LinearGradientBrush(bounds, StartColor, EndColor, 12f))
                    e.Graphics.FillPath(fill, path);
            }
        }

        private sealed class NavigationButton : Button
        {
            private bool _hovered;

            internal NavigationButton()
            {
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                BackColor = Color.Transparent;
                ForeColor = Color.White;
                Cursor = Cursors.Hand;
                Font = new Font("Segoe UI Semibold", 8.3f, FontStyle.Bold);
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);
                MouseEnter += (_, __) => { _hovered = true; Invalidate(); };
                MouseLeave += (_, __) => { _hovered = false; Invalidate(); };
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = Rounded(bounds, 8))
                using (var fill = new SolidBrush(Color.FromArgb(_hovered ? 52 : 30, Color.White)))
                using (var border = new Pen(Color.FromArgb(_hovered ? 105 : 65, Color.White)))
                using (var text = new SolidBrush(Color.White))
                using (var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap,
                })
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                    e.Graphics.DrawString(Text, Font, text, bounds, format);
                }
                if (Focused) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -4, -4));
            }
        }

        private sealed class WorkspaceButton : Button
        {
            private readonly bool _primary;
            private bool _hovered;

            internal WorkspaceButton(bool primary)
            {
                _primary = primary;
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                BackColor = Color.Transparent;
                ForeColor = primary ? Color.White : ComparisonAccent;
                Cursor = Cursors.Hand;
                Font = new Font("Segoe UI Semibold", 8.2f, FontStyle.Bold);
                UseMnemonic = false;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);
                MouseEnter += (_, __) => { _hovered = true; Invalidate(); };
                MouseLeave += (_, __) => { _hovered = false; Invalidate(); };
                GotFocus += (_, __) => Invalidate();
                LostFocus += (_, __) => Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                Color fillColor = _primary
                    ? (_hovered ? Color.FromArgb(86, 81, 230) : Accent)
                    : (_hovered ? Color.FromArgb(246, 245, 253) : Color.White);
                Color borderColor = _primary
                    ? Color.FromArgb(65, 64, 194)
                    : (_hovered ? Color.FromArgb(150, 143, 214) : Color.FromArgb(199, 195, 225));
                using (GraphicsPath path = Rounded(bounds, 8))
                using (var fill = new SolidBrush(fillColor))
                using (var border = new Pen(borderColor))
                using (var text = new SolidBrush(_primary ? Color.White : ComparisonAccent))
                using (var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap,
                })
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                    e.Graphics.DrawString(Text, Font, text, bounds, format);
                }
                if (Focused) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -4, -4));
            }
        }

        private sealed class PillLabel : SmoothLabel
        {
            internal Color FillColor { get; set; }
            internal Color BorderColor { get; set; }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = Rounded(bounds, Math.Max(2, bounds.Height / 2)))
                using (var fill = new SolidBrush(FillColor))
                using (var border = new Pen(BorderColor))
                using (var text = new SolidBrush(ForeColor))
                using (var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap,
                })
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                    e.Graphics.DrawString(Text, Font, text, bounds, format);
                }
            }
        }

        private sealed class ActionCard : Panel
        {
            private readonly bool _primary;
            private readonly SmoothLabel _icon;
            private readonly SmoothLabel _title;
            private readonly SmoothLabel _description;
            private readonly SmoothLabel _arrow;
            private bool _hovered;

            internal ActionCard(string icon, string title, string description, bool primary)
            {
                _primary = primary;
                BackColor = Color.Transparent;
                Cursor = Cursors.Hand;
                TabStop = true;
                AccessibleRole = AccessibleRole.PushButton;
                AccessibleDefaultActionDescription = "Open";
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor | ControlStyles.Selectable, true);

                _icon = new SmoothLabel
                {
                    Text = icon,
                    Location = new Point(11, 16),
                    Size = new Size(43, 43),
                    BackColor = Color.Transparent,
                    ForeColor = primary ? Color.White : ComparisonAccent,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI Semibold", icon.Length > 3 ? 7f : 8.5f, FontStyle.Bold),
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_icon);

                _title = new SmoothLabel
                {
                    Text = title,
                    Location = new Point(64, 10),
                    Height = 24,
                    BackColor = Color.Transparent,
                    ForeColor = primary ? Color.White : Ink,
                    Font = new Font("Segoe UI Semibold", 9.2f, FontStyle.Bold),
                    AutoEllipsis = true,
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_title);

                _description = new SmoothLabel
                {
                    Text = description,
                    Location = new Point(64, 34),
                    Height = 35,
                    BackColor = Color.Transparent,
                    ForeColor = primary ? Color.FromArgb(224, 226, 255) : Muted,
                    Font = new Font("Segoe UI", 8f),
                    AutoEllipsis = true,
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_description);

                _arrow = new SmoothLabel
                {
                    Text = "›",
                    Size = new Size(20, 30),
                    Location = new Point(0, 23),
                    BackColor = Color.Transparent,
                    ForeColor = primary ? Color.White : ComparisonAccent,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI Symbol", 16f),
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_arrow);

                Bind(this);
                KeyDown += (_, e) =>
                {
                    if (e.KeyCode != Keys.Enter && e.KeyCode != Keys.Space) return;
                    Activated?.Invoke(this, EventArgs.Empty);
                    e.Handled = e.SuppressKeyPress = true;
                };
                GotFocus += (_, __) => Invalidate();
                LostFocus += (_, __) => Invalidate();
                Resize += (_, __) => Arrange();
                Arrange();
            }

            internal event EventHandler Activated;

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                Color fillColor = _primary
                    ? (_hovered ? Color.FromArgb(91, 83, 235) : Accent)
                    : (_hovered ? Color.FromArgb(249, 249, 255) : Color.White);
                Color borderColor = _primary
                    ? Color.FromArgb(72, 68, 197)
                    : (_hovered ? Color.FromArgb(168, 165, 235) : Border);
                using (GraphicsPath path = Rounded(bounds, 9))
                using (var fill = new SolidBrush(fillColor))
                using (var border = new Pen(borderColor, _hovered ? 1.35f : 1f))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }

                if (_icon.Visible)
                {
                    Rectangle iconBounds = new Rectangle(10, 15, 44, 44);
                    using (GraphicsPath iconPath = Rounded(iconBounds, 9))
                    using (var iconFill = new SolidBrush(_primary
                               ? Color.FromArgb(42, Color.White)
                               : Color.FromArgb(24, ComparisonAccent)))
                    using (var iconBorder = new Pen(_primary
                               ? Color.FromArgb(70, Color.White)
                               : Color.FromArgb(55, ComparisonAccent)))
                    {
                        e.Graphics.FillPath(iconFill, iconPath);
                        e.Graphics.DrawPath(iconBorder, iconPath);
                    }
                }
                if (Focused) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -4, -4));
            }

            private void Arrange()
            {
                bool compact = Width < 250;
                bool ultraCompact = Width < 210;
                _icon.Visible = !ultraCompact;
                _arrow.Visible = !compact;
                _title.Left = ultraCompact ? 12 : 64;
                _description.Left = _title.Left;
                _arrow.Left = Math.Max(65, Width - _arrow.Width - 8);
                int textRight = compact ? Width - 8 : _arrow.Left - 5;
                int textWidth = Math.Max(45, textRight - _title.Left);
                _title.Width = textWidth;
                _description.Width = textWidth;
            }

            private void Bind(Control control)
            {
                control.MouseEnter += (_, __) => SetHovered(true);
                control.MouseLeave += (_, __) => SetHovered(ClientRectangle.Contains(PointToClient(Cursor.Position)));
                control.MouseDown += (_, __) => Focus();
                control.Click += (_, __) => Activated?.Invoke(this, EventArgs.Empty);
                foreach (Control child in control.Controls) Bind(child);
            }

            private void SetHovered(bool value)
            {
                if (_hovered == value) return;
                _hovered = value;
                Invalidate();
            }
        }

        private sealed class WorkflowStageRow : Panel
        {
            private readonly bool _implemented;
            private readonly SmoothLabel _title;
            private readonly SmoothLabel _description;
            private readonly PillLabel _status;

            internal WorkflowStageRow(string title, string description, bool implemented)
            {
                _implemented = implemented;
                BackColor = Color.Transparent;
                AccessibleRole = AccessibleRole.StaticText;
                AccessibleName = title;
                AccessibleDescription = (implemented ? "Implemented. " : "Not yet available. ") + description;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);

                var glyph = new SmoothLabel
                {
                    Text = implemented ? "✓" : "○",
                    Location = new Point(9, 8),
                    Size = new Size(26, 26),
                    BackColor = Color.Transparent,
                    ForeColor = implemented ? Ready : DisabledInk,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI Symbol", implemented ? 11f : 9.5f, FontStyle.Bold),
                };
                Controls.Add(glyph);

                _title = new SmoothLabel
                {
                    Text = title,
                    Location = new Point(42, 4),
                    Height = 19,
                    BackColor = Color.Transparent,
                    ForeColor = implemented ? Ink : DisabledInk,
                    Font = new Font("Segoe UI Semibold", 8.3f, FontStyle.Bold),
                    AutoEllipsis = true,
                };
                Controls.Add(_title);

                _description = new SmoothLabel
                {
                    Text = description,
                    Location = new Point(42, 22),
                    Height = 17,
                    BackColor = Color.Transparent,
                    ForeColor = implemented ? Muted : Color.FromArgb(157, 164, 180),
                    Font = new Font("Segoe UI", 7.6f),
                    AutoEllipsis = true,
                };
                Controls.Add(_description);

                _status = new PillLabel
                {
                    Text = implemented ? "READY" : "NOT AVAILABLE",
                    Size = new Size(implemented ? 54 : 88, 18),
                    Location = new Point(0, 6),
                    FillColor = implemented
                        ? Color.FromArgb(229, 247, 238)
                        : Color.FromArgb(239, 241, 245),
                    BorderColor = implemented
                        ? Color.FromArgb(180, 229, 205)
                        : Color.FromArgb(218, 222, 231),
                    ForeColor = implemented ? Ready : DisabledInk,
                    Font = new Font("Segoe UI Semibold", 6.2f, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleCenter,
                };
                Controls.Add(_status);
                Resize += (_, __) => Arrange();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = Rounded(bounds, 7))
                using (var fill = new SolidBrush(_implemented
                           ? Color.FromArgb(250, 253, 252)
                           : Color.FromArgb(247, 248, 250)))
                using (var border = new Pen(_implemented
                           ? Color.FromArgb(222, 235, 228)
                           : Color.FromArgb(231, 233, 238)))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }
            }

            private void Arrange()
            {
                _status.Left = Math.Max(42, Width - _status.Width - 8);
                if (Width < 332)
                {
                    _status.Top = Math.Max(24, Height - _status.Height - 5);
                    _title.Width = Math.Max(40, Width - _title.Left - 8);
                    _description.Width = Math.Max(35, _status.Left - _description.Left - 5);
                }
                else
                {
                    _status.Top = 6;
                    _title.Width = Math.Max(40, _status.Left - _title.Left - 6);
                    _description.Width = Math.Max(55, Width - _description.Left - 10);
                }
            }
        }

        private sealed class AlphaInfoRow : Panel
        {
            private readonly SmoothLabel _label;
            private readonly SmoothLabel _value;

            internal AlphaInfoRow(string label, string value)
            {
                BackColor = Color.Transparent;
                _label = new SmoothLabel
                {
                    Text = label,
                    BackColor = Color.Transparent,
                    ForeColor = ComparisonAccent,
                    Font = new Font("Segoe UI Semibold", 6.8f, FontStyle.Bold),
                    AutoEllipsis = true,
                };
                Controls.Add(_label);
                _value = new SmoothLabel
                {
                    Text = value,
                    BackColor = Color.Transparent,
                    ForeColor = Muted,
                    Font = new Font("Segoe UI", 7.9f),
                };
                Controls.Add(_value);
            }

            internal int ArrangeAndMeasure(int width)
            {
                bool stacked = width < 340;
                if (stacked)
                {
                    _label.SetBounds(8, 6, Math.Max(80, width - 16), 17);
                    int valueWidth = Math.Max(80, width - 16);
                    int textHeight = MeasureWrapped(_value.Text, _value.Font, valueWidth);
                    _value.SetBounds(8, 24, valueWidth, textHeight + 3);
                    return Math.Max(55, 31 + textHeight);
                }
                int keyWidth = 80;
                _label.SetBounds(8, 9, keyWidth, 18);
                int valueLeft = keyWidth + 16;
                int sideValueWidth = Math.Max(90, width - valueLeft - 8);
                int valueHeight = MeasureWrapped(_value.Text, _value.Font, sideValueWidth);
                _value.SetBounds(valueLeft, 7, sideValueWidth, valueHeight + 3);
                return Math.Max(44, valueHeight + 15);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(Color.FromArgb(235, 237, 242)))
                    e.Graphics.DrawLine(pen, 7, Height - 1, Width - 7, Height - 1);
            }

            private static int MeasureWrapped(string text, Font font, int width)
            {
                Size measured = TextRenderer.MeasureText(
                    text ?? "",
                    font,
                    new Size(Math.Max(1, width), int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
                return Math.Max(font.Height, measured.Height);
            }
        }

        private static GraphicsPath Rounded(Rectangle bounds, int radius)
        {
            int safeRadius = Math.Max(1, Math.Min(radius, Math.Min(bounds.Width, bounds.Height) / 2));
            int diameter = safeRadius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    internal sealed class ReviewWorkbookOpenRequestedEventArgs : EventArgs
    {
        internal ReviewWorkbookOpenRequestedEventArgs(ReviewWorkbookHydrationSession session)
        {
            Session = session ?? throw new ArgumentNullException(nameof(session));
        }

        internal ReviewWorkbookHydrationSession Session { get; }
        internal ReviewWorkbookValidationSummary Summary => Session.Summary;
    }
}
