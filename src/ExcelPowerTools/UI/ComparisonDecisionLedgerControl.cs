using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ExcelPowerTools.Core.Comparison.Decisions;
using ExcelPowerTools.Core.Comparison.Model;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Virtualized, copy-on-write review surface. Right-click actions update only an in-memory immutable
    /// generation; Save publishes the complete chained request batch into one newly named Review workbook.
    /// </summary>
    internal sealed class ComparisonDecisionLedgerControl : Panel
    {
        private static readonly Color Ink = Color.FromArgb(28, 35, 56);
        private static readonly Color Muted = Color.FromArgb(101, 112, 136);
        private static readonly Color Border = Color.FromArgb(215, 219, 231);
        private static readonly Color Accent = Color.FromArgb(104, 70, 198);
        private static readonly Color Ready = Color.FromArgb(31, 139, 91);
        private static readonly Color Warning = Color.FromArgb(153, 91, 18);

        private readonly SmoothDecisionLabel _title = new SmoothDecisionLabel();
        private readonly SmoothDecisionLabel _description = new SmoothDecisionLabel();
        private readonly SmoothDecisionLabel _filterLabel = new SmoothDecisionLabel();
        private readonly SmoothDecisionLabel _reviewerLabel = new SmoothDecisionLabel();
        private readonly SmoothDecisionLabel _noteLabel = new SmoothDecisionLabel();
        private readonly TextBox _filter = new TextBox();
        private readonly TextBox _reviewer = new TextBox();
        private readonly TextBox _note = new TextBox();
        private readonly ListView _ledger = new ListView();
        private readonly SmoothDecisionLabel _status = new SmoothDecisionLabel();
        private readonly Button _save = new Button();
        private readonly Button _discard = new Button();
        private readonly string _officeReviewerName;
        private readonly string _addInBuild;
        private readonly List<ReviewDecisionRequest> _pending = new List<ReviewDecisionRequest>();
        private List<ReviewDecisionWorkspaceRow> _visibleRows = new List<ReviewDecisionWorkspaceRow>();
        private ReviewWorkbookHydrationSession _session;
        private ReviewDecisionGeneration _baseGeneration;
        private ReviewDecisionGeneration _workingGeneration;
        private ReviewDecisionWorkspaceProjection _projection;

        internal ComparisonDecisionLedgerControl()
        {
            Height = 500;
            Margin = new Padding(0, 0, 0, 10);
            BackColor = Color.White;
            BorderStyle = BorderStyle.FixedSingle;
            Font = new Font("Segoe UI", 8.5f);
            TabStop = false;
            AccessibleRole = AccessibleRole.Pane;
            AccessibleName = "Review decisions";
            _officeReviewerName = ReadOfficeReviewerName();
            _addInBuild = ReadAddInBuild();

            _title.Text = "REVIEW DECISIONS";
            _title.ForeColor = Accent;
            _title.Font = new Font("Segoe UI Semibold", 8f, FontStyle.Bold);
            Controls.Add(_title);

            _description.Text =
                "Select one or many cells, then right-click for relevant choices. Changes remain pending until saved as a new Review revision.";
            _description.ForeColor = Muted;
            _description.Font = new Font("Segoe UI", 8f);
            Controls.Add(_description);

            ConfigureField(_filterLabel, _filter, "FILTER", "");
            _filter.TextChanged += (_, __) => ApplyFilter();
            ConfigureField(_reviewerLabel, _reviewer, "REVIEWER", _officeReviewerName);
            _reviewer.MaxLength = 512;
            ConfigureField(_noteLabel, _note, "AUDIT NOTE", "");
            _note.MaxLength = 4096;

            _ledger.View = View.Details;
            _ledger.FullRowSelect = true;
            _ledger.MultiSelect = true;
            _ledger.HideSelection = false;
            _ledger.GridLines = true;
            _ledger.VirtualMode = true;
            _ledger.VirtualListSize = 0;
            _ledger.BorderStyle = BorderStyle.FixedSingle;
            _ledger.Font = new Font("Segoe UI", 8f);
            _ledger.Columns.Add("Status", 105);
            _ledger.Columns.Add("Aspect", 70);
            _ledger.Columns.Add("Cell", 105);
            _ledger.Columns.Add("Master", 180);
            _ledger.Columns.Add("Latest candidate", 210);
            _ledger.RetrieveVirtualItem += RetrieveVirtualItem;
            _ledger.SelectedIndexChanged += (_, __) => UpdateStatus();
            _ledger.MouseUp += LedgerMouseUp;
            _ledger.KeyDown += LedgerKeyDown;
            Controls.Add(_ledger);

            _status.ForeColor = Muted;
            _status.Font = new Font("Segoe UI", 7.7f);
            _status.AutoEllipsis = true;
            Controls.Add(_status);

            ConfigureButton(_save, primary: true);
            _save.Text = "Save new Review revision";
            _save.Click += (_, __) => RequestSave();
            Controls.Add(_save);

            ConfigureButton(_discard, primary: false);
            _discard.Text = "Discard pending";
            _discard.Click += (_, __) => DiscardPending();
            Controls.Add(_discard);

            Resize += (_, __) => ArrangeControls();
            ArrangeControls();
            UpdateStatus();
        }

        internal event EventHandler<ComparisonDecisionCommitRequestedEventArgs> CommitRequested;

        internal int PendingDecisionCount => _pending.Count;
        internal bool HasPendingDecisions => _pending.Count != 0;
        internal ReviewWorkbookHydrationSession Session => _session;

        internal void FocusPageStart()
        {
            if (_filter.CanFocus) _filter.Focus();
        }

        internal void LoadSession(ReviewWorkbookHydrationSession session)
        {
            if (session == null) throw new ArgumentNullException(nameof(session));
            if (!session.ReadResult.HasDecisionState)
                throw new InvalidOperationException(
                    "This Review workbook has no editable decision ledger.");
            _session = session;
            _baseGeneration = session.ReadResult.DecisionGeneration;
            _workingGeneration = _baseGeneration;
            _pending.Clear();
            RefreshProjection();
        }

        internal void ClearSession()
        {
            _session = null;
            _baseGeneration = null;
            _workingGeneration = null;
            _projection = null;
            _pending.Clear();
            _visibleRows.Clear();
            _ledger.VirtualListSize = 0;
            _ledger.Invalidate();
            UpdateStatus();
        }

        internal void DiscardPending()
        {
            if (_baseGeneration == null) return;
            _workingGeneration = _baseGeneration;
            _pending.Clear();
            RefreshProjection();
        }

        private static void ConfigureField(
            Control label,
            TextBox field,
            string labelText,
            string initialValue)
        {
            label.Text = labelText;
            label.ForeColor = Muted;
            label.Font = new Font("Segoe UI Semibold", 6.8f, FontStyle.Bold);
            field.Text = initialValue ?? "";
            field.BorderStyle = BorderStyle.FixedSingle;
            field.Font = new Font("Segoe UI", 8.2f);
            field.BackColor = Color.White;
        }

        private void ConfigureField(
            SmoothDecisionLabel label,
            TextBox field,
            string labelText,
            string initialValue)
        {
            ConfigureField((Control)label, field, labelText, initialValue);
            Controls.Add(label);
            Controls.Add(field);
        }

        private static void ConfigureButton(Button button, bool primary)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = primary
                ? Color.FromArgb(79, 55, 173)
                : Border;
            button.BackColor = primary ? Accent : Color.White;
            button.ForeColor = primary ? Color.White : Ink;
            button.Font = new Font("Segoe UI Semibold", 8f, FontStyle.Bold);
            button.UseVisualStyleBackColor = false;
            button.Cursor = Cursors.Hand;
        }

        private void ArrangeControls()
        {
            int width = ClientSize.Width;
            if (width <= 0) return;
            int innerWidth = Math.Max(180, width - 24);
            _title.SetBounds(12, 9, innerWidth, 18);
            int descriptionHeight = Math.Max(
                36,
                TextRenderer.MeasureText(
                    _description.Text,
                    _description.Font,
                    new Size(innerWidth, int.MaxValue),
                    TextFormatFlags.WordBreak | TextFormatFlags.NoPadding).Height + 2);
            _description.SetBounds(12, 28, innerWidth, descriptionHeight);

            int fieldsTop = _description.Bottom + 4;
            if (width >= 690)
            {
                int gap = 8;
                int filterWidth = Math.Max(150, innerWidth * 35 / 100);
                int reviewerWidth = Math.Max(130, innerWidth * 24 / 100);
                int noteLeft = 12 + filterWidth + gap + reviewerWidth + gap;
                int noteWidth = Math.Max(150, 12 + innerWidth - noteLeft);
                ArrangeField(_filterLabel, _filter, 12, fieldsTop, filterWidth);
                ArrangeField(_reviewerLabel, _reviewer, 12 + filterWidth + gap, fieldsTop, reviewerWidth);
                ArrangeField(_noteLabel, _note, noteLeft, fieldsTop, noteWidth);
                int ledgerTop = fieldsTop + 44;
                _ledger.SetBounds(12, ledgerTop, innerWidth, Math.Max(210, Height - ledgerTop - 59));
            }
            else if (width >= 330)
            {
                int gap = 8;
                int leftWidth = (innerWidth - gap) / 2;
                ArrangeField(_filterLabel, _filter, 12, fieldsTop, leftWidth);
                ArrangeField(
                    _reviewerLabel,
                    _reviewer,
                    12 + leftWidth + gap,
                    fieldsTop,
                    innerWidth - leftWidth - gap);
                int noteTop = fieldsTop + 44;
                ArrangeField(_noteLabel, _note, 12, noteTop, innerWidth);
                int ledgerTop = noteTop + 44;
                _ledger.SetBounds(12, ledgerTop, innerWidth, Math.Max(170, Height - ledgerTop - 79));
            }
            else
            {
                ArrangeField(_filterLabel, _filter, 12, fieldsTop, innerWidth);
                int reviewerTop = fieldsTop + 44;
                ArrangeField(_reviewerLabel, _reviewer, 12, reviewerTop, innerWidth);
                int noteTop = reviewerTop + 44;
                ArrangeField(_noteLabel, _note, 12, noteTop, innerWidth);
                int ledgerTop = noteTop + 44;
                _ledger.SetBounds(12, ledgerTop, innerWidth, Math.Max(130, Height - ledgerTop - 79));
            }

            int footerTop = _ledger.Bottom + 7;
            if (innerWidth >= 500)
            {
                int discardWidth = 108;
                int saveWidth = 178;
                _save.SetBounds(12, footerTop, saveWidth, 31);
                _discard.SetBounds(20 + saveWidth, footerTop, discardWidth, 31);
                int statusLeft = 28 + saveWidth + discardWidth;
                _status.SetBounds(
                    statusLeft,
                    footerTop + 2,
                    Math.Max(40, 12 + innerWidth - statusLeft),
                    27);
            }
            else
            {
                int gap = 8;
                int saveWidth = Math.Max(95, (innerWidth - gap) * 3 / 5);
                int discardWidth = Math.Max(78, innerWidth - saveWidth - gap);
                _save.SetBounds(12, footerTop, saveWidth, 31);
                _discard.SetBounds(12 + saveWidth + gap, footerTop, discardWidth, 31);
                _status.SetBounds(12, footerTop + 33, innerWidth, 20);
            }
        }

        private static void ArrangeField(
            Control label,
            TextBox field,
            int left,
            int top,
            int width)
        {
            label.SetBounds(left, top, width, 14);
            field.SetBounds(left, top + 15, width, 23);
        }

        private void RefreshProjection()
        {
            if (_workingGeneration == null)
            {
                _projection = null;
                _visibleRows.Clear();
                _ledger.VirtualListSize = 0;
            }
            else
            {
                _projection = ReviewDecisionWorkspaceProjection.Build(_workingGeneration);
                ApplyFilter();
            }
            UpdateStatus();
        }

        private void ApplyFilter()
        {
            if (_ledger.SelectedIndices.Count != 0) _ledger.SelectedIndices.Clear();
            if (_projection == null)
            {
                _visibleRows = new List<ReviewDecisionWorkspaceRow>();
            }
            else
            {
                string query = (_filter.Text ?? "").Trim();
                _visibleRows = query.Length == 0
                    ? new List<ReviewDecisionWorkspaceRow>(_projection.Rows)
                    : _projection.Rows.Where(item => Matches(item, query)).ToList();
            }
            _ledger.VirtualListSize = _visibleRows.Count;
            _ledger.Invalidate();
            UpdateStatus();
        }

        private static bool Matches(ReviewDecisionWorkspaceRow row, string query) =>
            Contains(row.Location, query) ||
            Contains(row.Aspect.ToString(), query) ||
            Contains(StatusText(row.Status), query) ||
            Contains(row.MasterContent, query) ||
            Contains(row.LatestCandidateContent, query) ||
            Contains(row.LatestCandidateRevisionLabel, query) ||
            Contains(row.LogicalSubjectKey, query);

        private static bool Contains(string value, string query) =>
            (value ?? "").IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;

        private void RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            if (e.ItemIndex < 0 || e.ItemIndex >= _visibleRows.Count)
            {
                e.Item = new ListViewItem("");
                return;
            }
            ReviewDecisionWorkspaceRow row = _visibleRows[e.ItemIndex];
            var item = new ListViewItem(StatusText(row.Status));
            item.SubItems.Add(AspectText(row.Aspect));
            item.SubItems.Add(row.Location);
            item.SubItems.Add(row.MasterContent);
            item.SubItems.Add(
                row.LatestCandidateRevisionLabel.Length == 0
                    ? row.LatestCandidateContent
                    : row.LatestCandidateRevisionLabel + " · " + row.LatestCandidateContent);
            item.ForeColor = StatusColor(row.Status);
            item.Tag = row;
            e.Item = item;
        }

        private void LedgerMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            ListViewItem hit = _ledger.GetItemAt(e.X, e.Y);
            if (hit != null && !hit.Selected)
            {
                _ledger.SelectedIndices.Clear();
                hit.Selected = true;
            }
            ShowDecisionMenu(new Point(e.X, e.Y));
        }

        private void LedgerKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Apps || (e.Shift && e.KeyCode == Keys.F10))
            {
                ShowDecisionMenu(new Point(24, 24));
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void ShowDecisionMenu(Point location)
        {
            List<ReviewDecisionWorkspaceRow> selected = SelectedRows();
            if (selected.Count == 0) return;
            var menu = new ContextMenuStrip
            {
                Font = new Font("Segoe UI", 8.5f),
                ShowImageMargin = false,
            };
            var heading = new ToolStripMenuItem(
                selected.Count == 1
                    ? selected[0].Location + " · " + AspectText(selected[0].Aspect)
                    : selected.Count.ToString("N0", CultureInfo.CurrentCulture) + " selected changes")
            {
                Enabled = false,
            };
            menu.Items.Add(heading);
            menu.Items.Add(new ToolStripSeparator());

            bool allSelectable = selected.All(item => item.CanSelectObservation);
            bool allHaveMaster = allSelectable && selected.All(item => item.Master != null);
            if (allHaveMaster)
            {
                menu.Items.Add(ActionItem(
                    selected.Count == 1 ? "Use master" : "Use master for selected",
                    () => QueueObservationStatus(
                        selected,
                        ReviewDecisionStatus.UseMaster,
                        row => row.Master)));
            }

            if (allSelectable)
            {
                ToolStripMenuItem candidates = BuildCommonCandidateMenu(selected);
                if (candidates.DropDownItems.Count > 0) menu.Items.Add(candidates);
            }

            if (allHaveMaster || allSelectable) menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(ActionItem(
                selected.Count == 1 ? "Mark equivalent" : "Mark selected equivalent",
                () => QueueStatus(selected, ReviewDecisionStatus.Equivalent)));
            menu.Items.Add(ActionItem(
                selected.Count == 1 ? "Needs investigation" : "Mark selected for investigation",
                () => QueueStatus(selected, ReviewDecisionStatus.NeedsInvestigation)));
            if (selected.All(item => !item.CanSelectObservation))
            {
                menu.Items.Add(ActionItem(
                    "Reviewed — no structural merge",
                    () => QueueStatus(selected, ReviewDecisionStatus.ReviewedNoStructuralMerge)));
            }
            if (selected.Any(item => item.Status != ReviewDecisionStatus.Unreviewed))
            {
                menu.Items.Add(ActionItem(
                    selected.Count == 1 ? "Clear decision" : "Clear selected decisions",
                    () => QueueStatus(selected, ReviewDecisionStatus.Unreviewed)));
            }

            if (selected.Count == 1 && selected[0].GroupIds.Count > 0)
            {
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add(BuildGroupMenu(selected[0]));
            }
            menu.Closed += (_, __) => menu.Dispose();
            menu.Show(_ledger, location);
        }

        private ToolStripMenuItem BuildCommonCandidateMenu(
            IReadOnlyList<ReviewDecisionWorkspaceRow> rows,
            string auditGroupId = "")
        {
            var menu = new ToolStripMenuItem(
                rows.Count == 1 ? "Use candidate" : "Use candidate revision for selected");
            IEnumerable<string> common = rows[0].Candidates.Select(item => item.RevisionId);
            foreach (ReviewDecisionWorkspaceRow row in rows.Skip(1))
            {
                var revisions = new HashSet<string>(
                    row.Candidates.Select(item => item.RevisionId),
                    StringComparer.Ordinal);
                common = common.Where(revisions.Contains).ToList();
            }
            foreach (string revisionId in common.Distinct(StringComparer.Ordinal))
            {
                ReviewDecisionObservationChoice example = rows[0].Candidates.First(
                    item => string.Equals(item.RevisionId, revisionId, StringComparison.Ordinal));
                string label = rows.Count == 1
                    ? example.RevisionLabel + " · " + example.Content
                    : example.RevisionLabel;
                menu.DropDownItems.Add(ActionItem(
                    Cap(label, 110),
                    () => QueueObservationStatus(
                        rows,
                        ReviewDecisionStatus.UseCandidate,
                        row => row.Candidates.First(item => string.Equals(
                            item.RevisionId,
                            revisionId,
                            StringComparison.Ordinal)),
                        auditGroupId)));
            }
            return menu;
        }

        private ToolStripMenuItem BuildGroupMenu(ReviewDecisionWorkspaceRow row)
        {
            var root = new ToolStripMenuItem("Apply to whole group");
            foreach (string groupId in row.GroupIds)
            {
                if (!_workingGeneration.Catalog.TryGetGroup(groupId, out ReviewDecisionGroup group)) continue;
                var groupItem = new ToolStripMenuItem(
                    "Group " + Prefix(group.GroupId) + " · " +
                    group.SubjectIds.Count.ToString("N0", CultureInfo.CurrentCulture) + " changes");
                List<ReviewDecisionWorkspaceRow> groupRows = group.SubjectIds
                    .Select(subjectId => _projection.TryGetRow(subjectId, out ReviewDecisionWorkspaceRow found)
                        ? found
                        : null)
                    .Where(item => item != null)
                    .ToList();
                if (groupRows.Count == group.SubjectIds.Count &&
                    groupRows.All(item => item.CanSelectObservation && item.Master != null))
                {
                    groupItem.DropDownItems.Add(ActionItem(
                        "Use master for group",
                        () => QueueObservationStatus(
                            groupRows,
                            ReviewDecisionStatus.UseMaster,
                            item => item.Master,
                            group.GroupId)));
                    ToolStripMenuItem candidates = BuildCommonCandidateMenu(
                        groupRows,
                        group.GroupId);
                    candidates.Text = "Use candidate revision for group";
                    if (candidates.DropDownItems.Count > 0)
                        groupItem.DropDownItems.Add(candidates);
                    groupItem.DropDownItems.Add(new ToolStripSeparator());
                }
                groupItem.DropDownItems.Add(ActionItem(
                    "Mark group equivalent",
                    () => QueueGroupStatus(group.GroupId, ReviewDecisionStatus.Equivalent)));
                groupItem.DropDownItems.Add(ActionItem(
                    "Group needs investigation",
                    () => QueueGroupStatus(group.GroupId, ReviewDecisionStatus.NeedsInvestigation)));
                if (groupRows.Count == group.SubjectIds.Count && groupRows.All(item => !item.CanSelectObservation))
                {
                    groupItem.DropDownItems.Add(ActionItem(
                        "Reviewed — no structural merge",
                        () => QueueGroupStatus(
                            group.GroupId,
                            ReviewDecisionStatus.ReviewedNoStructuralMerge)));
                }
                groupItem.DropDownItems.Add(ActionItem(
                    "Clear group decisions",
                    () => QueueGroupStatus(group.GroupId, ReviewDecisionStatus.Unreviewed)));
                root.DropDownItems.Add(groupItem);
            }
            return root;
        }

        private static ToolStripMenuItem ActionItem(string text, Action action)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (_, __) => action();
            return item;
        }

        private void QueueStatus(
            IReadOnlyList<ReviewDecisionWorkspaceRow> rows,
            ReviewDecisionStatus status)
        {
            TryQueue(() =>
            {
                string runId = NewRunId();
                ReviewerSnapshot reviewer = Reviewer();
                DateTimeOffset created = DateTimeOffset.UtcNow;
                string note = _note.Text.Trim();
                var pending = new List<ReviewDecisionRequest>();
                ReviewDecisionGeneration next = _workingGeneration;
                bool homogeneous = HaveHomogeneousPriorState(rows);
                if (rows.Count == 1 || homogeneous)
                {
                    ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                        next,
                        rows.Select(item => item.SubjectId),
                        status,
                        reviewer,
                        created,
                        runId: runId,
                        addInBuild: _addInBuild,
                        note: note);
                    next = ReviewDecisionEngine.Apply(next, request);
                    pending.Add(request);
                }
                else
                {
                    foreach (ReviewDecisionWorkspaceRow row in rows.OrderBy(
                                 item => item.SubjectId,
                                 StringComparer.Ordinal))
                    {
                        if (row.Status == status && row.SelectedObservationId.Length == 0) continue;
                        ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                            next,
                            new[] { row.SubjectId },
                            status,
                            reviewer,
                            created,
                            runId: runId,
                            addInBuild: _addInBuild,
                            note: note);
                        next = ReviewDecisionEngine.Apply(next, request);
                        pending.Add(request);
                    }
                }
                CommitWorkingBatch(next, pending);
            });
        }

        private void QueueGroupStatus(string groupId, ReviewDecisionStatus status)
        {
            TryQueue(() =>
            {
                if (!_workingGeneration.Catalog.TryGetGroup(groupId, out ReviewDecisionGroup group))
                    throw new InvalidOperationException(
                        "The selected review group is not in this decision generation.");
                var rows = group.SubjectIds.Select(subjectId =>
                    _projection.TryGetRow(subjectId, out ReviewDecisionWorkspaceRow row)
                        ? row
                        : throw new InvalidOperationException(
                            "The decision workspace is missing a group subject."))
                    .ToList();
                ReviewerSnapshot reviewer = Reviewer();
                DateTimeOffset created = DateTimeOffset.UtcNow;
                string runId = NewGroupRunId(groupId);
                string note = _note.Text.Trim();
                ReviewDecisionGeneration next = _workingGeneration;
                var pending = new List<ReviewDecisionRequest>();
                if (HaveHomogeneousPriorState(rows))
                {
                    ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForGroup(
                        next,
                        groupId,
                        status,
                        reviewer,
                        created,
                        runId: runId,
                        addInBuild: _addInBuild,
                        note: note);
                    next = ReviewDecisionEngine.Apply(next, request);
                    pending.Add(request);
                }
                else
                {
                    // One group event cannot truthfully name several superseded events. Keep the user's one
                    // action correlated by run ID while recording an exact event per prior-provenance state.
                    foreach (ReviewDecisionWorkspaceRow row in rows.OrderBy(
                                 item => item.SubjectId,
                                 StringComparer.Ordinal))
                    {
                        if (row.Status == status && row.SelectedObservationId.Length == 0) continue;
                        ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                            next,
                            new[] { row.SubjectId },
                            status,
                            reviewer,
                            created,
                            runId: runId,
                            addInBuild: _addInBuild,
                            note: note);
                        next = ReviewDecisionEngine.Apply(next, request);
                        pending.Add(request);
                    }
                }
                CommitWorkingBatch(next, pending);
            });
        }

        private void QueueObservationStatus(
            IReadOnlyList<ReviewDecisionWorkspaceRow> rows,
            ReviewDecisionStatus status,
            Func<ReviewDecisionWorkspaceRow, ReviewDecisionObservationChoice> choice,
            string auditGroupId = "")
        {
            TryQueue(() =>
            {
                string runId = auditGroupId.Length == 0
                    ? NewRunId()
                    : NewGroupRunId(auditGroupId);
                ReviewerSnapshot reviewer = Reviewer();
                DateTimeOffset created = DateTimeOffset.UtcNow;
                string note = _note.Text.Trim();
                var pending = new List<ReviewDecisionRequest>();
                ReviewDecisionGeneration next = _workingGeneration;
                foreach (ReviewDecisionWorkspaceRow row in rows.OrderBy(
                             item => item.SubjectId,
                             StringComparer.Ordinal))
                {
                    ReviewDecisionObservationChoice selected = choice(row) ??
                        throw new InvalidOperationException(
                            "The selected revision has no exact observation for " + row.Location + ".");
                    ReviewDecisionState current = next.GetSubjectState(row.SubjectId);
                    if (current.Status == status && string.Equals(
                            current.SelectedObservationId,
                            selected.ObservationId,
                            StringComparison.Ordinal))
                        continue;
                    ReviewDecisionRequest request = ReviewDecisionWorkspaceRequestBuilder.ForSubjects(
                        next,
                        new[] { row.SubjectId },
                        status,
                        reviewer,
                        created,
                        selected.ObservationId,
                        runId,
                        _addInBuild,
                        note);
                    next = ReviewDecisionEngine.Apply(next, request);
                    pending.Add(request);
                }
                CommitWorkingBatch(next, pending);
            });
        }

        private void TryQueue(Action action)
        {
            try
            {
                if (_workingGeneration == null)
                    throw new InvalidOperationException("Load a validated Review workbook first.");
                action();
            }
            catch (Exception exception)
            {
                Dialogs.Info(
                    "The decision was not queued. No file was changed.\n\n" +
                    exception.Message);
            }
        }

        private void CommitWorkingBatch(
            ReviewDecisionGeneration next,
            IEnumerable<ReviewDecisionRequest> requests)
        {
            var frozen = new List<ReviewDecisionRequest>(requests ?? Array.Empty<ReviewDecisionRequest>());
            if (frozen.Count == 0)
                throw new InvalidOperationException("The selected decision already is effective.");
            _workingGeneration = next ?? throw new ArgumentNullException(nameof(next));
            _pending.AddRange(frozen);
            RefreshProjection();
        }

        private void RequestSave()
        {
            if (_session == null || _pending.Count == 0) return;
            CommitRequested?.Invoke(
                this,
                new ComparisonDecisionCommitRequestedEventArgs(
                    _session,
                    _workingGeneration.GenerationId,
                    _pending));
        }

        private ReviewerSnapshot Reviewer()
        {
            string confirmed = _reviewer.Text.Trim();
            if (confirmed.Length == 0 && _officeReviewerName.Length == 0)
                return ReviewerSnapshot.Empty;
            var sources = new List<string>();
            if (confirmed.Length != 0) sources.Add("comparison-workspace-confirmed");
            if (_officeReviewerName.Length != 0) sources.Add("excel-application-username");
            return new ReviewerSnapshot(confirmed, _officeReviewerName, sources);
        }

        private List<ReviewDecisionWorkspaceRow> SelectedRows()
        {
            var result = new List<ReviewDecisionWorkspaceRow>();
            foreach (int index in _ledger.SelectedIndices.Cast<int>().OrderBy(item => item))
                if (index >= 0 && index < _visibleRows.Count)
                    result.Add(_visibleRows[index]);
            return result;
        }

        private static bool HaveHomogeneousPriorState(
            IReadOnlyList<ReviewDecisionWorkspaceRow> rows)
        {
            if (rows.Count <= 1) return true;
            ReviewDecisionWorkspaceRow first = rows[0];
            return rows.Skip(1).All(item =>
                item.Status == first.Status &&
                string.Equals(
                    item.SelectedObservationId,
                    first.SelectedObservationId,
                    StringComparison.Ordinal) &&
                string.Equals(item.EffectiveEventId, first.EffectiveEventId, StringComparison.Ordinal));
        }

        private void UpdateStatus()
        {
            int total = _projection?.Rows.Count ?? 0;
            int selected = _ledger.SelectedIndices.Count;
            _save.Enabled = _pending.Count > 0;
            _discard.Enabled = _pending.Count > 0;
            _save.Text = _pending.Count == 0
                ? "Save new Review revision"
                : "Save revision (" + _pending.Count.ToString("N0", CultureInfo.CurrentCulture) + ")";
            string visible = _visibleRows.Count == total
                ? total.ToString("N0", CultureInfo.CurrentCulture) + " changes"
                : _visibleRows.Count.ToString("N0", CultureInfo.CurrentCulture) + " of " +
                  total.ToString("N0", CultureInfo.CurrentCulture);
            _status.Text = visible +
                           (selected == 0 ? "" : " · " + selected.ToString("N0", CultureInfo.CurrentCulture) + " selected") +
                           (_pending.Count == 0
                               ? " · right-click to review"
                               : " · " + _pending.Count.ToString("N0", CultureInfo.CurrentCulture) + " pending");
            _status.ForeColor = _pending.Count == 0 ? Muted : Warning;
        }

        private static string StatusText(ReviewDecisionStatus status)
        {
            switch (status)
            {
                case ReviewDecisionStatus.UseMaster: return "Use master";
                case ReviewDecisionStatus.UseCandidate: return "Use candidate";
                case ReviewDecisionStatus.Equivalent: return "Equivalent";
                case ReviewDecisionStatus.NeedsInvestigation: return "Investigate";
                case ReviewDecisionStatus.ReviewedNoStructuralMerge: return "Reviewed / no merge";
                default: return "Unreviewed";
            }
        }

        private static string AspectText(ComparisonAspect aspect)
        {
            switch (aspect)
            {
                case ComparisonAspect.InteriorFill: return "Fill (context only)";
                case ComparisonAspect.CachedResult: return "Saved calculated value";
                default: return aspect.ToString();
            }
        }

        private static Color StatusColor(ReviewDecisionStatus status)
        {
            switch (status)
            {
                case ReviewDecisionStatus.UseCandidate: return Ready;
                case ReviewDecisionStatus.UseMaster: return Color.FromArgb(45, 91, 154);
                case ReviewDecisionStatus.NeedsInvestigation: return Warning;
                case ReviewDecisionStatus.Unreviewed: return Muted;
                default: return Ink;
            }
        }

        private static string NewRunId() => "review-ui-" + Guid.NewGuid().ToString("N");

        private static string NewGroupRunId(string groupId) =>
            "review-ui-group-" + (groupId ?? "") + "-" + Guid.NewGuid().ToString("N");

        private static string Prefix(string value) =>
            string.IsNullOrEmpty(value)
                ? "?"
                : value.Substring(0, Math.Min(10, value.Length)) + (value.Length > 10 ? "…" : "");

        private static string Cap(string value, int max) =>
            (value ?? "").Length <= max
                ? value ?? ""
                : value.Substring(0, max - 1) + "…";

        private static string ReadOfficeReviewerName()
        {
            try { return Convert.ToString(ExcelApp.App.UserName, CultureInfo.CurrentCulture)?.Trim() ?? ""; }
            catch { return ""; }
        }

        private static string ReadAddInBuild()
        {
            try
            {
                AssemblyName name = typeof(ComparisonDecisionLedgerControl).Assembly.GetName();
                return name.Name + " " + (name.Version?.ToString() ?? "unknown");
            }
            catch { return "ExcelPowerTools unknown"; }
        }

        private sealed class SmoothDecisionLabel : Label
        {
            internal SmoothDecisionLabel()
            {
                UseCompatibleTextRendering = true;
                UseMnemonic = false;
                BackColor = Color.Transparent;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                base.OnPaint(e);
            }
        }
    }

    internal sealed class ComparisonDecisionCommitRequestedEventArgs : EventArgs
    {
        internal ComparisonDecisionCommitRequestedEventArgs(
            ReviewWorkbookHydrationSession session,
            string expectedFinalGenerationId,
            IEnumerable<ReviewDecisionRequest> requests)
        {
            Session = session ?? throw new ArgumentNullException(nameof(session));
            if (string.IsNullOrWhiteSpace(expectedFinalGenerationId))
                throw new ArgumentException(
                    "An expected final decision generation is required.",
                    nameof(expectedFinalGenerationId));
            ExpectedFinalGenerationId = expectedFinalGenerationId;
            var frozen = new List<ReviewDecisionRequest>(
                requests ?? throw new ArgumentNullException(nameof(requests)));
            if (frozen.Count == 0 || frozen.Any(item => item == null))
                throw new ArgumentException(
                    "A decision commit requires one or more complete requests.",
                    nameof(requests));
            Requests = new ReadOnlyCollection<ReviewDecisionRequest>(frozen);
        }

        internal ReviewWorkbookHydrationSession Session { get; }
        internal string ExpectedFinalGenerationId { get; }
        internal IReadOnlyList<ReviewDecisionRequest> Requests { get; }
    }
}
