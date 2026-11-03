using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelPowerTools.Core.Comparison.RevisionBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>One-screen setup for the deliberately small keyed revision-comparison workflow.</summary>
    internal sealed class RevisionComparisonSetupForm : Form
    {
        private readonly RevisionComparisonInspection _inspection;
        private readonly ListBox _revisions = new ListBox();
        private readonly DataGridView _pairs = new DataGridView();
        private readonly ComboBox _transition = new ComboBox();
        private readonly TextBox _preview = new TextBox();
        private readonly Button _create = new Button();
        private bool _updatingGrid;

        internal RevisionComparisonSetupForm(RevisionComparisonInspection inspection)
        {
            _inspection = inspection ?? throw new ArgumentNullException(nameof(inspection));
            if (inspection.Sources.Count < 2)
                throw new ArgumentException("At least two revisions are required.", nameof(inspection));

            Text = "Compare table revisions";
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimumSize = new Size(980, 650);
            Size = new Size(1220, 780);
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.FromArgb(247, 248, 252);
            BuildLayout();
            LoadRevisions();
            RebuildDefaultPairs();
            GdiText.Apply(this);
        }

        internal RevisionComparisonPlan ResultPlan { get; private set; }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(14),
                ColumnCount = 1,
                RowCount = 4,
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 68F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 32F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            root.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "Order the revisions, pair one or more tables for every adjacent step, then confirm the key columns. " +
                       "The first revision is copied as the apply base; sources are never overwritten. " +
                       "Suggested keys are never accepted until Confirm is ticked.",
                ForeColor = Color.FromArgb(45, 50, 65),
                Margin = new Padding(0, 0, 0, 10),
            }, 0, 0);

            var work = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 270,
                FixedPanel = FixedPanel.Panel1,
                Margin = new Padding(0),
            };
            root.Controls.Add(work, 0, 1);
            BuildRevisionPanel(work.Panel1);
            BuildPairPanel(work.Panel2);

            _preview.Dock = DockStyle.Fill;
            _preview.Multiline = true;
            _preview.ReadOnly = true;
            _preview.ScrollBars = ScrollBars.Vertical;
            _preview.BackColor = Color.White;
            _preview.Text = "Preview keys to see row counts, blank keys and duplicate keys before creating the report.";
            _preview.Margin = new Padding(0, 10, 0, 10);
            root.Controls.Add(_preview, 0, 2);

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = new Padding(0),
            };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            _create.Text = "Create review";
            _create.AutoSize = true;
            _create.Click += (_, __) => CreatePlan();
            var previewButton = new Button { Text = "Preview keys", AutoSize = true };
            previewButton.Click += (_, __) => PreviewKeys();
            actions.Controls.Add(cancel);
            actions.Controls.Add(_create);
            actions.Controls.Add(previewButton);
            root.Controls.Add(actions, 0, 3);
            AcceptButton = _create;
            CancelButton = cancel;
        }

        private void BuildRevisionPanel(Control parent)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0, 0, 10, 0),
            };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            parent.Controls.Add(panel);
            panel.Controls.Add(new Label
            {
                Text = "Revision order — first revision is the apply base",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 6),
            }, 0, 0);
            _revisions.Dock = DockStyle.Fill;
            _revisions.IntegralHeight = false;
            _revisions.HorizontalScrollbar = true;
            panel.Controls.Add(_revisions, 0, 1);
            var move = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 6, 0, 0),
            };
            var up = new Button { Text = "Move up", AutoSize = true };
            var down = new Button { Text = "Move down", AutoSize = true };
            up.Click += (_, __) => MoveRevision(-1);
            down.Click += (_, __) => MoveRevision(1);
            move.Controls.Add(up);
            move.Controls.Add(down);
            panel.Controls.Add(move, 0, 2);
        }

        private void BuildPairPanel(Control parent)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
            };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            parent.Controls.Add(panel);
            panel.Controls.Add(new Label
            {
                Text = "Adjacent table pairs — separate composite key columns with semicolons",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 6),
            }, 0, 0);

            _pairs.Dock = DockStyle.Fill;
            _pairs.AllowUserToAddRows = false;
            _pairs.AllowUserToDeleteRows = false;
            _pairs.AllowUserToResizeRows = false;
            _pairs.RowHeadersVisible = false;
            _pairs.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _pairs.MultiSelect = false;
            _pairs.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Transition", ReadOnly = true, FillWeight = 110F });
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Before sheet", FillWeight = 90F });
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "After sheet", FillWeight = 90F });
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Before header", FillWeight = 55F });
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "After header", FillWeight = 55F });
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Before key(s)", FillWeight = 90F });
            _pairs.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "After key(s)", FillWeight = 90F });
            _pairs.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Confirm", FillWeight = 45F });
            _pairs.CurrentCellDirtyStateChanged += (_, __) =>
            {
                if (_pairs.IsCurrentCellDirty) _pairs.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _pairs.CellValueChanged += PairCellChanged;
            _pairs.DataError += (_, args) => { args.ThrowException = false; };
            panel.Controls.Add(_pairs, 0, 1);

            var edit = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 6, 0, 0),
            };
            _transition.DropDownStyle = ComboBoxStyle.DropDownList;
            _transition.Width = 250;
            var add = new Button { Text = "Add pair", AutoSize = true };
            var remove = new Button { Text = "Remove selected pair", AutoSize = true };
            add.Click += (_, __) => AddManualPair();
            remove.Click += (_, __) => RemoveSelectedPair();
            edit.Controls.Add(_transition);
            edit.Controls.Add(add);
            edit.Controls.Add(remove);
            panel.Controls.Add(edit, 0, 2);
        }

        private void LoadRevisions()
        {
            _revisions.Items.Clear();
            for (int index = 0; index < _inspection.Sources.Count; index++)
                _revisions.Items.Add(new RevisionItem(index, _inspection.Sources[index].Path));
            _revisions.SelectedIndex = 0;
        }

        private void MoveRevision(int delta)
        {
            int index = _revisions.SelectedIndex;
            int target = index + delta;
            if (index < 0 || target < 0 || target >= _revisions.Items.Count) return;
            object item = _revisions.Items[index];
            _revisions.Items.RemoveAt(index);
            _revisions.Items.Insert(target, item);
            _revisions.SelectedIndex = target;
            RebuildDefaultPairs();
        }

        private void RebuildDefaultPairs()
        {
            _updatingGrid = true;
            try
            {
                _pairs.Rows.Clear();
                _transition.Items.Clear();
                IReadOnlyList<int> order = OrderedIndexes();
                for (int position = 0; position < order.Count - 1; position++)
                {
                    int before = order[position];
                    int after = order[position + 1];
                    var transition = new TransitionItem(before, after, TransitionText(before, after));
                    _transition.Items.Add(transition);
                    string[] common = _inspection.Sources[before].Sheets.Select(item => item.Name)
                        .Intersect(_inspection.Sources[after].Sheets.Select(item => item.Name), StringComparer.Ordinal)
                        .OrderBy(item => item, StringComparer.Ordinal)
                        .ToArray();
                    if (common.Length == 0)
                    {
                        AddPairRow(
                            transition,
                            _inspection.Sources[before].Sheets.FirstOrDefault()?.Name ?? "",
                            _inspection.Sources[after].Sheets.FirstOrDefault()?.Name ?? "");
                    }
                    else
                    {
                        foreach (string sheetName in common) AddPairRow(transition, sheetName, sheetName);
                    }
                }
                if (_transition.Items.Count > 0) _transition.SelectedIndex = 0;
                InvalidatePreview("Pairs rebuilt for the new revision order. Review every row and tick Confirm.");
            }
            finally
            {
                _updatingGrid = false;
            }
        }

        private void AddManualPair()
        {
            if (!(_transition.SelectedItem is TransitionItem item)) return;
            _updatingGrid = true;
            try
            {
                AddPairRow(
                    item,
                    _inspection.Sources[item.BeforeIndex].Sheets.FirstOrDefault()?.Name ?? "",
                    _inspection.Sources[item.AfterIndex].Sheets.FirstOrDefault()?.Name ?? "");
            }
            finally
            {
                _updatingGrid = false;
            }
            InvalidatePreview();
        }

        private void AddPairRow(TransitionItem transition, string beforeSheet, string afterSheet)
        {
            int rowIndex = _pairs.Rows.Add();
            DataGridViewRow row = _pairs.Rows[rowIndex];
            row.Tag = new PairTag(transition.BeforeIndex, transition.AfterIndex);
            row.Cells[0].Value = transition.Text;
            row.Cells[1] = SheetCell(transition.BeforeIndex, beforeSheet);
            row.Cells[2] = SheetCell(transition.AfterIndex, afterSheet);
            FillSuggestions(row, true);
            FillSuggestions(row, false);
            row.Cells[7].Value = false;
        }

        private DataGridViewComboBoxCell SheetCell(int revisionIndex, string selected)
        {
            var cell = new DataGridViewComboBoxCell
            {
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                FlatStyle = FlatStyle.Flat,
            };
            foreach (RevisionComparisonSheetInspection sheet in _inspection.Sources[revisionIndex].Sheets)
                cell.Items.Add(sheet.Name);
            if (cell.Items.Contains(selected)) cell.Value = selected;
            return cell;
        }

        private void FillSuggestions(DataGridViewRow row, bool before)
        {
            PairTag tag = (PairTag)row.Tag;
            int sheetColumn = before ? 1 : 2;
            int headerColumn = before ? 3 : 4;
            int keyColumn = before ? 5 : 6;
            int revision = before ? tag.BeforeIndex : tag.AfterIndex;
            string sheetName = Convert.ToString(row.Cells[sheetColumn].Value, CultureInfo.InvariantCulture) ?? "";
            RevisionComparisonSheetInspection sheet = _inspection.Sources[revision].FindSheet(sheetName);
            row.Cells[headerColumn].Value = (sheet?.SuggestedHeaderRow ?? 1).ToString(CultureInfo.InvariantCulture);
            row.Cells[keyColumn].Value = sheet?.SuggestedKeyHeader ?? "";
        }

        private void PairCellChanged(object sender, DataGridViewCellEventArgs args)
        {
            if (_updatingGrid || args.RowIndex < 0 || args.ColumnIndex < 1 || args.ColumnIndex > 6) return;
            _updatingGrid = true;
            try
            {
                DataGridViewRow row = _pairs.Rows[args.RowIndex];
                if (args.ColumnIndex == 1) FillSuggestions(row, true);
                if (args.ColumnIndex == 2) FillSuggestions(row, false);
                row.Cells[7].Value = false;
            }
            finally
            {
                _updatingGrid = false;
            }
            InvalidatePreview();
        }

        private void RemoveSelectedPair()
        {
            if (_pairs.SelectedRows.Count == 0) return;
            _pairs.Rows.Remove(_pairs.SelectedRows[0]);
            InvalidatePreview();
        }

        private void InvalidatePreview(string message = null)
        {
            _preview.Text = message ??
                "Configuration changed. Preview keys again to refresh row, blank-key and duplicate-key counts.";
        }

        private void PreviewKeys()
        {
            try
            {
                RevisionComparisonPlan plan = BuildPlan(requireConfirmation: false);
                IReadOnlyList<RevisionComparisonKeyPreview> previews = RevisionComparisonEngine.PreviewKeys(_inspection, plan);
                _preview.Text = string.Join(Environment.NewLine, previews.Select(item =>
                    item.Transition + " | " + item.SheetPair + Environment.NewLine +
                    "  Before: " + item.BeforeRows + " rows, " + item.BeforeBlankKeys + " blank keys, " +
                    item.BeforeDuplicateRows + " duplicate rows" + Environment.NewLine +
                    "  After:  " + item.AfterRows + " rows, " + item.AfterBlankKeys + " blank keys, " +
                    item.AfterDuplicateRows + " duplicate rows"));
            }
            catch (Exception exception)
            {
                _preview.Text = "Cannot preview yet: " + exception.Message;
            }
        }

        private void CreatePlan()
        {
            try
            {
                ResultPlan = BuildPlan(requireConfirmation: true);
                RevisionComparisonEngine.PreviewKeys(_inspection, ResultPlan);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                Dialogs.Info("Comparison setup needs attention\n\n" + exception.Message);
            }
        }

        private RevisionComparisonPlan BuildPlan(bool requireConfirmation)
        {
            IReadOnlyList<int> order = OrderedIndexes();
            var pairs = new List<RevisionComparisonPairSpec>();
            foreach (DataGridViewRow row in _pairs.Rows)
            {
                PairTag tag = row.Tag as PairTag ?? throw new InvalidOperationException("A table-pair row is incomplete.");
                bool confirmed = Convert.ToBoolean(row.Cells[7].Value ?? false, CultureInfo.InvariantCulture);
                if (requireConfirmation && !confirmed)
                    throw new InvalidOperationException("Tick Confirm for every table pair after reviewing its sheets and key columns.");
                string beforeSheet = CellText(row, 1);
                string afterSheet = CellText(row, 2);
                int beforeHeader = ParsePositiveRow(CellText(row, 3), "Before header");
                int afterHeader = ParsePositiveRow(CellText(row, 4), "After header");
                string[] beforeKeys = ParseKeys(CellText(row, 5));
                string[] afterKeys = ParseKeys(CellText(row, 6));
                pairs.Add(new RevisionComparisonPairSpec(
                    tag.BeforeIndex,
                    tag.AfterIndex,
                    beforeSheet,
                    afterSheet,
                    beforeHeader,
                    afterHeader,
                    beforeKeys,
                    afterKeys));
            }
            return new RevisionComparisonPlan(order, pairs, order[0]);
        }

        private IReadOnlyList<int> OrderedIndexes() =>
            _revisions.Items.Cast<RevisionItem>().Select(item => item.SourceIndex).ToList();

        private string TransitionText(int before, int after) =>
            SourceLabel(before) + " -> " + SourceLabel(after);

        private string SourceLabel(int index)
        {
            RevisionComparisonSourceInspection source = _inspection.Sources[index];
            string parent = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(source.Path));
            return source.FileName + (string.IsNullOrEmpty(parent) ? "" : " [" + parent + "]");
        }

        private static string CellText(DataGridViewRow row, int column) =>
            (Convert.ToString(row.Cells[column].Value, CultureInfo.InvariantCulture) ?? "").Trim();

        private static int ParsePositiveRow(string value, string label)
        {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int row) ||
                row < 1 || row > 1048576)
                throw new InvalidOperationException(label + " must be an Excel row number from 1 to 1,048,576.");
            return row;
        }

        private static string[] ParseKeys(string value)
        {
            string[] keys = (value ?? "").Split(';').Select(item => item.Trim())
                .Where(item => item.Length > 0).ToArray();
            if (keys.Length == 0) throw new InvalidOperationException("Every table pair needs at least one explicit key column.");
            return keys;
        }

        private sealed class RevisionItem
        {
            internal RevisionItem(int sourceIndex, string path)
            {
                SourceIndex = sourceIndex;
                Path = path;
            }
            internal int SourceIndex { get; }
            internal string Path { get; }
            public override string ToString() => Path;
        }

        private sealed class TransitionItem
        {
            internal TransitionItem(int beforeIndex, int afterIndex, string text)
            {
                BeforeIndex = beforeIndex;
                AfterIndex = afterIndex;
                Text = text;
            }
            internal int BeforeIndex { get; }
            internal int AfterIndex { get; }
            internal string Text { get; }
            public override string ToString() => Text;
        }

        private sealed class PairTag
        {
            internal PairTag(int beforeIndex, int afterIndex)
            {
                BeforeIndex = beforeIndex;
                AfterIndex = afterIndex;
            }
            internal int BeforeIndex { get; }
            internal int AfterIndex { get; }
        }
    }
}
