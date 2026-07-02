using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ExcelPowerTools.Core.Text;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Regex find &amp; replace with a before/after preview. The form owns only the UI — the owner wires
    /// <see cref="ComputeChanges"/> (scan the sheet) and <see cref="ApplyChanges"/> (write the cells).
    /// </summary>
    internal sealed class RegexReplaceForm : Form
    {
        private const int PreviewCap = 1000;   // rows shown in the list; Apply always applies all

        private readonly TextBox _find = new TextBox();
        private readonly TextBox _replace = new TextBox();
        private readonly CheckBox _ignoreCase = new CheckBox { Text = "Ignore case" };
        private readonly CheckBox _literal = new CheckBox { Text = "Plain text (no regex)" };
        private readonly CheckBox _wholeCell = new CheckBox { Text = "Whole cell" };
        private readonly CheckBox _visibleOnly = new CheckBox { Text = "Visible cells only", Checked = true };
        private readonly RadioButton _scopeSelection = new RadioButton { Text = "Selection" };
        private readonly RadioButton _scopeSheet = new RadioButton { Text = "Whole sheet" };
        private readonly Button _preview = new Button { Text = "Preview" };
        private readonly Button _apply = new Button { Text = "Apply", Enabled = false };
        private readonly Button _closeBtn = new Button { Text = "Close" };
        private readonly ListView _list = new ListView();
        private readonly Label _status = new Label { AutoSize = false };

        private List<RegexCellChange> _pending;

        /// <summary>(spec, selectionScope, visibleOnly) → the changes that would be made.</summary>
        public Func<RegexReplaceSpec, bool, bool, List<RegexCellChange>> ComputeChanges;

        /// <summary>Writes the changes to the sheet.</summary>
        public Action<List<RegexCellChange>> ApplyChanges;

        public RegexReplaceSpec CurrentSpec => new RegexReplaceSpec
        {
            Pattern = _find.Text,
            Replacement = _replace.Text,
            IgnoreCase = _ignoreCase.Checked,
            LiteralPattern = _literal.Checked,
            WholeCell = _wholeCell.Checked,
        };

        public bool VisibleOnly => _visibleOnly.Checked;

        public RegexReplaceForm(RegexReplaceSpec initial, bool preferSelectionScope)
        {
            Text = "Regex Replace";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            ShowInTaskbar = false;
            MinimumSize = new Size(560, 420);
            ClientSize = new Size(660, 460);
            Font = new Font("Segoe UI", 9f);

            const int labelX = 12, inputX = 90;
            int y = 14;

            AddLabel("Find:", labelX, y);
            Place(_find, inputX, y, ClientSize.Width - inputX - 12, anchorRight: true);
            y += 30;

            AddLabel("Replace:", labelX, y);
            Place(_replace, inputX, y, ClientSize.Width - inputX - 12, anchorRight: true);
            y += 32;

            _ignoreCase.Location = new Point(inputX, y); _ignoreCase.AutoSize = true;
            _literal.Location = new Point(inputX + 110, y); _literal.AutoSize = true;
            _wholeCell.Location = new Point(inputX + 280, y); _wholeCell.AutoSize = true;
            _visibleOnly.Location = new Point(inputX + 390, y); _visibleOnly.AutoSize = true;
            Controls.AddRange(new Control[] { _ignoreCase, _literal, _wholeCell, _visibleOnly });
            y += 28;

            AddLabel("Scope:", labelX, y);
            _scopeSelection.Location = new Point(inputX, y - 2); _scopeSelection.AutoSize = true;
            _scopeSheet.Location = new Point(inputX + 110, y - 2); _scopeSheet.AutoSize = true;
            (preferSelectionScope ? _scopeSelection : _scopeSheet).Checked = true;
            Controls.Add(_scopeSelection);
            Controls.Add(_scopeSheet);

            _preview.Location = new Point(ClientSize.Width - 102, y - 6);
            _preview.Size = new Size(90, 28);
            _preview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _preview.Click += (_, __) => RunPreview();
            Controls.Add(_preview);
            y += 34;

            _list.Location = new Point(12, y);
            _list.Size = new Size(ClientSize.Width - 24, ClientSize.Height - y - 52);
            _list.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            _list.Columns.Add("Cell", 70);
            _list.Columns.Add("Before", 270);
            _list.Columns.Add("After", 270);
            Controls.Add(_list);

            _status.Location = new Point(12, ClientSize.Height - 40);
            _status.Size = new Size(ClientSize.Width - 220, 30);
            _status.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _status.Text = "Enter a pattern and press Preview.";
            Controls.Add(_status);

            _apply.Size = new Size(90, 30);
            _apply.Location = new Point(ClientSize.Width - 200, ClientSize.Height - 42);
            _apply.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _apply.Click += (_, __) => RunApply();
            Controls.Add(_apply);

            _closeBtn.Size = new Size(90, 30);
            _closeBtn.Location = new Point(ClientSize.Width - 102, ClientSize.Height - 42);
            _closeBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _closeBtn.DialogResult = DialogResult.Cancel;
            Controls.Add(_closeBtn);

            AcceptButton = _preview;
            CancelButton = _closeBtn;

            if (initial != null)
            {
                _find.Text = initial.Pattern;
                _replace.Text = initial.Replacement;
                _ignoreCase.Checked = initial.IgnoreCase;
                _literal.Checked = initial.LiteralPattern;
                _wholeCell.Checked = initial.WholeCell;
            }

            // Any input change invalidates a shown preview.
            EventHandler stale = (_, __) => InvalidatePreview();
            _find.TextChanged += stale;
            _replace.TextChanged += stale;
            _ignoreCase.CheckedChanged += stale;
            _literal.CheckedChanged += stale;
            _wholeCell.CheckedChanged += stale;
            _visibleOnly.CheckedChanged += stale;
            _scopeSelection.CheckedChanged += stale;

            Shown += (_, __) => _find.Focus();
        }

        private void InvalidatePreview()
        {
            _pending = null;
            _apply.Enabled = false;
        }

        private void RunPreview()
        {
            if (string.IsNullOrEmpty(_find.Text)) { _status.Text = "Enter a pattern first."; return; }

            try { RegexReplacer.Build(CurrentSpec); }
            catch (ArgumentException ex) { _status.Text = "Invalid pattern: " + ex.Message; return; }

            try
            {
                Cursor = Cursors.WaitCursor;
                _pending = ComputeChanges?.Invoke(CurrentSpec, _scopeSelection.Checked, _visibleOnly.Checked)
                           ?? new List<RegexCellChange>();
            }
            catch (Exception ex)
            {
                _status.Text = "Scan failed: " + ex.Message;
                _pending = null;
                return;
            }
            finally { Cursor = Cursors.Default; }

            _list.BeginUpdate();
            _list.Items.Clear();
            int shown = 0;
            foreach (var c in _pending)
            {
                if (shown++ >= PreviewCap) break;
                var item = new ListViewItem(CellRef(c.Row, c.Column));
                item.SubItems.Add(c.Before);
                item.SubItems.Add(c.After);
                _list.Items.Add(item);
            }
            _list.EndUpdate();

            _status.Text = _pending.Count == 0
                ? "No matches."
                : $"{_pending.Count} cell(s) will change" + (_pending.Count > PreviewCap ? $" (showing first {PreviewCap})" : "") + ".";
            _apply.Enabled = _pending.Count > 0;
        }

        private void RunApply()
        {
            if (_pending == null || _pending.Count == 0) return;
            try
            {
                Cursor = Cursors.WaitCursor;
                ApplyChanges?.Invoke(_pending);
                _status.Text = $"Applied {_pending.Count} change(s).";
            }
            catch (Exception ex)
            {
                _status.Text = "Apply failed: " + ex.Message;
                return;
            }
            finally { Cursor = Cursors.Default; }

            InvalidatePreview();
            _list.Items.Clear();
        }

        /// <summary>R1C1 → "B5"-style reference for the preview list.</summary>
        private static string CellRef(int row, int col)
        {
            string letters = "";
            int c = col;
            while (c > 0)
            {
                c--;
                letters = (char)('A' + c % 26) + letters;
                c /= 26;
            }
            return letters + row;
        }

        private void AddLabel(string text, int x, int y)
        {
            Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), AutoSize = true });
        }

        private void Place(Control ctl, int x, int y, int width, bool anchorRight = false)
        {
            ctl.Location = new Point(x, y);
            ctl.Size = new Size(width, 24);
            if (anchorRight) ctl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(ctl);
        }
    }
}
