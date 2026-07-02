using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Live view of the undo timeline: redoable steps on top, "now", then undoable steps (newest first).
    /// Select a step and jump straight to it with "Undo to here" / "Redo to here". Modeless — stays open
    /// while you work and refreshes itself on every change.
    /// </summary>
    internal sealed class TimelineForm : Form
    {
        private static TimelineForm _open;

        private readonly ListView _list = new ListView();
        private readonly Button _undoTo = new Button { Text = "Undo to here", Enabled = false };
        private readonly Button _redoTo = new Button { Text = "Redo to here", Enabled = false };
        private readonly Button _closeBtn = new Button { Text = "Close" };
        private readonly Label _hint = new Label { AutoSize = false };

        public static void ShowTimeline()
        {
            if (_open != null && !_open.IsDisposed) { _open.Activate(); return; }
            _open = new TimelineForm();
            _open.Show(ExcelWindow.Current);
        }

        private TimelineForm()
        {
            Text = "Power Tools — Timeline";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            ShowInTaskbar = false;
            MinimumSize = new Size(430, 320);
            ClientSize = new Size(520, 420);
            Font = new Font("Segoe UI", 9f);

            _list.Location = new Point(12, 12);
            _list.Size = new Size(ClientSize.Width - 24, ClientSize.Height - 96);
            _list.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.MultiSelect = false;
            _list.HideSelection = false;
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            _list.Columns.Add("", 30);
            _list.Columns.Add("Action", 240);
            _list.Columns.Add("Sheet", 110);
            _list.Columns.Add("Cells", 50);
            _list.Columns.Add("When", 60);
            _list.SelectedIndexChanged += (_, __) => UpdateButtons();
            _list.DoubleClick += (_, __) => JumpToSelected();
            Controls.Add(_list);

            _hint.Location = new Point(12, ClientSize.Height - 74);
            _hint.Size = new Size(ClientSize.Width - 24, 24);
            _hint.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _hint.Text = "Double-click a step to jump to it. ✎ = manual edit, ▸ = Power Tools operation.";
            Controls.Add(_hint);

            _undoTo.Size = new Size(110, 30);
            _undoTo.Location = new Point(12, ClientSize.Height - 42);
            _undoTo.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _undoTo.Click += (_, __) => JumpToSelected();
            Controls.Add(_undoTo);

            _redoTo.Size = new Size(110, 30);
            _redoTo.Location = new Point(130, ClientSize.Height - 42);
            _redoTo.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            _redoTo.Click += (_, __) => JumpToSelected();
            Controls.Add(_redoTo);

            _closeBtn.Size = new Size(90, 30);
            _closeBtn.Location = new Point(ClientSize.Width - 102, ClientSize.Height - 42);
            _closeBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _closeBtn.Click += (_, __) => Close();
            Controls.Add(_closeBtn);

            CancelButton = _closeBtn;

            UndoManager.Changed += OnStacksChanged;
            FormClosed += (_, __) =>
            {
                UndoManager.Changed -= OnStacksChanged;
                if (_open == this) _open = null;
            };

            RefreshList();
        }

        private void OnStacksChanged()
        {
            if (IsDisposed) return;
            try { BeginInvoke((Action)RefreshList); } catch { }
        }

        private void RefreshList()
        {
            if (IsDisposed) return;

            _list.BeginUpdate();
            _list.Items.Clear();

            var redo = UndoManager.RedoEntries;
            var undo = UndoManager.UndoEntries;

            // Redoable steps: oldest at the top, the NEXT redo ends up adjacent to the "now" marker.
            for (int i = 0; i < redo.Count; i++)
                _list.Items.Add(MakeItem(redo[i], isRedo: true, steps: redo.Count - i));

            var now = new ListViewItem("");
            now.SubItems.Add("——— now ———");
            now.SubItems.Add("");
            now.SubItems.Add("");
            now.SubItems.Add("");
            now.ForeColor = Color.DimGray;
            now.Font = new Font(Font, FontStyle.Bold);
            _list.Items.Add(now);

            // Undoable steps: newest first (the next undo right under the marker).
            for (int i = undo.Count - 1; i >= 0; i--)
                _list.Items.Add(MakeItem(undo[i], isRedo: false, steps: undo.Count - i));

            _list.EndUpdate();
            UpdateButtons();
        }

        private ListViewItem MakeItem(UndoEntry e, bool isRedo, int steps)
        {
            var item = new ListViewItem(e.IsManualEdit ? "✎" : "▸");
            item.SubItems.Add(e.Description ?? "");
            item.SubItems.Add(e.SheetName ?? "");
            item.SubItems.Add(e.CellCount.ToString());
            item.SubItems.Add(e.Timestamp.ToString("HH:mm:ss"));
            item.Tag = (isRedo, steps);
            if (isRedo) item.ForeColor = Color.Gray;
            return item;
        }

        private void UpdateButtons()
        {
            _undoTo.Enabled = false;
            _redoTo.Enabled = false;
            if (_list.SelectedItems.Count != 1 || !(_list.SelectedItems[0].Tag is ValueTuple<bool, int> tag)) return;

            if (tag.Item1) _redoTo.Enabled = true;
            else _undoTo.Enabled = true;
        }

        private void JumpToSelected()
        {
            if (_list.SelectedItems.Count != 1 || !(_list.SelectedItems[0].Tag is ValueTuple<bool, int> tag)) return;

            bool isRedo = tag.Item1;
            int steps = tag.Item2;

            // Run in a macro context so the write is safe even from a modeless window.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (isRedo) UndoManager.RedoSteps(steps);
                else UndoManager.UndoSteps(steps);
            });
        }
    }
}
