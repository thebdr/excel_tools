using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Analysis;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Live profile of the active cell's column: counts, numeric range, lengths, top values. Modeless;
    /// follows the selection as you move around the sheet.
    /// </summary>
    internal sealed class ColumnStatsForm : Form
    {
        private const int MaxRows = 200000;

        private static ColumnStatsForm _open;

        private readonly Label _title = new Label { AutoSize = false };
        private readonly ListView _stats = new ListView();
        private readonly ListView _top = new ListView();

        public static void ShowPanel()
        {
            if (_open != null && !_open.IsDisposed) { _open.Activate(); return; }
            _open = new ColumnStatsForm();
            _open.Show(ExcelWindow.Current);
        }

        private ColumnStatsForm()
        {
            Text = "Column Stats";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            ShowInTaskbar = false;
            MinimumSize = new Size(340, 420);
            ClientSize = new Size(360, 520);
            Font = new Font("Segoe UI", 9f);

            _title.Location = new Point(12, 10);
            _title.Size = new Size(ClientSize.Width - 24, 22);
            _title.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _title.Font = new Font(Font, FontStyle.Bold);
            Controls.Add(_title);

            _stats.Location = new Point(12, 38);
            _stats.Size = new Size(ClientSize.Width - 24, 190);
            _stats.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _stats.View = View.Details;
            _stats.HeaderStyle = ColumnHeaderStyle.None;
            _stats.FullRowSelect = true;
            _stats.Columns.Add("Stat", 150);
            _stats.Columns.Add("Value", 160);
            Controls.Add(_stats);

            var topLabel = new Label { Text = "Top values:", Location = new Point(12, 236), AutoSize = true };
            Controls.Add(topLabel);

            _top.Location = new Point(12, 258);
            _top.Size = new Size(ClientSize.Width - 24, ClientSize.Height - 270);
            _top.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _top.View = View.Details;
            _top.FullRowSelect = true;
            _top.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            _top.Columns.Add("Value", 180);
            _top.Columns.Add("Count", 60);
            _top.Columns.Add("%", 50);
            Controls.Add(_top);

            SelectionEvents.Changed += OnSelectionChanged;
            FormClosed += (_, __) =>
            {
                SelectionEvents.Changed -= OnSelectionChanged;
                if (_open == this) _open = null;
            };

            Refresh_();
        }

        private void OnSelectionChanged()
        {
            if (IsDisposed) return;
            try { BeginInvoke((Action)Refresh_); } catch { }
        }

        private void Refresh_()
        {
            if (IsDisposed) return;

            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic cellRef = app.ActiveCell;
                if (cellRef == null) { _title.Text = "(no active cell)"; return; }

                dynamic sheet = cellRef.Worksheet;
                int col = (int)cellRef.Column;
                long lastRow = RangeHelpers.LastRow(sheet);
                bool truncated = lastRow > MaxRows;
                if (truncated) lastRow = MaxRows;
                if (lastRow < 1) lastRow = 1;

                dynamic range = sheet.Range(sheet.Cells(1, col), sheet.Cells(lastRow, col));
                object raw = range.Value2;

                var values = new System.Collections.Generic.List<string>((int)lastRow);
                if (raw is object[,] g)
                {
                    int rb = g.GetLowerBound(0), cb = g.GetLowerBound(1);
                    int n = g.GetLength(0);
                    for (int r = 0; r < n; r++)
                        values.Add(CsvSheet.CellText(g[rb + r, cb], sheet, r + 1, col));
                }
                else
                {
                    values.Add(CsvSheet.CellText(raw, sheet, 1, col));
                }

                var s = ColumnStats.Compute(values);

                _title.Text = $"Column {ColLetters(col)} — {s.Total} row(s) on {(string)sheet.Name}" +
                              (truncated ? " (first 200k)" : "");

                _stats.BeginUpdate();
                _stats.Items.Clear();
                AddStat("Non-empty", s.NonEmpty.ToString());
                AddStat("Empty", s.Empty.ToString());
                AddStat("Distinct", s.Distinct.ToString());
                AddStat("Numeric", s.NumericCount.ToString());
                AddStat("Min / Max", s.Min == null ? "—" : $"{s.Min} / {s.Max}");
                AddStat("Length min / max", s.NonEmpty == 0 ? "—" : $"{s.MinLength} / {s.MaxLength}");
                _stats.EndUpdate();

                _top.BeginUpdate();
                _top.Items.Clear();
                foreach (var kv in s.TopValues)
                {
                    var item = new ListViewItem(kv.Key.Length > 60 ? kv.Key.Substring(0, 60) + "…" : kv.Key);
                    item.SubItems.Add(kv.Value.ToString());
                    item.SubItems.Add(s.NonEmpty > 0 ? (100.0 * kv.Value / s.NonEmpty).ToString("0.#") : "0");
                    _top.Items.Add(item);
                }
                _top.EndUpdate();
            }
            catch
            {
                _title.Text = "(couldn't read the selection)";
            }
        }

        private void AddStat(string name, string value)
        {
            var item = new ListViewItem(name);
            item.SubItems.Add(value);
            _stats.Items.Add(item);
        }

        internal static string ColLetters(int col)
        {
            string s = "";
            while (col > 0)
            {
                col--;
                s = (char)('A' + col % 26) + s;
                col /= 26;
            }
            return s;
        }
    }
}
