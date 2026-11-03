using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelPowerTools.Core.Comparison.RevisionBridge;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>Collects revisions across one or more folders before workbook inspection begins.</summary>
    internal sealed class RevisionComparisonFilePickerForm : Form
    {
        private readonly string _filter;
        private readonly ListBox _paths = new ListBox();
        private readonly Button _continue = new Button();
        private bool _openedInitialPicker;

        internal RevisionComparisonFilePickerForm(string filter)
        {
            _filter = filter ?? throw new ArgumentNullException(nameof(filter));
            Text = "Choose table revisions";
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimumSize = new Size(720, 390);
            Size = new Size(900, 520);
            Font = new Font("Segoe UI", 9F);
            BuildLayout();
            Shown += (_, __) =>
            {
                if (_openedInitialPicker) return;
                _openedInitialPicker = true;
                BeginInvoke(new Action(AddFiles));
            };
            GdiText.Apply(this);
        }

        internal IReadOnlyList<string> SelectedPaths =>
            new ReadOnlyCollection<string>(_paths.Items.Cast<string>().ToList());

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(14),
                ColumnCount = 1,
                RowCount = 3,
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            root.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "Add 2 to 20 .xlsx or .xlsm revisions. You can use Add files more than once " +
                       "to collect revisions from different folders; order them on the next screen.",
                Margin = new Padding(0, 0, 0, 10),
            }, 0, 0);

            _paths.Dock = DockStyle.Fill;
            _paths.IntegralHeight = false;
            _paths.SelectionMode = SelectionMode.MultiExtended;
            _paths.HorizontalScrollbar = true;
            root.Controls.Add(_paths, 0, 1);

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = new Padding(0, 10, 0, 0),
            };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            _continue.Text = "Continue";
            _continue.AutoSize = true;
            _continue.Click += (_, __) => Continue();
            var remove = new Button { Text = "Remove selected", AutoSize = true };
            remove.Click += (_, __) => RemoveSelected();
            var add = new Button { Text = "Add files...", AutoSize = true };
            add.Click += (_, __) => AddFiles();
            actions.Controls.Add(cancel);
            actions.Controls.Add(_continue);
            actions.Controls.Add(remove);
            actions.Controls.Add(add);
            root.Controls.Add(actions, 0, 2);
            AcceptButton = _continue;
            CancelButton = cancel;
        }

        private void AddFiles()
        {
            using (var picker = new OpenFileDialog
            {
                Title = "Add table revisions",
                Filter = _filter,
                Multiselect = true,
                CheckFileExists = true,
                DereferenceLinks = true,
            })
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                var existing = new HashSet<string>(_paths.Items.Cast<string>(), StringComparer.OrdinalIgnoreCase);
                foreach (string selected in picker.FileNames.Select(Path.GetFullPath))
                {
                    if (existing.Contains(selected)) continue;
                    if (_paths.Items.Count >= RevisionComparisonEngine.MaximumRevisionCount)
                    {
                        Dialogs.Info("A comparison can contain at most 20 revisions.", Text, this);
                        break;
                    }
                    existing.Add(selected);
                    _paths.Items.Add(selected);
                }
            }
        }

        private void RemoveSelected()
        {
            int[] indexes = _paths.SelectedIndices.Cast<int>().OrderByDescending(item => item).ToArray();
            foreach (int index in indexes) _paths.Items.RemoveAt(index);
        }

        private void Continue()
        {
            if (_paths.Items.Count < 2)
            {
                Dialogs.Info("Add at least two workbook revisions before continuing.", Text, this);
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
