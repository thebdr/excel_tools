using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.IoPipeline;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.IoPipeline.IoBridge
{
    internal sealed class IoBridgeSetupResult
    {
        public string IoListPath = "";
        public string InputsCsvPath = "";
        public string OutputsCsvPath = "";
        public List<string> CardSheets = new List<string>();
        public int IoColumnLimit;   // 0 = all source columns
        public string VirtualIncludeRegex = "";
        public string ExcludeMnemonicRegex = "";
    }

    /// <summary>
    /// Workspace setup: the IO-list workbook, which of its card sheets, and the Inputs/Outputs CSV
    /// pair. Sheet names are listed via OpenXml over a share-tolerant byte copy, so a workbook already
    /// open in Excel still lists. Only R{rack}.S{slot}.C{card}-named sheets are selectable — the card
    /// number comes from the sheet name and nothing else can join.
    /// </summary>
    internal sealed class IoBridgeSetupForm : Form
    {
        private readonly TextBox _ioList = new TextBox { ReadOnly = true };
        private readonly TextBox _inputs = new TextBox { ReadOnly = true };
        private readonly TextBox _outputs = new TextBox { ReadOnly = true };
        private readonly CheckedListBox _sheets = new CheckedListBox { CheckOnClick = true, IntegralHeight = false };
        private readonly NumericUpDown _columnLimit = new NumericUpDown
        {
            Minimum = 0,
            Maximum = 200,
            Value = 0,
            Width = 70,
        };
        private readonly TextBox _virtualInclude = new TextBox();
        private readonly TextBox _excludeMnemonics = new TextBox();
        private readonly Label _hint = new Label { AutoSize = false, ForeColor = System.Drawing.Color.FromArgb(156, 0, 6) };
        private readonly Button _ok = new Button { Text = "Build workspace", DialogResult = DialogResult.OK, Enabled = false };

        public IoBridgeSetupResult Result { get; private set; }

        public IoBridgeSetupForm()
        {
            Text = "IO Bridge — build workspace";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 540);
            MinimumSize = new Size(560, 480);
            Font = SystemFonts.MessageBoxFont;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 3,
                RowCount = 9,
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

            AddPathRow(layout, 0, "IO list", _ioList, BrowseIoList);
            AddPathRow(layout, 1, "Inputs CSV", _inputs, () => BrowseCsv(_inputs, "Inputs"));
            AddPathRow(layout, 2, "Outputs CSV", _outputs, () => BrowseCsv(_outputs, "Outputs"));

            layout.Controls.Add(new Label { Text = "Card sheets", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 3);
            layout.Controls.Add(_sheets, 1, 3);
            layout.SetColumnSpan(_sheets, 2);
            // Exactly ONE style per row, in row order — AddPathRow added rows 0-2; this is row 3.
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _sheets.Dock = DockStyle.Fill;

            layout.Controls.Add(new Label
            {
                Text = "IO column limit",
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
            }, 0, 4);
            var limitPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill };
            limitPanel.Controls.Add(_columnLimit);
            limitPanel.Controls.Add(new Label
            {
                Text = "first N source columns per sheet (0 = all)",
                AutoSize = true,
                Padding = new Padding(4, 6, 0, 0),
            });
            layout.Controls.Add(limitPanel, 1, 4);
            layout.SetColumnSpan(limitPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

            AddRegexRow(layout, 5, "Compare virtuals", _virtualInclude,
                "regex — VIRTUAL signals (types 5/6/9/10/13/14) whose mnemonic matches join by name, e.g. ^CLASSDIAGAL");
            AddRegexRow(layout, 6, "Exclude mnemonics", _excludeMnemonics,
                "regex — IO-list MNEMONICS that match drop out of the comparison (generator keys beside a real one, e.g. ^SAFETY_NUMDOORS); a cell whose every mnemonic matches is ignored whole");

            _hint.Dock = DockStyle.Fill;
            _hint.Height = 34;
            layout.Controls.Add(_hint, 1, 7);
            layout.SetColumnSpan(_hint, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 88 };
            _ok.Width = 120;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_ok);
            layout.Controls.Add(buttons, 1, 8);
            layout.SetColumnSpan(buttons, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            Controls.Add(layout);
            ExcelPowerTools.UI.GdiText.Apply(this);   // GDI text, like the native controls
            AcceptButton = _ok;
            CancelButton = cancel;

            // ItemCheck also fires for Items.Add(name, isChecked: true) during the constructor prefill,
            // before the handle exists — BeginInvoke would throw and abort the check before it is stored.
            _sheets.ItemCheck += (s, e) => { if (IsHandleCreated) BeginInvoke(new Action(Revalidate)); };

            // Long paths show their tail: caret at the end and the view scrolled there whenever
            // code fills a box (they are read-only, so TextChanged only ever comes from code).
            foreach (TextBox box in new[] { _ioList, _inputs, _outputs })
            {
                box.HandleCreated += (s, e) => { var b = (TextBox)s; b.SelectionStart = b.Text.Length; b.ScrollToCaret(); };
                box.TextChanged += (s, e) => { var b = (TextBox)s; b.SelectionStart = b.Text.Length; b.ScrollToCaret(); };
            }
            Shown += (s, e) =>
            {
                foreach (TextBox box in new[] { _ioList, _inputs, _outputs })
                {
                    box.SelectionStart = box.Text.Length;
                    box.ScrollToCaret();
                }
            };
            FormClosing += OnClosingCollect;

            // Registry-remembered paths pre-populate the pickers.
            _ioList.Text = IoBridgeRegistry.GetPath("IoListPath") ?? "";
            _inputs.Text = IoBridgeRegistry.GetPath("InputsCsvPath") ?? "";
            _outputs.Text = IoBridgeRegistry.GetPath("OutputsCsvPath") ?? "";
            if (int.TryParse(IoBridgeRegistry.GetPath("IoColumnLimit"), out int limit) &&
                limit >= _columnLimit.Minimum && limit <= _columnLimit.Maximum)
                _columnLimit.Value = limit;
            _virtualInclude.Text = IoBridgeRegistry.GetPath("VirtualIncludeRegex") ?? "";
            _excludeMnemonics.Text = IoBridgeRegistry.GetPath("ExcludeMnemonicRegex") ?? "";
            if (File.Exists(_ioList.Text)) LoadSheetList();
            Revalidate();
        }

        private void AddRegexRow(TableLayoutPanel layout, int row, string label, TextBox box, string hint)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0) };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            box.Dock = DockStyle.Fill;
            panel.Controls.Add(box, 0, 0);
            panel.Controls.Add(new Label
            {
                Text = hint,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = System.Drawing.SystemColors.GrayText,
            }, 1, 0);
            layout.Controls.Add(panel, 1, row);
            layout.SetColumnSpan(panel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            box.TextChanged += (s, e) => Revalidate();
        }

        private void AddPathRow(TableLayoutPanel layout, int row, string label, TextBox box, Action browse)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            box.Dock = DockStyle.Fill;
            layout.Controls.Add(box, 1, row);
            var button = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            button.Click += (s, e) => browse();
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        }

        private void BrowseIoList()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "Excel IO list|*.xlsx;*.xlsm|All files|*.*",
                CheckFileExists = true,
                FileName = _ioList.Text,
            })
            {
                if (File.Exists(_ioList.Text))
                    try { dialog.InitialDirectory = Path.GetDirectoryName(_ioList.Text); } catch { }
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                _ioList.Text = dialog.FileName;
                LoadSheetList();
                Revalidate();
            }
        }

        private void BrowseCsv(TextBox box, string direction)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "Plant CSV|*.csv|All files|*.*",
                CheckFileExists = true,
                Title = "Select the " + direction + " CSV",
            })
            {
                string start = File.Exists(box.Text) ? box.Text
                    : File.Exists(_inputs.Text) ? _inputs.Text : _ioList.Text;
                if (File.Exists(start))
                    try { dialog.InitialDirectory = Path.GetDirectoryName(start); } catch { }
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                box.Text = dialog.FileName;
                Revalidate();
            }
        }

        private void LoadSheetList()
        {
            _sheets.Items.Clear();
            try
            {
                foreach (string name in ListSheetNames(_ioList.Text))
                    if (IoListLogic.TryParseCardSheetName(name, out _, out _, out _))
                        _sheets.Items.Add(name, true);

                if (_sheets.Items.Count == 0)
                    _hint.Text = "No R{rack}.S{slot}.C{card} sheet in this workbook — nothing can join.";
            }
            catch (Exception ex)
            {
                _hint.Text = "Could not list the workbook's sheets: " + ex.Message;
            }
        }

        /// <summary>Sheet names via OpenXml over an in-memory copy — works while Excel holds the file.</summary>
        private static List<string> ListSheetNames(string path)
        {
            byte[] bytes;
            using (var input = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var buffer = new MemoryStream())
            {
                input.CopyTo(buffer);
                bytes = buffer.ToArray();
            }

            var names = new List<string>();
            using (var stream = new MemoryStream(bytes, writable: false))
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                Sheets sheets = document.WorkbookPart?.Workbook?.Sheets;
                if (sheets != null)
                    foreach (Sheet sheet in sheets.Elements<Sheet>())
                        if (sheet.Name != null && sheet.Name.HasValue) names.Add(sheet.Name.Value);
            }
            return names;
        }

        private void Revalidate()
        {
            var problems = new List<string>();
            if (!File.Exists(_ioList.Text)) problems.Add("pick the IO-list workbook");
            if (!File.Exists(_inputs.Text)) problems.Add("pick the Inputs CSV");
            if (!File.Exists(_outputs.Text)) problems.Add("pick the Outputs CSV");
            if (_sheets.CheckedItems.Count == 0) problems.Add("tick at least one card sheet");

            // Direction sanity: a file named the other way round is almost always a mixed-up pick.
            if (File.Exists(_inputs.Text) && IoListLogic.DirectionFromName(Path.GetFileName(_inputs.Text)) == 'O')
                problems.Add("the Inputs pick is named like an OUTPUTS file");
            if (File.Exists(_outputs.Text) && IoListLogic.DirectionFromName(Path.GetFileName(_outputs.Text)) == 'I')
                problems.Add("the Outputs pick is named like an INPUTS file");
            if (File.Exists(_inputs.Text) && File.Exists(_outputs.Text) &&
                string.Equals(Path.GetFullPath(_inputs.Text), Path.GetFullPath(_outputs.Text), StringComparison.OrdinalIgnoreCase))
                problems.Add("Inputs and Outputs are the same file");

            if (!IsValidRegex(_virtualInclude.Text)) problems.Add("the virtuals regex does not parse");
            if (!IsValidRegex(_excludeMnemonics.Text)) problems.Add("the exclude regex does not parse");

            _hint.Text = problems.Count == 0 ? "" : "· " + string.Join("   · ", problems);
            _ok.Enabled = problems.Count == 0;
        }

        private void OnClosingCollect(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK) return;

            Result = new IoBridgeSetupResult
            {
                IoListPath = _ioList.Text,
                InputsCsvPath = _inputs.Text,
                OutputsCsvPath = _outputs.Text,
                IoColumnLimit = (int)_columnLimit.Value,
                VirtualIncludeRegex = _virtualInclude.Text.Trim(),
                ExcludeMnemonicRegex = _excludeMnemonics.Text.Trim(),
            };
            foreach (object item in _sheets.CheckedItems) Result.CardSheets.Add((string)item);

            IoBridgeRegistry.SetPath("IoListPath", Result.IoListPath);
            IoBridgeRegistry.SetPath("InputsCsvPath", Result.InputsCsvPath);
            IoBridgeRegistry.SetPath("OutputsCsvPath", Result.OutputsCsvPath);
            IoBridgeRegistry.SetPath("IoColumnLimit", Result.IoColumnLimit.ToString(CultureInfo.InvariantCulture));
            IoBridgeRegistry.SetPath("VirtualIncludeRegex", Result.VirtualIncludeRegex);
            IoBridgeRegistry.SetPath("ExcludeMnemonicRegex", Result.ExcludeMnemonicRegex);
        }

        private static bool IsValidRegex(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern)) return true;
            try
            {
                _ = new System.Text.RegularExpressions.Regex(pattern.Trim());
                return true;
            }
            catch (ArgumentException) { return false; }
        }

        public static IoBridgeSetupResult Ask()
        {
            using (var form = new IoBridgeSetupForm())
                return form.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? form.Result : null;
        }
    }
}
