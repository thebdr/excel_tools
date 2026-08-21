using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelPowerTools.Core.IoPipeline;
using ExcelPowerTools.Core.IoPipeline.SwpBridge;

namespace ExcelPowerTools.IoPipeline.SwpBridge
{
    /// <summary>
    /// Workspace setup. Picking the SWP file inside a dated version folder derives the system
    /// root (its parent) and the current version; a sidecar found at that root pre-fills every
    /// other field, so returning to a system is one pick + OK. All plant CSVs are optional —
    /// their checks simply stay off. The .xls format is refused (the headless patcher writes
    /// OpenXml); .xlsx/.xlsm both work.
    /// </summary>
    internal sealed class SwpBridgeSetupForm : Form
    {
        private readonly TextBox _swp = new TextBox { ReadOnly = true };
        private readonly Label _derived = new Label { AutoSize = false, ForeColor = SystemColors.GrayText };
        private readonly RadioButton _msc = new RadioButton { Text = "MSC (sorter)", AutoSize = true, Checked = true };
        private readonly RadioButton _ispc = new RadioButton { Text = "ISPC (induction)", AutoSize = true };
        private readonly RadioButton _ups = new RadioButton { Text = "UPS (upstream)", AutoSize = true };
        private readonly RadioButton _safety = new RadioButton { Text = "SAFETY (PLC)", AutoSize = true };
        private readonly NumericUpDown _firstLine = new NumericUpDown { Minimum = 1, Maximum = 99, Value = 1, Width = 56 };
        private readonly NumericUpDown _group = new NumericUpDown { Minimum = 1, Maximum = 99, Value = 1, Width = 56 };
        private readonly TextBox _secondLangCode = new TextBox { Width = 44, CharacterCasing = CharacterCasing.Upper, MaxLength = 2 };
        private readonly TextBox _ioList = new TextBox { ReadOnly = true };
        private readonly CheckedListBox _sheets = new CheckedListBox { CheckOnClick = true, IntegralHeight = false };
        private readonly TextBox _plantInputs = new TextBox();
        private readonly TextBox _busfield = new TextBox();
        private readonly TextBox _destinationParams = new TextBox();
        private readonly TextBox _sensor = new TextBox();
        private readonly TextBox _strobe = new TextBox();
        private readonly TextBox _glossary1 = new TextBox();
        private readonly TextBox _glossary2 = new TextBox();
        private readonly TextBox _referenceSwp = new TextBox();
        private readonly TextBox _alarmStrings = new TextBox();
        private readonly TextBox _includeRegex = new TextBox();
        private readonly TextBox _excludeRegex = new TextBox();
        private readonly CheckBox _glossaryAlways = new CheckBox
        {
            Text = "Glossary always wins (replace the IO list's language 2; otherwise it only fills the gaps)",
            AutoSize = true,
        };
        private readonly CheckBox _trustConfig = new CheckBox { Text = "Trust CONFIG_PLANT descriptions (generated rows write only the label addition)", AutoSize = true };
        private readonly CheckBox _rewriteFormulas = new CheckBox { Text = "Rewrite the Q/R description formulas wholesale", AutoSize = true };
        private readonly RadioButton _mismatchAppend = new RadioButton
        {
            Text = "append identity to descriptions (device blanks; template formulas extend)",
            AutoSize = true,
            Checked = true,
        };
        private readonly RadioButton _mismatchOverwrite = new RadioButton
        {
            Text = "overwrite with the IO values as-is",
            AutoSize = true,
        };
        private readonly RadioButton _inPlace = new RadioButton { Text = "in place (with .bak)", AutoSize = true, Checked = true };
        private readonly RadioButton _versionFolder = new RadioButton { Text = "new dated version folder", AutoSize = true };
        private readonly Label _hint = new Label { AutoSize = false, ForeColor = System.Drawing.Color.FromArgb(156, 0, 6) };
        private readonly Button _ok = new Button { Text = "Build workspace", DialogResult = DialogResult.OK, Enabled = false };

        public SwpSidecarDocument Result { get; private set; }
        public string SwpRootFolder { get; private set; } = "";

        private SwpSidecarDocument _loaded;   // an existing sidecar at the derived root

        public SwpBridgeSetupForm()
        {
            Text = "SWP Bridge — build workspace";
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(720, 846);
            MinimumSize = new Size(640, 700);
            Font = SystemFonts.MessageBoxFont;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 3,
                RowCount = 20,
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

            AddPathRow(layout, 0, "SWP file", _swp, BrowseSwp);

            _derived.Dock = DockStyle.Fill;
            layout.Controls.Add(_derived, 1, 1);
            layout.SetColumnSpan(_derived, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));   // two wrapped lines of path

            var plantPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, WrapContents = true };
            plantPanel.Controls.Add(_msc);
            plantPanel.Controls.Add(_ispc);
            plantPanel.Controls.Add(_ups);
            plantPanel.Controls.Add(_safety);
            plantPanel.SetFlowBreak(_safety, true);   // line 1 = plant radios, line 2 = parameters
            plantPanel.Controls.Add(new Label { Text = "first line/system", AutoSize = true, Padding = new Padding(0, 5, 0, 0) });
            plantPanel.Controls.Add(_firstLine);
            plantPanel.Controls.Add(new Label { Text = "group", AutoSize = true, Padding = new Padding(10, 5, 0, 0) });
            plantPanel.Controls.Add(_group);
            plantPanel.Controls.Add(new Label { Text = "2nd language code", AutoSize = true, Padding = new Padding(10, 5, 0, 0) });
            plantPanel.Controls.Add(_secondLangCode);
            plantPanel.Controls.Add(new Label
            {
                Text = "2 letters, the glossary column (e.g. PT, FR, IT)",
                AutoSize = true,
                Padding = new Padding(4, 5, 0, 0),
                ForeColor = SystemColors.GrayText,
            });
            layout.Controls.Add(new Label { Text = "Plant type", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 2);
            layout.Controls.Add(plantPanel, 1, 2);
            layout.SetColumnSpan(plantPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));   // two flow lines: radios · parameters

            AddPathRow(layout, 3, "IO list", _ioList, BrowseIoList);

            layout.Controls.Add(new Label { Text = "Card sheets", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 4);
            _sheets.Dock = DockStyle.Fill;
            layout.Controls.Add(_sheets, 1, 4);
            layout.SetColumnSpan(_sheets, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            AddOptionalCsvRow(layout, 5, "Plant_Inputs", _plantInputs);
            AddOptionalCsvRow(layout, 6, "Busfield", _busfield);
            AddOptionalCsvRow(layout, 7, "Dest. params", _destinationParams);
            AddOptionalCsvRow(layout, 8, "Sensor (UPS)", _sensor);
            AddOptionalCsvRow(layout, 9, "Strobe (UPS)", _strobe);
            AddOptionalWorkbookRow(layout, 10, "Glossary", _glossary1);
            AddOptionalWorkbookRow(layout, 11, "Glossary ombra", _glossary2);
            AddOptionalWorkbookRow(layout, 12, "Reference SWP", _referenceSwp);
            AddAlarmStringsRow(layout, 13);

            AddTextRow(layout, 14, "Include regex", _includeRegex, "extra mnemonics accepted as diagnostics");
            AddTextRow(layout, 15, "Exclude regex", _excludeRegex, "accepted mnemonics dropped anyway");

            var optionsPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, Dock = DockStyle.Fill, WrapContents = false };
            optionsPanel.Controls.Add(_trustConfig);
            optionsPanel.Controls.Add(_rewriteFormulas);
            optionsPanel.Controls.Add(_glossaryAlways);
            // The mismatch treatment is one exclusive choice; the warning fires regardless and
            // is suppressed per row with the [Mismatch OK] override.
            optionsPanel.Controls.Add(new Label { Text = "On FU/Location mismatch (always warns; per-row [Mismatch OK] silences it):", AutoSize = true, Padding = new Padding(0, 4, 0, 0) });
            var mismatchPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Margin = new Padding(16, 0, 0, 0) };
            mismatchPanel.Controls.Add(_mismatchAppend);
            mismatchPanel.Controls.Add(_mismatchOverwrite);
            optionsPanel.Controls.Add(mismatchPanel);
            layout.Controls.Add(new Label { Text = "Write rules", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 16);
            layout.Controls.Add(optionsPanel, 1, 16);
            layout.SetColumnSpan(optionsPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 132));

            var applyPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, WrapContents = false };
            applyPanel.Controls.Add(_inPlace);
            applyPanel.Controls.Add(_versionFolder);
            layout.Controls.Add(new Label { Text = "Apply writes", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 17);
            layout.Controls.Add(applyPanel, 1, 17);
            layout.SetColumnSpan(applyPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            _hint.Dock = DockStyle.Fill;
            layout.Controls.Add(_hint, 1, 18);
            layout.SetColumnSpan(_hint, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 88 };
            _ok.Width = 130;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_ok);
            layout.Controls.Add(buttons, 1, 19);
            layout.SetColumnSpan(buttons, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            Controls.Add(layout);
            AcceptButton = _ok;
            CancelButton = cancel;

            _sheets.ItemCheck += (s, e) => BeginInvoke(new Action(Revalidate));
            foreach (TextBox box in new[] { _plantInputs, _busfield, _destinationParams, _sensor, _strobe,
                _glossary1, _glossary2, _referenceSwp, _alarmStrings, _includeRegex, _excludeRegex })
                box.TextChanged += (s, e) => Revalidate();
            FormClosing += OnClosingCollect;

            // Free 2-letter code — any glossary column works, listed on _Languages or not. The
            // second row of the global table is the usual second language, so its code is the
            // default for a fresh system.
            List<(string Name, string Code, string Text)> languages = SwpBridgeRegistry.GetLanguages();
            _secondLangCode.Text = languages.Count > 1 ? languages[1].Code : "PT";

            string lastRoot = SwpBridgeRegistry.GetLastRoot();
            if (!string.IsNullOrEmpty(lastRoot) && Directory.Exists(lastRoot))
                TryPrefillFromRoot(lastRoot);
            Revalidate();
        }

        // ---- rows ---------------------------------------------------------------------------

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

        private void AddOptionalWorkbookRow(TableLayoutPanel layout, int row, string label, TextBox box)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            box.Dock = DockStyle.Fill;
            layout.Controls.Add(box, 1, row);
            var button = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            button.Click += (s, e) =>
            {
                using (var dialog = new OpenFileDialog { Filter = "Glossary workbook|*.xlsx;*.xlsm|All files|*.*", CheckFileExists = true })
                {
                    string start = File.Exists(box.Text) ? box.Text : _glossary1.Text;
                    if (File.Exists(start))
                        try { dialog.InitialDirectory = Path.GetDirectoryName(start); } catch { }
                    if (dialog.ShowDialog(this) == DialogResult.OK) box.Text = dialog.FileName;
                }
            };
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        }

        private void AddAlarmStringsRow(TableLayoutPanel layout, int row)
        {
            layout.Controls.Add(new Label { Text = "Alarm strings", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            _alarmStrings.Dock = DockStyle.Fill;
            layout.Controls.Add(_alarmStrings, 1, row);
            var button = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            button.Click += (s, e) =>
            {
                using (var dialog = new OpenFileDialog
                {
                    Filter = "Exported alarms|Alarms_*.txt;*.txt|All files|*.*",
                    CheckFileExists = true,
                    Multiselect = true,
                    Title = "Select the exported Alarms_*.txt files (multi-select)",
                })
                {
                    string first = FirstAlarmPath();
                    if (File.Exists(first))
                        try { dialog.InitialDirectory = Path.GetDirectoryName(first); } catch { }
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                        _alarmStrings.Text = string.Join("; ", dialog.FileNames);
                }
            };
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        }

        private string FirstAlarmPath()
        {
            foreach (string part in _alarmStrings.Text.Split(';'))
                if (part.Trim().Length > 0) return part.Trim();
            return "";
        }

        private List<string> AlarmStringsPaths()
        {
            var paths = new List<string>();
            foreach (string part in _alarmStrings.Text.Split(';'))
                if (part.Trim().Length > 0) paths.Add(part.Trim());
            return paths;
        }

        private void AddOptionalCsvRow(TableLayoutPanel layout, int row, string label, TextBox box)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            box.Dock = DockStyle.Fill;
            layout.Controls.Add(box, 1, row);
            var button = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            button.Click += (s, e) =>
            {
                using (var dialog = new OpenFileDialog { Filter = "CSV|*.csv|All files|*.*", CheckFileExists = true })
                {
                    string start = File.Exists(box.Text) ? box.Text : _plantInputs.Text;
                    if (File.Exists(start))
                        try { dialog.InitialDirectory = Path.GetDirectoryName(start); } catch { }
                    if (dialog.ShowDialog(this) == DialogResult.OK) box.Text = dialog.FileName;
                }
            };
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        }

        private void AddTextRow(TableLayoutPanel layout, int row, string label, TextBox box, string hint)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0) };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            box.Dock = DockStyle.Fill;
            panel.Controls.Add(box, 0, 0);
            panel.Controls.Add(new Label
            {
                Text = hint,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
            }, 1, 0);
            layout.Controls.Add(panel, 1, row);
            layout.SetColumnSpan(panel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        }

        // ---- picking + prefill --------------------------------------------------------------

        private void BrowseSwp()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "SWP workbook|*.xlsx;*.xlsm|All files|*.*",
                CheckFileExists = true,
                FileName = _swp.Text,
            })
            {
                if (File.Exists(_swp.Text))
                    try { dialog.InitialDirectory = Path.GetDirectoryName(_swp.Text); } catch { }
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                if (dialog.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    _hint.Text = "The old .xls format cannot be patched — save the SWP as .xlsx first.";
                    return;
                }

                _swp.Text = dialog.FileName;
                DeriveFromSwpPath();
                Revalidate();
            }
        }

        private void DeriveFromSwpPath()
        {
            try
            {
                string versionFolderPath = Path.GetDirectoryName(_swp.Text);
                SwpRootFolder = Path.GetDirectoryName(versionFolderPath) ?? "";
                string versionFolder = Path.GetFileName(versionFolderPath);
                string config = "";
                foreach (string candidate in Directory.GetFiles(versionFolderPath, "CONFIG_PLANT_*.xls*"))
                {
                    config = Path.GetFileName(candidate);
                    break;
                }
                _derived.Text = "root " + SwpRootFolder + "   ·   version " + versionFolder +
                    (config.Length > 0 ? "   ·   config " + config : "   ·   (no CONFIG_PLANT here)");

                TryPrefillFromRoot(SwpRootFolder);
            }
            catch (Exception ex)
            {
                _derived.Text = "could not derive the system root: " + ex.Message;
            }
        }

        private void TryPrefillFromRoot(string root)
        {
            SwpRootFolder = root;
            try
            {
                _loaded = SwpBridgeSidecarStore.Load(root);
            }
            catch (Exception ex)
            {
                _hint.Text = ex.Message;
                _loaded = null;
                return;
            }
            if (_loaded == null) return;

            SwpBridgeSettings s = _loaded.Settings;
            if (_swp.Text.Length == 0 && s.SwpFileName.Length > 0)
            {
                string candidate = Path.Combine(root, s.CurrentVersionFolder ?? "", s.SwpFileName);
                if (File.Exists(candidate))
                {
                    _swp.Text = candidate;
                    _derived.Text = "root " + root + "   ·   version " + s.CurrentVersionFolder + "   (from the sidecar)";
                }
            }
            _msc.Checked = s.PlantType == SwpPlantType.Msc;
            _ispc.Checked = s.PlantType == SwpPlantType.Ispc;
            _ups.Checked = s.PlantType == SwpPlantType.Ups;
            _safety.Checked = s.PlantType == SwpPlantType.Safety;
            if (s.FirstLineId >= _firstLine.Minimum && s.FirstLineId <= _firstLine.Maximum) _firstLine.Value = s.FirstLineId;
            if (s.GroupIndex >= _group.Minimum && s.GroupIndex <= _group.Maximum) _group.Value = s.GroupIndex;
            if (File.Exists(s.IoListPath))
            {
                _ioList.Text = s.IoListPath;
                LoadSheetList(s.CardSheets);
            }
            _plantInputs.Text = s.PlantInputsCsvPath;
            _busfield.Text = s.BusfieldCsvPath;
            _destinationParams.Text = s.DestinationParamsCsvPath;
            _sensor.Text = s.SensorCsvPath;
            _strobe.Text = s.StrobeDeviceCsvPath;
            _glossary1.Text = s.GlossaryPath1;
            _glossary2.Text = s.GlossaryPath2;
            _referenceSwp.Text = s.ReferenceSwpPath;
            _alarmStrings.Text = string.Join("; ", s.AlarmStringsPaths);
            _glossaryAlways.Checked = s.GlossaryTranslateAlways;
            _includeRegex.Text = s.IncludeRegex;
            _excludeRegex.Text = s.ExcludeRegex;
            _trustConfig.Checked = s.TrustConfigPlant;
            _rewriteFormulas.Checked = s.UpdateDescriptionFormulas;
            _mismatchAppend.Checked = s.MismatchMode == SwpMismatchMode.AppendToDescription;
            _mismatchOverwrite.Checked = s.MismatchMode == SwpMismatchMode.Overwrite;
            _inPlace.Checked = !s.ApplyToVersionFolder;
            _versionFolder.Checked = s.ApplyToVersionFolder;
            if (s.SecondLanguage.Length > 0)
                _secondLangCode.Text = SwpBridgeBuilder.NormalizeLanguageCode(s.SecondLanguage);
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
                LoadSheetList(null);
                Revalidate();
            }
        }

        private void LoadSheetList(List<string> checkedNames)
        {
            _sheets.Items.Clear();
            try
            {
                var wanted = checkedNames == null
                    ? null
                    : new HashSet<string>(checkedNames, StringComparer.OrdinalIgnoreCase);
                foreach (string name in ListSheetNames(_ioList.Text))
                    if (IoListLogic.TryParseCardSheetName(name, out _, out _, out _))
                        _sheets.Items.Add(name, wanted == null || wanted.Contains(name));

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

        // ---- validation + collect -----------------------------------------------------------

        private void Revalidate()
        {
            var problems = new List<string>();
            if (!File.Exists(_swp.Text)) problems.Add("pick the SWP workbook");
            if (!File.Exists(_ioList.Text)) problems.Add("pick the IO-list workbook");
            if (_sheets.CheckedItems.Count == 0) problems.Add("tick at least one card sheet");

            foreach (var pair in new[]
            {
                (_plantInputs, "Plant_Inputs"), (_busfield, "Busfield"),
                (_destinationParams, "Destination_Parameters"), (_sensor, "Sensor"), (_strobe, "StrobeDevice"),
                (_glossary1, "Glossary"), (_glossary2, "Glossary ombra"), (_referenceSwp, "Reference SWP"),
            })
                if (pair.Item1.Text.Trim().Length > 0 && !File.Exists(pair.Item1.Text.Trim()))
                    problems.Add("the " + pair.Item2 + " path does not exist");

            foreach (string path in AlarmStringsPaths())
                if (!File.Exists(path))
                {
                    problems.Add("an alarm-strings path does not exist");
                    break;
                }

            if (!IsValidRegex(_includeRegex.Text)) problems.Add("the include regex does not parse");
            if (!IsValidRegex(_excludeRegex.Text)) problems.Add("the exclude regex does not parse");

            _hint.Text = problems.Count == 0 ? "" : "· " + string.Join("   · ", problems);
            _ok.Enabled = problems.Count == 0;
        }

        private void OnClosingCollect(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK) return;

            SwpSidecarDocument document = _loaded ?? new SwpSidecarDocument();
            if (_loaded == null)
                document.StrobeTemplates = SwpDefaults.DefaultStrobeTemplates();

            SwpBridgeSettings s = document.Settings;
            s.PlantType = _ispc.Checked ? SwpPlantType.Ispc
                : _ups.Checked ? SwpPlantType.Ups
                : _safety.Checked ? SwpPlantType.Safety
                : SwpPlantType.Msc;
            s.CurrentVersionFolder = Path.GetFileName(Path.GetDirectoryName(_swp.Text)) ?? "";
            s.SwpFileName = Path.GetFileName(_swp.Text);
            s.ConfigFileName = FindConfigName(Path.GetDirectoryName(_swp.Text));
            s.IoListPath = _ioList.Text;
            s.CardSheets = new List<string>();
            foreach (object item in _sheets.CheckedItems) s.CardSheets.Add((string)item);
            s.PlantInputsCsvPath = _plantInputs.Text.Trim();
            s.BusfieldCsvPath = _busfield.Text.Trim();
            s.DestinationParamsCsvPath = _destinationParams.Text.Trim();
            s.SensorCsvPath = _sensor.Text.Trim();
            s.StrobeDeviceCsvPath = _strobe.Text.Trim();
            s.GlossaryPath1 = _glossary1.Text.Trim();
            s.GlossaryPath2 = _glossary2.Text.Trim();
            s.ReferenceSwpPath = _referenceSwp.Text.Trim();
            s.AlarmStringsPaths = AlarmStringsPaths();
            s.GlossaryTranslateAlways = _glossaryAlways.Checked;
            s.FirstLineId = (int)_firstLine.Value;
            s.GroupIndex = (int)_group.Value;
            s.IncludeRegex = _includeRegex.Text.Trim();
            s.ExcludeRegex = _excludeRegex.Text.Trim();
            s.TrustConfigPlant = _trustConfig.Checked;
            s.UpdateDescriptionFormulas = _rewriteFormulas.Checked;
            s.MismatchMode = _mismatchOverwrite.Checked
                ? SwpMismatchMode.Overwrite
                : SwpMismatchMode.AppendToDescription;
            s.ApplyToVersionFolder = _versionFolder.Checked;
            s.SecondLanguage = _secondLangCode.Text.Trim().ToUpperInvariant();

            Result = document;
            SwpBridgeRegistry.SetLastRoot(SwpRootFolder);
        }

        private static string FindConfigName(string folder)
        {
            try
            {
                foreach (string candidate in Directory.GetFiles(folder ?? "", "CONFIG_PLANT_*.xls*"))
                    return Path.GetFileName(candidate);
            }
            catch { }
            return "";
        }

        private static bool IsValidRegex(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern)) return true;
            try
            {
                _ = new System.Text.RegularExpressions.Regex(pattern.Trim());
                return true;
            }
            catch { return false; }
        }
    }
}
