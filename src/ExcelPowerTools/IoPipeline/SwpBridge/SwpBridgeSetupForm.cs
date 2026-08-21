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
    /// Workspace setup, in five sections: SWP SYSTEM (the file, its base, the plant), IO LIST,
    /// PRODUCTION FILES (one realtime root resolved live by the PRODUCTION-tree layout),
    /// INTERNATIONALIZATION (glossaries, second language) and RULES. Picking the SWP file inside
    /// a dated version folder derives the system root (its parent) and the current version; a
    /// sidecar found at that root pre-fills every other field, so returning to a system is one
    /// pick + OK. Every production input is optional — a missing one keeps its check off. The
    /// .xls format is refused (the headless patcher writes OpenXml); .xlsx/.xlsm both work.
    /// Rows whose content varies (plant numbers, the resolved list, the write rules) size
    /// themselves; the fixed ones derive from the font, which is one point above the system's.
    /// </summary>
    internal sealed class SwpBridgeSetupForm : Form
    {
        // SWP SYSTEM
        private readonly TextBox _swp = new TextBox { ReadOnly = true };
        private readonly Label _derived = new Label { AutoSize = false, ForeColor = SystemColors.GrayText };
        private readonly TextBox _referenceSwp = new TextBox();
        private readonly RadioButton _msc = new RadioButton { Text = "MSC (sorter)", AutoSize = true, Checked = true };
        private readonly RadioButton _ispc = new RadioButton { Text = "ISPC (induction)", AutoSize = true };
        private readonly RadioButton _ups = new RadioButton { Text = "UPS (upstream)", AutoSize = true };
        private readonly RadioButton _safety = new RadioButton { Text = "SAFETY (PLC)", AutoSize = true };
        private readonly Label _firstLineLabel = new Label { Text = "first line/system", AutoSize = true, Padding = new Padding(0, 5, 0, 0) };
        private readonly NumericUpDown _firstLine = new NumericUpDown { Minimum = 1, Maximum = 99, Value = 1, Width = 60 };
        private readonly Label _groupLabel = new Label { Text = "group", AutoSize = true, Padding = new Padding(10, 5, 0, 0) };
        private readonly NumericUpDown _group = new NumericUpDown { Minimum = 1, Maximum = 99, Value = 1, Width = 60 };
        private readonly Label _plantNote = new Label
        {
            Text = "ISPC: first induction line id · line group — UPS: first system id · conveyor group",
            AutoSize = true,
            Padding = new Padding(10, 5, 0, 0),
            ForeColor = SystemColors.GrayText,
        };

        // IO LIST
        private readonly TextBox _ioList = new TextBox { ReadOnly = true };
        private readonly CheckedListBox _sheets = new CheckedListBox { CheckOnClick = true, IntegralHeight = false };

        // PRODUCTION FILES
        private readonly TextBox _productionRoot = new TextBox();
        private readonly FlowLayoutPanel _resolved = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(3, 2, 3, 6),
        };
        private bool _rootSuggested;   // the root came from a legacy sidecar's Busfield path
        private static readonly System.Drawing.Color ResolvedGreen = System.Drawing.Color.FromArgb(0, 140, 60);
        private static readonly System.Drawing.Color ResolvedRed = System.Drawing.Color.FromArgb(200, 30, 30);

        // INTERNATIONALIZATION
        private readonly TextBox _glossary1 = new TextBox();
        private readonly TextBox _glossary2 = new TextBox();
        private readonly TextBox _secondLangCode = new TextBox { Width = 48, CharacterCasing = CharacterCasing.Upper, MaxLength = 2 };

        // RULES
        private readonly TextBox _includeRegex = new TextBox();
        private readonly TextBox _excludeRegex = new TextBox();
        private readonly CheckBox _trustConfig = new CheckBox { Text = "Never overwrite CONFIG_PLANT descriptions", AutoSize = true };
        private readonly CheckBox _rewriteFormulas = new CheckBox { Text = "Rewrite the Q/R description formulas wholesale", AutoSize = true };
        private readonly CheckBox _glossaryAlways = new CheckBox { Text = "Glossary always wins", AutoSize = true };
        private readonly RadioButton _mismatchAppend = new RadioButton { Text = "append identity to descriptions", AutoSize = true, Checked = true };
        private readonly RadioButton _mismatchOverwrite = new RadioButton { Text = "overwrite the row's FU/Loc/Device cells", AutoSize = true };
        private readonly RadioButton _inPlace = new RadioButton { Text = "in place (with .bak)", AutoSize = true, Checked = true };
        private readonly RadioButton _versionFolder = new RadioButton { Text = "new dated version folder", AutoSize = true };

        private readonly Label _hint = new Label { AutoSize = false, ForeColor = System.Drawing.Color.FromArgb(156, 0, 6) };
        private readonly Button _ok = new Button { Text = "Build workspace", DialogResult = DialogResult.OK, Enabled = false };
        private readonly ToolTip _tips = new ToolTip { AutoPopDelay = 30000, InitialDelay = 400, ReshowDelay = 100, ShowAlways = true };

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
            Font = new System.Drawing.Font(SystemFonts.MessageBoxFont.FontFamily,
                SystemFonts.MessageBoxFont.SizeInPoints + 1f, FontStyle.Regular, GraphicsUnit.Point);
            ClientSize = new Size(780, 960);
            MinimumSize = new Size(720, 840);

            int line = Font.Height;
            int pathRow = line + 14;
            int textRow = line + 12;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 3,
                RowCount = 22,
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));

            int row = 0;

            // ---- SWP SYSTEM ----------------------------------------------------------------
            AddSection(layout, row++, "SWP SYSTEM");
            AddPathRow(layout, row++, "SWP file", _swp, BrowseSwp, pathRow);

            _derived.Dock = DockStyle.Fill;
            layout.Controls.Add(_derived, 1, row);
            layout.SetColumnSpan(_derived, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 2 * line + 6));   // two wrapped lines of path
            row++;

            AddOptionalWorkbookRow(layout, row++, "Reference SWP (base)", _referenceSwp,
                "Reference SWP|*.xlsx;*.xlsm|All files|*.*",
                "A known-good SWP of the same plant type. The formula-structure audit compares every " +
                "block's H/I/Q/R formulas against it and queues repairs for the ones that drifted.", textRow);

            var plantPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(3, 4, 3, 2),
            };
            plantPanel.Controls.Add(_msc);
            plantPanel.Controls.Add(_ispc);
            plantPanel.Controls.Add(_ups);
            plantPanel.Controls.Add(_safety);
            plantPanel.SetFlowBreak(_safety, true);   // line 1 = plant radios, line 2 = the ISPC/UPS numbers
            plantPanel.Controls.Add(_firstLineLabel);
            plantPanel.Controls.Add(_firstLine);
            plantPanel.Controls.Add(_groupLabel);
            plantPanel.Controls.Add(_group);
            plantPanel.Controls.Add(_plantNote);
            layout.Controls.Add(new Label { Text = "Plant type", TextAlign = ContentAlignment.TopLeft, Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 0) }, 0, row);
            layout.Controls.Add(plantPanel, 1, row);
            layout.SetColumnSpan(plantPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row++;

            // ---- IO LIST -------------------------------------------------------------------
            AddSection(layout, row++, "IO LIST");
            AddPathRow(layout, row++, "IO list", _ioList, BrowseIoList, pathRow);

            layout.Controls.Add(new Label { Text = "Card sheets", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            _sheets.Dock = DockStyle.Fill;
            layout.Controls.Add(_sheets, 1, row);
            layout.SetColumnSpan(_sheets, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            row++;

            // ---- PRODUCTION FILES ----------------------------------------------------------
            AddSection(layout, row++, "PRODUCTION FILES");
            layout.Controls.Add(new Label { Text = "Production root", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            _productionRoot.Dock = DockStyle.Fill;
            layout.Controls.Add(_productionRoot, 1, row);
            var rootButton = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            rootButton.Click += (s, e) => BrowseProductionRoot();
            layout.Controls.Add(rootButton, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, pathRow));
            _tips.SetToolTip(_productionRoot,
                "The realtime PRODUCTION folder — or the backup folder that contains it. Resolved by the " +
                "real tree layout: IODriver\\BusField.csv for every plant; MSC\\DB_Csv\\ (sorter) or " +
                "ISPC\\db_csv\\ (induction) Plant_Inputs; the AlarmLists folder; ConveyorSys\\SysGroupNN\\ " +
                "for upstream. Flat copies directly in the folder work too.");
            row++;

            _resolved.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            layout.Controls.Add(_resolved, 1, row);
            layout.SetColumnSpan(_resolved, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row++;

            // ---- INTERNATIONALIZATION ------------------------------------------------------
            AddSection(layout, row++, "INTERNATIONALIZATION");
            AddOptionalWorkbookRow(layout, row++, "Glossary", _glossary1, "Glossary workbook|*.xlsx;*.xlsm|All files|*.*",
                "The standard glossario: translation pairs, the bus-failure phrase family and the T.S. cards (\"Numero scheda\").", textRow);
            AddOptionalWorkbookRow(layout, row++, "Glossary ombra", _glossary2, "Glossary workbook|*.xlsx;*.xlsm|All files|*.*",
                "The project shadow glossary — its pairs add to (and win over) the standard one.", textRow);
            AddTextRow(layout, row++, "2nd language", _secondLangCode,
                "2 letters — the glossary column used for Description 2 (e.g. PT, FR, IT)", textRow, boxWidth: 48);

            // ---- RULES ---------------------------------------------------------------------
            AddSection(layout, row++, "RULES");
            AddTextRow(layout, row++, "Include regex", _includeRegex, "extra mnemonics accepted as diagnostics", textRow);
            AddTextRow(layout, row++, "Exclude regex", _excludeRegex, "accepted mnemonics dropped anyway", textRow);

            var optionsPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(3, 4, 3, 2),
            };
            optionsPanel.Controls.Add(_trustConfig);
            optionsPanel.Controls.Add(_rewriteFormulas);
            optionsPanel.Controls.Add(_glossaryAlways);
            // The mismatch treatment is one exclusive choice; the warning fires regardless and
            // is suppressed per row with the [Mismatch OK] override.
            optionsPanel.Controls.Add(new Label { Text = "On FU/Location mismatch (always warns; per-row [Mismatch OK] silences it):", AutoSize = true, Padding = new Padding(0, 4, 0, 0) });
            var mismatchPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                Margin = new Padding(16, 0, 0, 0),
            };
            mismatchPanel.Controls.Add(_mismatchAppend);
            mismatchPanel.Controls.Add(_mismatchOverwrite);
            optionsPanel.Controls.Add(mismatchPanel);
            layout.Controls.Add(new Label { Text = "Write rules", TextAlign = ContentAlignment.TopLeft, Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 0) }, 0, row);
            layout.Controls.Add(optionsPanel, 1, row);
            layout.SetColumnSpan(optionsPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row++;

            _tips.SetToolTip(_trustConfig,
                "SCADA shows CONFIG_PLANT's template text followed by the SWP's description cell, so writing the " +
                "IO text there stacks a second sentence onto a template that often already says it. Ticked: rows " +
                "this bridge generated (photocells, strobes) and \"P*\"-typed module rows keep CONFIG's wording — " +
                "the bridge adds only the FLD label, and nothing at all when there is no mismatch. Per row, the " +
                "[Keep CONFIG text] column decides outright.");
            _tips.SetToolTip(_rewriteFormulas,
                "Replaces the Q/R final-text formulas with a full CONCAT rewrite. Off: the pending append only extends them to show the description cell.");
            _tips.SetToolTip(_glossaryAlways,
                "The glossary translation replaces the IO list's language-2 text even when one exists; unchecked, it only fills empty gaps.");
            _tips.SetToolTip(_mismatchAppend,
                "The IO LIST's identity — functional unit + location + device, compacted — is appended to both " +
                "descriptions; the device cell (J) is left blank because the label already carries it; a " +
                "template-only Q/R formula is extended so the appended text actually reaches SCADA.");
            _tips.SetToolTip(_mismatchOverwrite,
                "The IO list's identity is written into the row's OWN cells — functional unit (H), location (I), " +
                "device (J) — leaving the descriptions clean. Cells that are shared formulas cannot be patched " +
                "headlessly: those rows fall back to the append, so the FLD is never lost either way.");

            var applyPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, WrapContents = false };
            applyPanel.Controls.Add(_inPlace);
            applyPanel.Controls.Add(_versionFolder);
            layout.Controls.Add(new Label { Text = "Apply writes", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            layout.Controls.Add(applyPanel, 1, row);
            layout.SetColumnSpan(applyPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, pathRow));
            row++;
            _tips.SetToolTip(_inPlace, "Patches the SWP where it is; the previous file stays beside it as .bak.");
            _tips.SetToolTip(_versionFolder, "Copies SWP + CONFIG_PLANT into the next NN_dd-MMM-yyyy sibling folder and patches the copy.");

            _hint.Dock = DockStyle.Fill;
            layout.Controls.Add(_hint, 1, row);
            layout.SetColumnSpan(_hint, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 2 * line + 8));
            row++;

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 92, Height = line + 14 };
            _ok.Width = 140;
            _ok.Height = line + 14;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_ok);
            layout.Controls.Add(buttons, 1, row);
            layout.SetColumnSpan(buttons, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, line + 24));

            Controls.Add(layout);
            ExcelPowerTools.UI.GdiText.Apply(this);   // GDI text — the resolved lines added later follow too
            AcceptButton = _ok;
            CancelButton = cancel;

            // ItemCheck also fires for Items.Add(name, isChecked: true) — during the sidecar
            // prefill in this constructor, BEFORE the handle exists. BeginInvoke would throw
            // there and abort the check before it is stored (every card showed unticked).
            _sheets.ItemCheck += (s, e) => { if (IsHandleCreated) BeginInvoke(new Action(Revalidate)); };
            foreach (TextBox box in new[] { _glossary1, _glossary2, _referenceSwp, _includeRegex, _excludeRegex })
                box.TextChanged += (s, e) => Revalidate();
            _productionRoot.TextChanged += (s, e) => { _rootSuggested = false; UpdateResolved(); Revalidate(); };
            foreach (RadioButton radio in new[] { _msc, _ispc, _ups, _safety })
                radio.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) UpdatePlantUi(); };
            _group.ValueChanged += (s, e) => UpdateResolved();
            FormClosing += OnClosingCollect;
            Load += (s, e) => EnsureCardListHeight();

            // Long paths: the telling part is the end, so every path box shows its tail.
            foreach (TextBox box in PathBoxes())
                box.HandleCreated += (s, e) => ScrollToEnd((TextBox)s);
            Shown += (s, e) => ScrollPathsToEnd();

            // Free 2-letter code — any glossary column works, listed on _Languages or not. The
            // second row of the global table is the usual second language, so its code is the
            // default for a fresh system.
            List<(string Name, string Code, string Text)> languages = SwpBridgeRegistry.GetLanguages();
            _secondLangCode.Text = languages.Count > 1 ? languages[1].Code : "PT";

            string lastRoot = SwpBridgeRegistry.GetLastRoot();
            if (!string.IsNullOrEmpty(lastRoot) && Directory.Exists(lastRoot))
                TryPrefillFromRoot(lastRoot);
            UpdatePlantUi();
            Revalidate();
        }

        // ---- long paths show their tail -----------------------------------------------------

        private TextBox[] PathBoxes() => new[] { _swp, _referenceSwp, _ioList, _productionRoot, _glossary1, _glossary2 };

        /// <summary>Caret at the end and the view scrolled there — applied whenever code fills a
        /// path box (prefill, browse, first show); typing is left alone.</summary>
        private static void ScrollToEnd(TextBox box)
        {
            try
            {
                box.SelectionStart = box.Text.Length;
                box.SelectionLength = 0;
                box.ScrollToCaret();
            }
            catch { }
        }

        private void ScrollPathsToEnd()
        {
            foreach (TextBox box in PathBoxes()) ScrollToEnd(box);
        }

        // ---- rows ---------------------------------------------------------------------------

        private void AddSection(TableLayoutPanel layout, int row, string title)
        {
            var label = new Label
            {
                Text = title,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                ForeColor = SystemColors.GrayText,
                Font = new System.Drawing.Font(Font, FontStyle.Bold),
            };
            layout.Controls.Add(label, 0, row);
            layout.SetColumnSpan(label, 3);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, Font.Height + 12));
        }

        private void AddPathRow(TableLayoutPanel layout, int row, string label, TextBox box, Action browse, int height)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            box.Dock = DockStyle.Fill;
            layout.Controls.Add(box, 1, row);
            var button = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            button.Click += (s, e) => browse();
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        }

        private void AddOptionalWorkbookRow(TableLayoutPanel layout, int row, string label, TextBox box, string filter, string tip, int height)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            box.Dock = DockStyle.Fill;
            layout.Controls.Add(box, 1, row);
            if (!string.IsNullOrEmpty(tip)) _tips.SetToolTip(box, tip);
            var button = new Button { Text = "Browse…", Dock = DockStyle.Fill };
            button.Click += (s, e) =>
            {
                using (var dialog = new OpenFileDialog { Filter = filter, CheckFileExists = true })
                {
                    string start = File.Exists(box.Text) ? box.Text : File.Exists(_glossary1.Text) ? _glossary1.Text : _swp.Text;
                    if (File.Exists(start))
                        try { dialog.InitialDirectory = Path.GetDirectoryName(start); } catch { }
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        box.Text = dialog.FileName;
                        ScrollToEnd(box);
                    }
                }
            };
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        }

        private void AddTextRow(TableLayoutPanel layout, int row, string label, TextBox box, string hint, int height, int boxWidth = 220)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0) };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, boxWidth));
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
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        }

        // ---- plant-dependent UI -------------------------------------------------------------

        private SwpPlantType SelectedPlantType()
        {
            return _ispc.Checked ? SwpPlantType.Ispc
                : _ups.Checked ? SwpPlantType.Ups
                : _safety.Checked ? SwpPlantType.Safety
                : SwpPlantType.Msc;
        }

        /// <summary>The line/group numbers only mean something for ISPC and UPS; the resolved
        /// production list changes with the plant (and, for UPS, with the group).</summary>
        private void UpdatePlantUi()
        {
            bool numbered = _ispc.Checked || _ups.Checked;
            _firstLineLabel.Visible = numbered;
            _firstLine.Visible = numbered;
            _groupLabel.Visible = numbered;
            _group.Visible = numbered;
            _plantNote.Visible = numbered;
            UpdateResolved();
        }

        /// <summary>The resolved list: ✓ green per input found, ✗ red per input missing — only the
        /// inputs this plant type consumes are mentioned at all.</summary>
        private void UpdateResolved()
        {
            _resolved.SuspendLayout();
            var old = new List<System.Windows.Forms.Control>();
            foreach (System.Windows.Forms.Control control in _resolved.Controls) old.Add(control);
            _resolved.Controls.Clear();
            foreach (System.Windows.Forms.Control control in old) control.Dispose();

            string root = _productionRoot.Text.Trim();
            if (root.Length == 0)
            {
                if (!AddLegacyPathLines())
                    AddResolvedLine("", ForeColor,
                        "No root picked — the plant checks (Plant_Inputs, Busfield, BUS_ExtendedDiagnostic, alarm strings) and the generators stay off.",
                        muted: true);
            }
            else if (!Directory.Exists(root))
            {
                AddResolvedLine("✗", ResolvedRed, "The folder does not exist.");
            }
            else
            {
                SwpProductionFiles files = SwpProductionTree.Resolve(root, SelectedPlantType(), (int)_group.Value);
                if (_rootSuggested)
                    AddResolvedLine("↳", SystemColors.GrayText, "suggested from the sidecar's Busfield path — check the list", muted: true);
                if (!SameFolder(files.Root, root))
                    AddResolvedLine("↳", SystemColors.GrayText, "PRODUCTION\\ found inside the picked folder — resolving from there", muted: true);
                foreach (SwpProductionFile file in files.ApplicableFiles)
                {
                    if (file.Found) AddResolvedLine("✓", ResolvedGreen, file.Relative);
                    else AddResolvedLine("✗", ResolvedRed, file.Label + " not found — its checks stay off");
                }
                if (files.AlarmStrings.Count > 0)
                    AddResolvedLine("✓", ResolvedGreen, files.AlarmStringsPattern + " (" + files.AlarmStrings.Count + " files)");
                else
                    AddResolvedLine("✗", ResolvedRed, "Alarms_*.txt not found — the alarm-strings check stays off");
            }
            _resolved.ResumeLayout(true);
            EnsureCardListHeight();
        }

        /// <summary>The card list is the one stretching row; when the self-sizing rows leave it
        /// fewer than eight lines, the window grows (within the screen) to give them back.</summary>
        private void EnsureCardListHeight()
        {
            if (!IsHandleCreated) return;
            int itemHeight = _sheets.ItemHeight > 0 ? _sheets.ItemHeight : Font.Height + 3;
            int wanted = 8 * itemHeight + 8;
            int missing = wanted - _sheets.Height;
            if (missing <= 0) return;

            System.Drawing.Rectangle area = Screen.FromControl(this).WorkingArea;
            int height = Math.Min(area.Height - 24, Height + missing);
            if (height <= Height) return;
            Height = height;
            if (Bottom > area.Bottom) Top = Math.Max(area.Top, area.Bottom - Height);
        }

        private void AddResolvedLine(string glyph, System.Drawing.Color glyphColor, string text, bool muted = false)
        {
            var lineRow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 2),
            };
            lineRow.Controls.Add(new Label
            {
                Text = glyph,
                AutoSize = false,
                Width = Font.Height + 8,
                Height = Font.Height + 2,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = glyphColor,
                Font = new System.Drawing.Font(Font, FontStyle.Bold),
                Margin = new Padding(0),
            });
            lineRow.Controls.Add(new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = muted ? SystemColors.GrayText : ForeColor,
                Margin = new Padding(0, 1, 0, 0),
            });
            _resolved.Controls.Add(lineRow);
        }

        /// <summary>A sidecar from before the production root still carries per-file paths; they
        /// keep working until a root replaces them, and the list says so.</summary>
        private bool AddLegacyPathLines()
        {
            SwpBridgeSettings s = _loaded?.Settings;
            if (s == null) return false;
            var paths = new List<string>();
            foreach (string path in new[] { s.BusfieldCsvPath, s.PlantInputsCsvPath, s.DestinationParamsCsvPath,
                s.BusExtendedDiagnosticCsvPath, s.SensorCsvPath, s.StrobeDeviceCsvPath })
                if (!string.IsNullOrWhiteSpace(path)) paths.Add(path);
            if (paths.Count == 0 && s.AlarmStringsPaths.Count == 0) return false;

            AddResolvedLine("", ForeColor, "per-file paths from the sidecar (pick a root to replace them):", muted: true);
            foreach (string path in paths)
                AddResolvedLine(File.Exists(path) ? "✓" : "✗", File.Exists(path) ? ResolvedGreen : ResolvedRed, path);
            if (s.AlarmStringsPaths.Count > 0)
                AddResolvedLine("✓", ResolvedGreen, s.AlarmStringsPaths.Count + " alarm-strings file(s)");
            return true;
        }

        private static bool SameFolder(string a, string b)
        {
            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\', '/'),
                    Path.GetFullPath(b).TrimEnd('\\', '/'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private void BrowseProductionRoot()
        {
            using (var dialog = new FolderBrowserDialog
            {
                Description = "Pick the PRODUCTION folder — or the backup folder that contains it",
                ShowNewFolderButton = false,
            })
            {
                string start = _productionRoot.Text.Trim();
                if (Directory.Exists(start)) dialog.SelectedPath = start;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _productionRoot.Text = dialog.SelectedPath;
                    ScrollToEnd(_productionRoot);
                }
            }
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
                ScrollToEnd(_swp);
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
                LoadSheetList(s.ExcludedCardSheets);
            }
            _referenceSwp.Text = s.ReferenceSwpPath;
            _glossary1.Text = s.GlossaryPath1;
            _glossary2.Text = s.GlossaryPath2;
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

            string suggested = s.ProductionRoot.Length > 0 ? "" : SuggestRootFromLegacy(s);
            _productionRoot.Text = s.ProductionRoot.Length > 0 ? s.ProductionRoot : suggested;
            _rootSuggested = suggested.Length > 0;
            UpdatePlantUi();
            ScrollPathsToEnd();
        }

        /// <summary>A legacy Busfield path under …\IODriver\ names the root two levels up.</summary>
        private static string SuggestRootFromLegacy(SwpBridgeSettings s)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(s.BusfieldCsvPath)) return "";
                string folder = Path.GetDirectoryName(s.BusfieldCsvPath);
                if (folder == null || !string.Equals(Path.GetFileName(folder), "IODriver", StringComparison.OrdinalIgnoreCase)) return "";
                string root = Path.GetDirectoryName(folder);
                return root != null && Directory.Exists(root) ? root : "";
            }
            catch { return ""; }
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
                ScrollToEnd(_ioList);
                LoadSheetList(null);
                Revalidate();
            }
        }

        /// <summary>Every card sheet starts ticked; only the ones the sidecar recorded as
        /// deliberately unticked stay off — a card added to the IO list since is in by default.</summary>
        private void LoadSheetList(List<string> excludedNames)
        {
            _sheets.Items.Clear();
            try
            {
                var excluded = new HashSet<string>(excludedNames ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                foreach (string name in ListSheetNames(_ioList.Text))
                    if (IoListLogic.TryParseCardSheetName(name, out _, out _, out _))
                        _sheets.Items.Add(name, !excluded.Contains(name));

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

            string root = _productionRoot.Text.Trim();
            if (root.Length > 0 && !Directory.Exists(root)) problems.Add("the production root does not exist");

            foreach (var pair in new[]
            {
                (_referenceSwp, "Reference SWP"), (_glossary1, "Glossary"), (_glossary2, "Glossary ombra"),
            })
                if (pair.Item1.Text.Trim().Length > 0 && !File.Exists(pair.Item1.Text.Trim()))
                    problems.Add("the " + pair.Item2 + " path does not exist");

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
            s.PlantType = SelectedPlantType();
            s.CurrentVersionFolder = Path.GetFileName(Path.GetDirectoryName(_swp.Text)) ?? "";
            s.SwpFileName = Path.GetFileName(_swp.Text);
            s.ConfigFileName = FindConfigName(Path.GetDirectoryName(_swp.Text));
            s.IoListPath = _ioList.Text;
            s.CardSheets = new List<string>();
            s.ExcludedCardSheets = new List<string>();
            for (int i = 0; i < _sheets.Items.Count; i++)
                (_sheets.GetItemChecked(i) ? s.CardSheets : s.ExcludedCardSheets).Add((string)_sheets.Items[i]);

            // The root is the authority: the per-file paths are filled by its resolution at build
            // time. Without a root, a legacy sidecar's per-file paths stay exactly as loaded.
            s.ProductionRoot = _productionRoot.Text.Trim();
            if (s.ProductionRoot.Length > 0)
            {
                s.PlantInputsCsvPath = "";
                s.BusfieldCsvPath = "";
                s.DestinationParamsCsvPath = "";
                s.BusExtendedDiagnosticCsvPath = "";
                s.SensorCsvPath = "";
                s.StrobeDeviceCsvPath = "";
                s.AlarmStringsPaths = new List<string>();
            }

            s.GlossaryPath1 = _glossary1.Text.Trim();
            s.GlossaryPath2 = _glossary2.Text.Trim();
            s.ReferenceSwpPath = _referenceSwp.Text.Trim();
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
