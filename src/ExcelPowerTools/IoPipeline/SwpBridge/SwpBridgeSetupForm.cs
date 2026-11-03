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
    /// Workspace setup, one section per QUESTION rather than one per control type:
    /// SWP SYSTEM (the session, the file, its base, the plant), IO LIST (the workbook, its card
    /// sheets, and the regexes deciding which of its mnemonics are diagnostics at all),
    /// PRODUCTION FILES (one realtime root resolved live by the PRODUCTION-tree layout),
    /// INTERNATIONALIZATION (glossaries, second language, and how the glossary competes with the
    /// IO list), CONFIG_PLANT (what happens to the template's text) and ELECTRICAL_DESIGNATION
    /// (where the functional unit / location / device land).
    ///
    /// The last two have the same shape, because they answer the same KIND of question: one
    /// exclusive default, one destination-sheet regex naming the sheets that take the other mode,
    /// and a per-row column that wins over both.
    ///
    /// Picking the SWP file inside a dated version folder derives the system root (its parent),
    /// the current version and — from the SWP's own tab names — the first-line and group numbers;
    /// a sidecar found at that root pre-fills every other field, so returning to a system is one
    /// pick + OK. A system can also be reached from the other end: "Session" opens a saved
    /// swp-bridge.json directly, and "Save as…" names a new one anywhere, which is what lets one
    /// IO list drive two SWP systems as two named sessions. Opening also recovers a "…json.bak".
    ///
    /// The reference SWP is required; every production input is optional — a missing one keeps
    /// its check off. The .xls format is refused (the headless patcher writes OpenXml);
    /// .xlsx/.xlsm both work. Rows whose content varies (plant numbers, the resolved list, the
    /// write rules) size themselves; the fixed ones derive from the font, one point above the
    /// system's.
    /// </summary>
    internal sealed class SwpBridgeSetupForm : Form
    {
        // SWP SYSTEM
        private readonly TextBox _session = new TextBox { ReadOnly = true };
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
        private readonly Label _sheetsLabel = new Label { Text = "Card sheets", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };

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
        private string _derivedNumbers = "";   // where the line/group numbers were read from
        private string _loadedSwpFileName = "";   // the SWP an OPENED session was saved for
        private string _loadedSessionPath = "";   // …and the file it came from, to spot an overwrite
        private static readonly System.Drawing.Color ResolvedGreen = System.Drawing.Color.FromArgb(0, 140, 60);
        private static readonly System.Drawing.Color ResolvedRed = System.Drawing.Color.FromArgb(200, 30, 30);

        // INTERNATIONALIZATION
        private readonly TextBox _glossary1 = new TextBox();
        private readonly TextBox _glossary2 = new TextBox();
        private readonly TextBox _secondLangCode = new TextBox { Width = 48, CharacterCasing = CharacterCasing.Upper, MaxLength = 2 };

        private readonly CheckBox _glossaryAlways = new CheckBox { Text = "Glossary always wins", AutoSize = true };

        // IO LIST — which mnemonics count as diagnostics at all
        private readonly TextBox _includeRegex = new TextBox();
        private readonly TextBox _excludeRegex = new TextBox();

        // CONFIG_PLANT — one exclusive choice: a row cannot both keep the template whole and cut
        // its tail off. Same three-level shape as the FLD treatment: per-row tick, sheet rule, default.
        private readonly RadioButton _keepConfigText = new RadioButton { Text = "Never overwrite CONFIG_PLANT descriptions", AutoSize = true };
        private readonly RadioButton _replaceTail = new RadioButton { Text = "Replace the template's trailing text with the row description", AutoSize = true, Checked = true };
        private readonly Label _configSheetRuleCaption = new Label { AutoSize = true, Padding = new Padding(0, 4, 6, 0) };
        private readonly TextBox _configSheetRegex = new TextBox { Width = 220 };

        // ELECTRICAL_DESIGNATION
        private readonly RadioButton _mismatchAppend = new RadioButton { Text = "append identity to descriptions", AutoSize = true, Checked = true };
        private readonly RadioButton _mismatchOverwrite = new RadioButton { Text = "overwrite the row's FU/Loc/Device cells", AutoSize = true };
        private readonly Label _fldSheetRuleCaption = new Label { AutoSize = true, Padding = new Padding(0, 4, 6, 0) };
        private readonly TextBox _fldSheetRegex = new TextBox { Width = 220 };
        private readonly RadioButton _inPlace = new RadioButton { Text = "in place (with .bak)", AutoSize = true, Checked = true };
        private readonly RadioButton _versionFolder = new RadioButton { Text = "new dated version folder", AutoSize = true };
        private readonly Label _copyCaption = new Label { Text = "…and copy along the files matching:", AutoSize = true, Padding = new Padding(16, 5, 6, 0) };
        private readonly TextBox _copyRegex = new TextBox { Width = 320 };

        private readonly Label _hint = new Label { AutoSize = false, ForeColor = System.Drawing.Color.FromArgb(156, 0, 6) };
        private readonly Button _ok = new Button { Text = "Build workspace", DialogResult = DialogResult.OK, Enabled = false };
        private readonly ToolTip _tips = new ToolTip { AutoPopDelay = 30000, InitialDelay = 400, ReshowDelay = 100, ShowAlways = true };

        public SwpSidecarDocument Result { get; private set; }
        public string SwpRootFolder { get; private set; } = "";

        /// <summary>The session FILE the build will write decisions to — the root's default name
        /// unless "Save as…" named another.</summary>
        public string SessionPath { get; private set; } = "";

        private SwpSidecarDocument _loaded;   // an existing sidecar at the derived root

        /// <summary>The running workspace a Save as moved to its new file — flushed there if the
        /// dialog is then cancelled, so the file says what is actually running.</summary>
        private SwpBridgeSession _retargetedLive;

        /// <summary>Every .json: a filter naming the exact default file ("swp-bridge.json") hid every
        /// named session. Patterns are separated by ";" — the Windows syntax; "," matches nothing.</summary>
        private const string SessionOpenFilter = "SWP Bridge session|*.json;*.json.bak|All files|*.*";
        private const string SessionSaveFilter = "SWP Bridge session|*.json|All files|*.*";

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
            // WIDTH IS WHAT FIXES THE CROP. The self-sizing rows (the plant numbers, the resolved
            // production list, the CONFIG_PLANT and ELECTRICAL_DESIGNATION panels) wrap their
            // contents; at the original 780 the wrapped lines made those rows tall enough to push
            // the bottom of the form off the window, and the buttons went with it. A scrollbar
            // would hide the shape of the dialog — widening stops the wrapping instead, so
            // everything fits at its natural height. 175% of the original, measured in the field.
            const int NaturalWidth = 780, NaturalMinimum = 720;
            System.Drawing.Rectangle work = Screen.PrimaryScreen.WorkingArea;

            // …but never wider than the desktop: a window running off the screen is its own crop.
            int width = Math.Min((int)(NaturalWidth * 1.75), work.Width - 32);
            MinimumSize = new Size(Math.Min((int)(NaturalMinimum * 1.75), width), 840);
            ClientSize = new Size(width, Math.Min(960, work.Height - 32));

            int line = Font.Height;
            int pathRow = line + 14;
            int textRow = line + 12;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 3,
                RowCount = 23,
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            // The single "Browse…" of every path row. The Session row's buttons do not live here —
            // they sit in a panel spanning this column and the path's, so no other row's button grows
            // to fit them.
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));

            int row = 0;

            // ---- SWP SYSTEM ----------------------------------------------------------------
            AddSection(layout, row++, "SWP SYSTEM");

            // Reaching a system by its SAVED SESSION rather than by its SWP file: the sidecar
            // carries every field this form asks for, so one pick fills the whole form — and it
            // works for a system whose SWP was renamed or moved to a new dated version folder,
            // which picking the workbook cannot express.
            Button[] sessionButtons = AddSessionRow(layout, row++, "Session", _session, pathRow,
                ("Open…", BrowseSession), ("Save as…", SaveSessionAs), ("Save as new…", SaveSessionAsNew));
            _tips.SetToolTip(sessionButtons[1],
                "Save this session under a new name, anywhere. Everything goes with it — settings, per-row " +
                "decisions, dictionaries and the queue's whole history — and both this dialog and a workspace " +
                "running on the session switch to the new file.");
            _tips.SetToolTip(sessionButtons[2],
                "Start a NEW session from this one: every setting and file reference, the dictionaries, the " +
                "glossary corrections and the strobe templates are kept; the queue's history and the per-row " +
                "decisions start empty. This dialog switches to the new file; a running workspace stays on its own.");
            _tips.SetToolTip(_session,
                "Open a saved " + SwpBridgeSidecarStore.FileName + " — the sidecar at an SWP system's root, " +
                "holding its settings, per-row decisions, dictionaries and queue. Picking one fills this whole " +
                "form from it; its folder becomes the system root, so decisions keep going back to the same " +
                "file. A \"" + SwpBridgeSidecarStore.FileName + ".bak\" opens too — that is how a damaged " +
                "sidecar is recovered: open the .bak, and the next decision writes it back as the live one.");

            AddPathRow(layout, row++, "SWP file", _swp, BrowseSwp, pathRow);

            _derived.Dock = DockStyle.Fill;
            layout.Controls.Add(_derived, 1, row);
            layout.SetColumnSpan(_derived, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 2 * line + 6));   // two wrapped lines of path
            row++;

            AddOptionalWorkbookRow(layout, row++, "Reference SWP (base)", _referenceSwp,
                "Reference SWP|*.xlsx;*.xlsm|All files|*.*",
                "REQUIRED. A known-good SWP of the same plant type. The formula-structure audit compares " +
                "every block's H/I/Q/R formulas against it and queues repairs for the ones that drifted — " +
                "and it is the only way back for a row whose CONFIG_PLANT template reference was destroyed, " +
                "so a workspace without one has no answer when a formula turns out damaged.", textRow);

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

            layout.Controls.Add(_sheetsLabel, 0, row);
            _sheets.Dock = DockStyle.Fill;
            layout.Controls.Add(_sheets, 1, row);
            layout.SetColumnSpan(_sheets, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            row++;

            // The two regexes decide which of this list's mnemonics are diagnostics at all, so
            // they belong to the IO list, not to the rules about what gets written.
            AddTextRow(layout, row++, "Include regex", _includeRegex, "extra mnemonics accepted as diagnostics", textRow);
            AddTextRow(layout, row++, "Exclude regex", _excludeRegex, "accepted mnemonics dropped anyway", textRow);

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
                "for upstream. Flat copies directly in the folder work too. SAFETY consumes only the " +
                "SCADA's Alarms_*.txt export (an \"Alarms List\" child beside the SWP works), so the " +
                "SWP's own version folder can be the root.");
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

            var languagePanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(3, 2, 3, 2),
            };
            languagePanel.Controls.Add(_glossaryAlways);
            layout.Controls.Add(languagePanel, 1, row);
            layout.SetColumnSpan(languagePanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row++;

            // ---- CONFIG_PLANT --------------------------------------------------------------
            // Two answers to ONE question — what happens to the template's text — so they are
            // radio buttons, with the same three-level resolution the FLD treatment uses:
            // per-row tick → destination-sheet rule → this default.
            AddSection(layout, row++, "CONFIG_PLANT");
            var configPanel = SectionPanel();
            configPanel.Controls.Add(new Label
            {
                Text = "SCADA shows CONFIG_PLANT's template followed by the SWP's description cell " +
                       "(per-row: the [Keep CONFIG text] column decides outright):",
                AutoSize = true,
                Padding = new Padding(0, 0, 0, 2),
            });
            var configModePanel = IndentedPanel();
            configModePanel.Controls.Add(_keepConfigText);
            configModePanel.Controls.Add(_replaceTail);
            configPanel.Controls.Add(configModePanel);
            var configSheetPanel = IndentedPanel();
            configSheetPanel.Controls.Add(_configSheetRuleCaption);
            configSheetPanel.Controls.Add(_configSheetRegex);
            configPanel.Controls.Add(configSheetPanel);
            _keepConfigText.CheckedChanged += (_, __) => UpdateConfigSheetRuleCaption();
            _replaceTail.CheckedChanged += (_, __) => UpdateConfigSheetRuleCaption();
            UpdateConfigSheetRuleCaption();
            layout.Controls.Add(configPanel, 1, row);
            layout.SetColumnSpan(configPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row++;

            // ---- ELECTRICAL_DESIGNATION ----------------------------------------------------
            AddSection(layout, row++, "ELECTRICAL_DESIGNATION");
            var optionsPanel = SectionPanel();
            // The mismatch treatment is one exclusive choice; the warning fires regardless and
            // is suppressed per row with the [Mismatch OK] override.
            optionsPanel.Controls.Add(new Label
            {
                Text = "On functional-unit / location mismatch (always warns; per-row [Mismatch OK] silences it):",
                AutoSize = true,
                Padding = new Padding(0, 0, 0, 2),
            });
            var mismatchPanel = IndentedPanel();
            mismatchPanel.Controls.Add(_mismatchAppend);
            mismatchPanel.Controls.Add(_mismatchOverwrite);
            optionsPanel.Controls.Add(mismatchPanel);
            // The destination-sheet exception to that choice. ONE regex is enough because the
            // treatment is binary — it names the SWP sheets that take the OTHER mode — and the
            // caption says which mode that is, the same opposite-of-the-default wording the
            // per-row [Append FLD] / [Overwrite FLD] column uses. A per-row tick still wins.
            var fldSheetPanel = IndentedPanel();
            fldSheetPanel.Controls.Add(_fldSheetRuleCaption);
            fldSheetPanel.Controls.Add(_fldSheetRegex);
            optionsPanel.Controls.Add(fldSheetPanel);
            _mismatchAppend.CheckedChanged += (_, __) => UpdateFldSheetRuleCaption();
            _mismatchOverwrite.CheckedChanged += (_, __) => UpdateFldSheetRuleCaption();
            UpdateFldSheetRuleCaption();
            layout.Controls.Add(optionsPanel, 1, row);
            layout.SetColumnSpan(optionsPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row++;

            _tips.SetToolTip(_keepConfigText,
                "The template already names the bit, so writing the IO text into the description cell stacks a " +
                "second sentence onto it. Chosen: every row keeps CONFIG's wording and the bridge contributes " +
                "only the FLD label — nothing at all when there is no mismatch, so a template-fed cell is never " +
                "blanked.");
            _tips.SetToolTip(_replaceTail,
                "A template reads \"INDUCTION %1 - BELT %2 - SPARE ALARM.63\": the head names the place, the " +
                "tail names the bit generically — and the SWP's own T/U cell already names that bit in the " +
                "plant's words. Chosen, every joined row's Q/R formula is rebuilt to keep the template up to " +
                "its last placeholder and put T/U where the tail was, so SCADA reads \"INDUCTION =S1+IN22 - " +
                "BELT =S1+IN22.05 - FIELD BUS FAILURE - DRIVE […]\". Rows this bridge generated (photocells, " +
                "strobes) and \"P*\"-typed module rows keep CONFIG's text anyway — it is the text this bridge " +
                "wrote — as do rows whose T/U would carry the FLD label alone, or the alarm's own name would " +
                "be lost.");
            _tips.SetToolTip(_configSheetRegex,
                "A .NET regex matched (case-insensitively) against the DESTINATION SWP sheet name, e.g. " +
                "\"BELT$\" or \"^ILGRP\". Rows landing on a sheet it matches take the treatment named to the " +
                "left instead of the default above; every other sheet follows the default. Empty = no " +
                "exception. A per-row [Keep CONFIG text] / [Write IO text] tick still wins over it, and the " +
                "build log reports how many rows the pattern moved.");
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
            _tips.SetToolTip(_fldSheetRegex,
                "A .NET regex matched (case-insensitively) against the DESTINATION SWP sheet name, e.g. " +
                "\"BELT$\" or \"^UP0[12]_\". Rows landing on a sheet it matches take the treatment named to the " +
                "left instead of the default above; every other sheet follows the default. Empty = no exception. " +
                "A per-row [Append FLD] / [Overwrite FLD] tick still wins over this, and the FLD rule holds " +
                "either way — only the place the identity lands changes. The build log reports how many rows " +
                "the pattern moved.");

            var applyPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill, WrapContents = false };
            applyPanel.Controls.Add(_inPlace);
            applyPanel.Controls.Add(_versionFolder);
            applyPanel.Controls.Add(_copyCaption);
            applyPanel.Controls.Add(_copyRegex);
            _versionFolder.CheckedChanged += (s, e) => UpdateCopyRegexEnabled();
            _copyRegex.TextChanged += (s, e) => Revalidate();
            UpdateCopyRegexEnabled();
            layout.Controls.Add(new Label { Text = "Apply writes", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            layout.Controls.Add(applyPanel, 1, row);
            layout.SetColumnSpan(applyPanel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, pathRow));
            row++;
            _tips.SetToolTip(_inPlace, "Patches the SWP where it is; the previous file stays beside it as .bak.");
            _tips.SetToolTip(_versionFolder,
                "Creates the next NN_dd-MMM-yyyy sibling folder, writes the patched SWP there and copies " +
                "CONFIG_PLANT alongside; the current folder is never written.");
            _tips.SetToolTip(_copyRegex,
                "A .NET regex matched (case-insensitively) against the FILE NAMES in the current version folder, " +
                "e.g. \"^SWP_02_Induction_generated_G1\\.xlsx$\" for the other SWP this IO list drives, or " +
                "\"\\.xlsx$\". When Apply creates the next dated folder, every match is copied into it too, so the " +
                "new folder is a complete version. Top-level files only; Excel's ~$ lock files and .bak copies " +
                "never. Empty = only the SWP and CONFIG_PLANT.");

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
                radio.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) DeriveNumbersFromSwpSheets(); };
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

            // Coming back to a session: the one RUNNING in this Excel first (reconfiguring what is
            // open), then the last one this dialog built or saved, and only then the last SWP root —
            // a named session is a file, and returning to it must not fall back to a root's default.
            string reopen = null;
            foreach (string candidate in new[] { SwpBridgeSession.Current?.EffectiveSidecarPath, SwpBridgeRegistry.GetLastSession() })
                if (!string.IsNullOrEmpty(candidate) && File.Exists(candidate)) { reopen = candidate; break; }
            if (reopen != null)
            {
                LoadSession(reopen);
            }
            else
            {
                string lastRoot = SwpBridgeRegistry.GetLastRoot();
                if (!string.IsNullOrEmpty(lastRoot) && Directory.Exists(lastRoot))
                    TryPrefillFromRoot(lastRoot);
            }
            UpdatePlantUi();
            Revalidate();
        }

        // ---- long paths show their tail -----------------------------------------------------

        private TextBox[] PathBoxes() => new[] { _session, _swp, _referenceSwp, _ioList, _productionRoot, _glossary1, _glossary2 };

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

        /// <summary>A section's stacked body — one question per line, top to bottom.</summary>
        private static FlowLayoutPanel SectionPanel() => new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Margin = new Padding(3, 2, 3, 6),
        };

        /// <summary>The answers to a question, indented under it.</summary>
        private static FlowLayoutPanel IndentedPanel() => new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = true,
            Margin = new Padding(16, 0, 0, 0),
        };

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

        private void AddPathRow(TableLayoutPanel layout, int row, string label, TextBox box, Action browse,
            int height, string buttonText = "Browse…")
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);
            box.Dock = DockStyle.Fill;
            layout.Controls.Add(box, 1, row);
            var button = new Button { Text = buttonText, Dock = DockStyle.Fill };
            button.Click += (s, e) => browse();
            layout.Controls.Add(button, 2, row);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
        }

        /// <summary>The Session row: its path and the ways to reach one — open an existing session,
        /// save it under a new name, start a new one from it. The buttons sit beside the path in a
        /// panel spanning the button column too, each as wide as its caption, so they never widen
        /// the single "Browse…" of every other row.</summary>
        private Button[] AddSessionRow(TableLayoutPanel layout, int row, string label, TextBox box, int height,
            params (string Text, Action Click)[] actions)
        {
            layout.Controls.Add(new Label { Text = label, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, row);

            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1 + actions.Length,
                RowCount = 1,
                Margin = new Padding(0),
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            box.Dock = DockStyle.Fill;
            panel.Controls.Add(box, 0, 0);

            var buttons = new Button[actions.Length];
            for (int i = 0; i < actions.Length; i++)
            {
                int width = Math.Max(96, TextRenderer.MeasureText(actions[i].Text, Font).Width + 28);
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, width));
                Action click = actions[i].Click;
                var button = new Button { Text = actions[i].Text, Dock = DockStyle.Fill, Margin = new Padding(i == 0 ? 3 : 1, 3, 3, 3) };
                button.Click += (s, e) => click();
                panel.Controls.Add(button, 1 + i, 0);
                buttons[i] = button;
            }
            layout.Controls.Add(panel, 1, row);
            layout.SetColumnSpan(panel, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, height));
            return buttons;
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
            _plantNote.Text = _derivedNumbers.Length > 0
                ? _derivedNumbers
                : "ISPC: first induction line id · line group — UPS: first system id · conveyor group";
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

        private static bool SameFile(string a, string b)
        {
            try
            {
                return a.Trim().Length > 0 && b.Trim().Length > 0 &&
                    string.Equals(Path.GetFullPath(a.Trim()), Path.GetFullPath(b.Trim()),
                        StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
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

        /// <summary>
        /// Opens a saved session: the picked sidecar fills the form, and the folder it lives in
        /// becomes the SWP root. Unlike the automatic prefill (which only ever completes what the
        /// SWP pick left blank), this REPLACES the form — asking for a session is asking for that
        /// session, including its SWP file.
        /// </summary>
        private void BrowseSession()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = SessionOpenFilter,
                CheckFileExists = true,
                Title = "Open an SWP Bridge session",
            })
            {
                string start = SessionsFolder();
                if (Directory.Exists(start)) dialog.InitialDirectory = start;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                LoadSession(dialog.FileName);
            }
        }

        private void LoadSession(string path)
        {
            SwpSidecarDocument document;
            try
            {
                document = SwpBridgeSidecarStore.LoadFile(path);
            }
            catch (Exception ex)
            {
                _hint.Text = ex.Message;
                return;
            }
            if (document == null)
            {
                _hint.Text = "That file no longer exists.";
                return;
            }

            _loaded = document;
            _session.Text = path;
            ScrollToEnd(_session);
            _swp.Text = "";                                    // the session's own SWP takes over
            SwpRootFolder = SwpBridgeSidecarStore.RootOf(path, document);
            _loadedSwpFileName = document.Settings.SwpFileName ?? "";
            _loadedSessionPath = path;
            _derived.Text = "root " + SwpRootFolder;
            ApplyLoadedSidecar(SwpRootFolder);
            DeriveNumbersFromSwpSheets();
            if (document.LoadNotes.Count > 0)
                _hint.Text = string.Join("   · ", document.LoadNotes);
            if (_swp.Text.Length == 0)
                _hint.Text = "The session's SWP (" + document.Settings.SwpFileName + " in \\" +
                    document.Settings.CurrentVersionFolder + ") is not there — pick the workbook " +
                    "to point this session at its current version folder.";
            Revalidate();
        }

        /// <summary>
        /// Saves this session under a new name, anywhere — and from then on the session IS that file:
        /// the dialog works on it, a workspace running on the session this dialog had open writes its
        /// next decision there, and Setup comes back to it. That is what lets one IO list drive two
        /// SWP systems as two named sessions kept together in one folder.
        ///
        /// The file is written NOW, with everything the session holds — its settings as the dialog
        /// shows them plus its per-row decisions, dictionaries and queue — so reopening it later
        /// finds all of it in place.
        /// </summary>
        private void SaveSessionAs()
        {
            if (!File.Exists(_swp.Text))
            {
                _hint.Text = "Pick the SWP workbook first — a session records which system it belongs to.";
                return;
            }
            using (var dialog = new SaveFileDialog
            {
                Filter = SessionSaveFilter,
                DefaultExt = "json",
                Title = "Save this SWP Bridge session as",
                FileName = SuggestedSessionFileName(),
                OverwritePrompt = true,
            })
            {
                string start = SessionsFolder();
                if (Directory.Exists(start)) dialog.InitialDirectory = start;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                SaveSessionTo(dialog.FileName);
            }
        }

        /// <summary>The Save as behind its file dialog: write the session, then switch everything to it.</summary>
        internal bool SaveSessionTo(string path)
        {
            SwpSidecarDocument document = CollectDocument(path);
            try
            {
                SwpBridgeSidecarStore.SaveFile(path, document);
            }
            catch (Exception ex)
            {
                _hint.Text = "The session could not be saved: " + ex.Message;
                return false;
            }

            // A workspace running on the session this dialog had open moves with it — otherwise its
            // next decision would land in the old file while Setup comes back to the new one.
            SwpBridgeSession live = SwpBridgeSession.Current;
            if (live != null && _loadedSessionPath.Length > 0 && SameFile(live.EffectiveSidecarPath, _loadedSessionPath))
            {
                live.SidecarPath = path;
                live.Sidecar.Settings.SwpRootFolder = document.Settings.SwpRootFolder;
                _retargetedLive = live;
            }

            _loaded = document;
            _loadedSessionPath = path;
            _loadedSwpFileName = document.Settings.SwpFileName ?? "";
            SessionPath = path;
            _session.Text = path;
            ScrollToEnd(_session);
            SwpBridgeRegistry.SetLastSession(path);
            Revalidate();
            if (_hint.Text.Length == 0) _hint.Text = "Saved — the session is now " + path;
            return true;
        }

        /// <summary>
        /// Starts a NEW session from this one, saved under a new name anywhere: every setting and file
        /// reference the dialog holds, and the knowledge that is not about any row — the dictionaries,
        /// glossary corrections and strobe templates — but none of the history: the queue and the
        /// per-row decisions start empty (<see cref="SwpSidecarJson.FreshSession"/>). A Save as copies
        /// that history on purpose; carried into a second SWP system driven by the same IO list, it is
        /// the first system's decisions riding along on mnemonics both cards share.
        ///
        /// The dialog switches to the new file like a Save as, but a running workspace does NOT move:
        /// its queue and decisions belong to the file they came from.
        /// </summary>
        private void SaveSessionAsNew()
        {
            if (!File.Exists(_swp.Text))
            {
                _hint.Text = "Pick the SWP workbook first — a session records which system it belongs to.";
                return;
            }
            using (var dialog = new SaveFileDialog
            {
                Filter = SessionSaveFilter,
                DefaultExt = "json",
                Title = "Save as a new SWP Bridge session — settings kept, history cleared",
                FileName = SuggestedSessionFileName(),
                OverwritePrompt = true,
            })
            {
                string start = SessionsFolder();
                if (Directory.Exists(start)) dialog.InitialDirectory = start;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                SaveNewSessionTo(dialog.FileName);
            }
        }

        /// <summary>The Save as new behind its file dialog: write the fresh session, then switch the
        /// dialog to it — and only the dialog.</summary>
        internal bool SaveNewSessionTo(string path)
        {
            SwpSidecarDocument fresh;
            try
            {
                fresh = SwpSidecarJson.FreshSession(CollectDocument(path));
                SwpBridgeSidecarStore.SaveFile(path, fresh);
            }
            catch (Exception ex)
            {
                _hint.Text = "The new session could not be saved: " + ex.Message;
                return false;
            }

            _loaded = fresh;
            _loadedSessionPath = path;
            _loadedSwpFileName = fresh.Settings.SwpFileName ?? "";
            SessionPath = path;
            _session.Text = path;
            ScrollToEnd(_session);
            SwpBridgeRegistry.SetLastSession(path);
            Revalidate();
            if (_hint.Text.Length == 0)
                _hint.Text = "Saved as a new session — settings kept, history cleared: " + path;
            return true;
        }

        /// <summary>
        /// Where both session dialogs start: the folder of the session in hand when it is a named one,
        /// else the last session's — sessions kept together stay together — and only then the root.
        /// </summary>
        private string SessionsFolder()
        {
            string current = _session.Text.Trim();
            foreach (string candidate in new[] { IsNamedSession(current) ? current : "", SwpBridgeRegistry.GetLastSession() ?? "", current })
            {
                if (candidate.Length == 0) continue;
                string folder = Path.GetDirectoryName(candidate);
                if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder)) return folder;
            }
            if (SwpRootFolder.Length > 0 && Directory.Exists(SwpRootFolder)) return SwpRootFolder;
            return File.Exists(_swp.Text) ? Path.GetDirectoryName(_swp.Text) : "";
        }

        /// <summary>The SWP workbook's own name is what tells two sessions of one IO list apart.</summary>
        private string SuggestedSessionFileName()
        {
            string current = _session.Text.Trim();
            if (current.Length > 0 &&
                !string.Equals(Path.GetFileName(current), SwpBridgeSidecarStore.FileName, StringComparison.OrdinalIgnoreCase))
                return Path.GetFileName(current);
            string swp = Path.GetFileNameWithoutExtension(_swp.Text);
            return swp.Length > 0 ? swp + ".swp-bridge.json" : SwpBridgeSidecarStore.FileName;
        }

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

                // A session opened or saved BY NAME stays the session: picking its SWP — or the other
                // SWP an IO list drives — must not swap it for the root's default sidecar. (Building it
                // for a different SWP still warns, and Save as… is how both are kept.)
                if (!IsNamedSession(_session.Text)) TryPrefillFromRoot(SwpRootFolder);
                DeriveNumbersFromSwpSheets();
            }
            catch (Exception ex)
            {
                _derived.Text = "could not derive the system root: " + ex.Message;
            }
        }

        /// <summary>
        /// Reads the first-line and group numbers off the SWP's OWN tabs (IL02…, ILGRP02_…) —
        /// the map builds its destination sheet names from exactly those numbers, so asking the
        /// user to retype them is asking them to re-derive what the file already states, and one
        /// wrong digit sends every row to a sheet that does not exist. The values stay editable;
        /// the note under them says where they came from.
        /// </summary>
        private void DeriveNumbersFromSwpSheets()
        {
            _derivedNumbers = "";
            if (!File.Exists(_swp.Text)) { UpdatePlantUi(); return; }

            SwpPlantNumbering.Result numbering;
            try
            {
                numbering = SwpPlantNumbering.FromSheetNames(ListSheetNames(_swp.Text), SelectedPlantType());
            }
            catch (Exception ex)
            {
                _derivedNumbers = "the SWP's sheet names could not be read (" + ex.Message + ") — the numbers stand as typed";
                UpdatePlantUi();
                return;
            }

            if (numbering.HasFirstLine && numbering.FirstLineId >= _firstLine.Minimum && numbering.FirstLineId <= _firstLine.Maximum)
                _firstLine.Value = numbering.FirstLineId;
            if (numbering.HasGroup && numbering.GroupIndex >= _group.Minimum && numbering.GroupIndex <= _group.Maximum)
                _group.Value = numbering.GroupIndex;

            if (numbering.Any)
                _derivedNumbers = "read from the SWP's sheets " + numbering.Evidence +
                    (numbering.HasGroup ? "" : " (no group in the tab names — it stands as typed)");
            else if (SelectedPlantType() == SwpPlantType.Ispc || SelectedPlantType() == SwpPlantType.Ups)
                _derivedNumbers = "no numbered sheet found in the SWP — the numbers stand as typed";
            UpdatePlantUi();
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
                _session.Text = "";
                return;
            }
            // A root with no sidecar is a NEW system: the session box must stop naming the last
            // one, or OK would look like it were about to reopen it.
            if (_loaded == null) { _session.Text = ""; _loadedSwpFileName = ""; _loadedSessionPath = ""; return; }
            _session.Text = SwpBridgeSidecarStore.PathFor(root);
            // The root's sidecar belongs to whichever SWP it was saved for: picking ANOTHER SWP in the
            // same root and building would overwrite it, so the hint must know whose it is.
            _loadedSessionPath = _session.Text;
            _loadedSwpFileName = _loaded.Settings.SwpFileName ?? "";
            ScrollToEnd(_session);
            ApplyLoadedSidecar(root);
        }

        /// <summary>Pours <see cref="_loaded"/> into the form. The SWP path is only FILLED, never
        /// replaced: a workbook the user just picked outranks the one the sidecar remembers.</summary>
        private void ApplyLoadedSidecar(string root)
        {
            SwpBridgeSettings s = _loaded.Settings;
            if (_swp.Text.Length == 0 && s.SwpFileName.Length > 0)
            {
                // Coming back after ANOTHER session's apply made a newer dated folder with this SWP
                // copied along: the session follows the newest version folder holding its workbook —
                // patching the old one would fork the chronology. The derived line says so.
                string recorded = s.CurrentVersionFolder ?? "";
                string newest = NewestVersionFolderHolding(root, s.SwpFileName);
                string folder = newest ?? recorded;
                string candidate = Path.Combine(root, folder, s.SwpFileName);
                if (File.Exists(candidate))
                {
                    _swp.Text = candidate;
                    _derived.Text = "root " + root + "   ·   version " + folder +
                        (newest != null && !string.Equals(newest, recorded, StringComparison.OrdinalIgnoreCase)
                            ? "   (the session said " + (recorded.Length > 0 ? recorded : "no version folder") + "; " +
                              newest + " is the newest version folder holding " + s.SwpFileName + ")"
                            : "   (from the sidecar)");
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
            _keepConfigText.Checked = s.ConfigMode == SwpConfigMode.KeepConfigText;
            _replaceTail.Checked = s.ConfigMode == SwpConfigMode.ReplaceTemplateTail;
            _configSheetRegex.Text = s.ConfigSheetRulePattern;
            UpdateConfigSheetRuleCaption();
            _mismatchAppend.Checked = s.MismatchMode == SwpMismatchMode.AppendToDescription;
            _mismatchOverwrite.Checked = s.MismatchMode == SwpMismatchMode.Overwrite;
            _fldSheetRegex.Text = s.FldSheetRulePattern;
            UpdateFldSheetRuleCaption();
            _inPlace.Checked = !s.ApplyToVersionFolder;
            _versionFolder.Checked = s.ApplyToVersionFolder;
            _copyRegex.Text = s.VersionFolderCopyPattern ?? "";
            UpdateCopyRegexEnabled();
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
        /// deliberately unticked stay off — a card added to the IO list since is in by default.
        /// Card sheets are the R{rack}.S{slot}.C{card} names when the workbook has any; otherwise
        /// the sheets whose header row carries Mnemonic + Bit (the safety PLC's list names its
        /// one card sheet by network), and the label says so.</summary>
        private void LoadSheetList(List<string> excludedNames)
        {
            _sheets.Items.Clear();
            _sheetsLabel.Text = "Card sheets";
            try
            {
                var excluded = new HashSet<string>(excludedNames ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                IoListCardSheets.Result cards = IoListCardSheets.Find(ReadBytesShared(_ioList.Text));
                foreach (string name in cards.Names)
                    _sheets.Items.Add(name, !excluded.Contains(name));

                if (_sheets.Items.Count == 0)
                    _hint.Text = "No R{rack}.S{slot}.C{card} sheet and no sheet headed Mnemonic + Bit in this workbook — nothing can join.";
                else if (cards.ByContent)
                    _sheetsLabel.Text = "Card sheets (by header: no R.S.C names)";
            }
            catch (Exception ex)
            {
                _hint.Text = "Could not list the workbook's sheets: " + ex.Message;
            }
        }

        /// <summary>The file as bytes, read shared — works while Excel holds the file.</summary>
        private static byte[] ReadBytesShared(string path)
        {
            using (var input = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var buffer = new MemoryStream())
            {
                input.CopyTo(buffer);
                return buffer.ToArray();
            }
        }

        /// <summary>Sheet names via OpenXml over an in-memory copy — works while Excel holds the file.</summary>
        private static List<string> ListSheetNames(string path)
        {
            byte[] bytes = ReadBytesShared(path);

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

            // The reference SWP is REQUIRED: it is the only thing that can tell a row whose
            // formulas drifted from one that is merely different, and the only way back for a
            // row whose CONFIG_PLANT template reference was destroyed.
            string reference = _referenceSwp.Text.Trim();
            if (reference.Length == 0) problems.Add("pick the reference SWP");
            else if (SameFile(reference, _swp.Text)) problems.Add("the reference SWP is the SWP itself");

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
            if (!IsValidRegex(_fldSheetRegex.Text))
                problems.Add("the FLD destination-sheet regex does not parse");
            if (!IsValidRegex(_configSheetRegex.Text))
                problems.Add("the CONFIG_PLANT destination-sheet regex does not parse");
            if (!IsValidRegex(_copyRegex.Text))
                problems.Add("the version-folder copy regex does not parse");

            // Pointing an OPENED session at a different SWP and building would overwrite that
            // session with the other system's decisions — which is exactly the "one IO list, two
            // SWPs" case, and exactly where Save as… is the answer. Said before OK is pressed.
            if (problems.Count == 0 && _loadedSwpFileName.Length > 0 && _swp.Text.Length > 0 &&
                !string.Equals(Path.GetFileName(_swp.Text), _loadedSwpFileName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_session.Text.Trim(), _loadedSessionPath, StringComparison.OrdinalIgnoreCase))
                _hint.Text = "This session was saved for " + _loadedSwpFileName +
                    " — building will overwrite its decisions with " + Path.GetFileName(_swp.Text) +
                    "'s. Use \"Save as…\" to keep both.";
            else
                _hint.Text = problems.Count == 0 ? "" : "· " + string.Join("   · ", problems);
            _ok.Enabled = problems.Count == 0;
        }

        private void OnClosingCollect(object sender, FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK)
            {
                // A Save as moved the running workspace to its new file; cancelling keeps that
                // workspace's settings as they run, so the new file must say exactly that.
                _retargetedLive?.SaveSidecarNow();
                return;
            }

            SessionPath = _session.Text.Trim().Length > 0
                ? _session.Text.Trim()
                : SwpBridgeSidecarStore.PathFor(SwpRootFolder);
            Result = CollectDocument(SessionPath);
            SwpBridgeRegistry.SetLastRoot(SwpRootFolder);
            SwpBridgeRegistry.SetLastSession(SessionPath);
        }

        /// <summary>
        /// The session as the dialog holds it, for the file <paramref name="sessionPath"/>: the loaded
        /// document — its decisions, dictionaries and queue — carrying the dialog's settings.
        /// </summary>
        private SwpSidecarDocument CollectDocument(string sessionPath)
        {
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
            s.ConfigMode = _keepConfigText.Checked
                ? SwpConfigMode.KeepConfigText
                : SwpConfigMode.ReplaceTemplateTail;
            s.ConfigSheetRulePattern = _configSheetRegex.Text.Trim();
            s.ConfigSheetRuleMode = ConfigSheetRuleMode;
            s.MismatchMode = _mismatchOverwrite.Checked
                ? SwpMismatchMode.Overwrite
                : SwpMismatchMode.AppendToDescription;
            s.FldSheetRulePattern = _fldSheetRegex.Text.Trim();
            // Stored as the mode it chose, not as "the opposite" — a later change of the default
            // can then never silently reverse what this rule does to the sheets it matches.
            s.FldSheetRuleMode = FldSheetRuleMode;
            s.ApplyToVersionFolder = _versionFolder.Checked;
            s.VersionFolderCopyPattern = _copyRegex.Text.Trim();
            s.SecondLanguage = _secondLangCode.Text.Trim().ToUpperInvariant();

            // A session saved at its own root stays self-locating (empty root = "my own folder"),
            // so the ordinary sidecar survives having its whole system folder moved. Only a
            // session saved ELSEWHERE has to record which system it belongs to.
            s.SwpRootFolder = SameFolder(Path.GetDirectoryName(sessionPath) ?? "", SwpRootFolder)
                ? ""
                : SwpRootFolder;
            return document;
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

        /// <summary>
        /// The treatment the destination-sheet rule imposes: always the OPPOSITE of the workspace
        /// default, because the mismatch treatment is binary — a rule naming the default would be
        /// no exception at all.
        /// </summary>
        private SwpMismatchMode FldSheetRuleMode =>
            _mismatchOverwrite.Checked ? SwpMismatchMode.AppendToDescription : SwpMismatchMode.Overwrite;

        /// <summary>The CONFIG_PLANT rule's mode — always the OPPOSITE of the default, for the
        /// same reason: a rule naming the default would be no exception at all.</summary>
        private SwpConfigMode ConfigSheetRuleMode =>
            _keepConfigText.Checked ? SwpConfigMode.ReplaceTemplateTail : SwpConfigMode.KeepConfigText;

        private void UpdateConfigSheetRuleCaption() =>
            _configSheetRuleCaption.Text = _keepConfigText.Checked
                ? "…but replace the trailing text on SWP sheets matching:"
                : "…but never overwrite on SWP sheets matching:";

        /// <summary>
        /// Keeps the rule's caption naming what it DOES, so flipping the default above can never
        /// leave the sentence describing the opposite of what will be written.
        /// </summary>
        private void UpdateFldSheetRuleCaption() =>
            _fldSheetRuleCaption.Text = _mismatchOverwrite.Checked
                ? "…but append identity to descriptions on SWP sheets matching:"
                : "…but overwrite the row's FU/Loc/Device cells on SWP sheets matching:";

        private void UpdateCopyRegexEnabled()
        {
            _copyRegex.Enabled = _versionFolder.Checked;
            _copyCaption.Enabled = _versionFolder.Checked;
        }

        /// <summary>A session file under a name of its own — not a root's default swp-bridge.json.</summary>
        private static bool IsNamedSession(string path) =>
            !string.IsNullOrWhiteSpace(path) &&
            !string.Equals(Path.GetFileName(path.Trim()), SwpBridgeSidecarStore.FileName, StringComparison.OrdinalIgnoreCase);

        /// <summary>The newest dated (NN_…) version folder under the root that holds this SWP, or null.</summary>
        private static string NewestVersionFolderHolding(string root, string swpFileName)
        {
            try
            {
                var holding = new List<string>();
                foreach (string dir in Directory.GetDirectories(root))
                    if (File.Exists(Path.Combine(dir, swpFileName))) holding.Add(Path.GetFileName(dir));
                return SwpVersionFolders.Newest(holding);
            }
            catch { return null; }
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
