using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelPowerTools.Commands;
using ExcelPowerTools.Core.Comparison.Persistence;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Searchable home for the complete add-in. This is intentionally a launcher/dashboard for the first
    /// iteration: transactional dialogs stay modal, while the shell can host richer pages as modules grow.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("1EE55874-BEC9-4A8B-9269-5B523E7C4C63")]
    public sealed class PowerToolsPane : UserControl
    {
        private const string SearchPlaceholder = "Search every Power Tool...";

        private static readonly Color Background = Color.FromArgb(246, 248, 252);
        private static readonly Color Ink = Color.FromArgb(28, 35, 56);
        private static readonly Color Muted = Color.FromArgb(101, 112, 136);
        private static readonly Color Accent = Color.FromArgb(75, 78, 225);
        private static readonly Color Border = Color.FromArgb(221, 226, 236);
        private static readonly string[] SectionOrder =
            { "Comparison", "Files", "Transform", "Data", "History", "Workbook", "IoPipeline", "Settings", "Developer" };

        private readonly TextBox _search = new TextBox();
        private readonly Button _clearSearch = new Button();
        private readonly TableLayoutPanel _root = new TableLayoutPanel();
        private readonly FlowLayoutPanel _categories = new FlowLayoutPanel();
        private readonly FlowLayoutPanel _catalog = new FlowLayoutPanel();
        private readonly ComparisonWorkspacePage _comparisonWorkspace = new ComparisonWorkspacePage();
        private readonly Label _context = new SmoothLabel();
        private readonly Label _resultCount = new SmoothLabel();
        private readonly Dictionary<string, CategoryChip> _chips =
            new Dictionary<string, CategoryChip>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Label> _sectionHeaders =
            new Dictionary<string, Label>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<ToolDefinition, ToolCard> _cards =
            new Dictionary<ToolDefinition, ToolCard>();
        private readonly List<ToolDefinition> _tools;

        private Panel _emptyState;
        private Label _emptyTitle;
        private System.Drawing.Image _logoImage;
        private string _selectedCategory = "All";
        private int _ownerWindowHandle;
        private bool _detached;
        private bool _refreshingCatalog;

        public PowerToolsPane()
        {
            Name = "PowerToolsPane";
            Dock = DockStyle.Fill;
            MinimumSize = new Size(260, 320);
            BackColor = Background;
            ForeColor = Ink;
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            DoubleBuffered = true;

            _tools = BuildToolCatalog();
            BuildLayout();
            InitializeCatalog();
            SetSearchPlaceholder();
            SelectCategory("All");
            Resize += (_, __) => UpdateResponsiveLayout();
            UpdateResponsiveLayout();

            SelectionEvents.Changed += OnExcelContextChanged;
            Disposed += (_, __) => Detach();
            RefreshContext();
        }

        internal void RefreshState()
        {
            if (_detached || IsDisposed) return;
            if (InvokeRequired)
            {
                try { BeginInvoke((Action)RefreshState); } catch { }
                return;
            }

            RefreshContext();
            foreach (ToolCard card in _cards.Values) card.RefreshSetting();
        }

        internal void AttachWindow(int windowHandle)
        {
            _ownerWindowHandle = windowHandle;
            RefreshContext();
        }

        internal void Detach()
        {
            if (_detached) return;
            _detached = true;
            SelectionEvents.Changed -= OnExcelContextChanged;
            _comparisonWorkspace.ClearLoadedReview();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) Detach();
            base.Dispose(disposing);
            if (!disposing) return;
            _logoImage?.Dispose();
            _logoImage = null;
        }

        private void BuildLayout()
        {
            SuspendLayout();

            _root.Dock = DockStyle.Fill;
            _root.BackColor = Background;
            _root.ColumnCount = 1;
            _root.RowCount = 5;
            _root.Margin = Padding.Empty;
            _root.Padding = Padding.Empty;
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 108));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

            _root.Controls.Add(BuildHeader(), 0, 0);
            _root.Controls.Add(BuildSearch(), 0, 1);
            _root.Controls.Add(BuildCategoryBar(), 0, 2);
            _root.Controls.Add(BuildCatalog(), 0, 3);
            _root.Controls.Add(BuildFooter(), 0, 4);

            Controls.Add(_root);

            _comparisonWorkspace.Visible = false;
            _comparisonWorkspace.BackRequested += (_, __) => ShowToolsCatalog();
            _comparisonWorkspace.NewComparisonRequested += (_, __) => QueuePaneCommand(
                "New comparison",
                ComparisonProjectCommands.StartComparisonProject);
            _comparisonWorkspace.ChooseReviewRequested += (_, __) => QueuePaneCommand(
                "Validate Review workbook",
                SelectReviewForWorkspace);
            _comparisonWorkspace.OpenValidatedReviewRequested += (_, e) =>
            {
                ReviewWorkbookHydrationSession selected = e.Session;
                QueuePaneCommand(
                    "Open protected Review copy",
                    () => OpenHydratedReviewForWorkspace(selected));
            };
            _comparisonWorkspace.DecisionCommitRequested += (_, e) =>
                QueuePaneCommand(
                    "Save review decision revision",
                    () => CommitDecisionRevisionForWorkspace(e));
            Controls.Add(_comparisonWorkspace);
            ResumeLayout(true);
        }

        private Control BuildHeader()
        {
            var header = new GradientPanel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = new Padding(16, 15, 12, 12),
                StartColor = Color.FromArgb(25, 34, 67),
                EndColor = Color.FromArgb(76, 70, 214),
            };

            _logoImage = RibbonIcons.Create("ptHubIcon", 48);
            var logo = new PictureBox
            {
                Location = new Point(15, 18),
                Size = new Size(48, 48),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Transparent,
                Image = _logoImage,
            };
            header.Controls.Add(logo);

            var title = new SmoothLabel
            {
                Text = "Power Tools",
                Location = new Point(76, 16),
                Size = new Size(250, 31),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.Transparent,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 17f, FontStyle.Bold),
                AutoEllipsis = true,
            };
            header.Controls.Add(title);

            var hubPill = new SmoothLabel
            {
                Text = "HUB",
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(0, 19),
                Size = new Size(43, 21),
                Anchor = AnchorStyles.Top,
                BackColor = Color.FromArgb(84, 255, 255, 255),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 7.5f, FontStyle.Bold),
            };
            header.Controls.Add(hubPill);

            var subtitle = new SmoothLabel
            {
                Text = "One searchable home for every workflow",
                Location = new Point(77, 48),
                Size = new Size(290, 22),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(218, 224, 249),
                Font = new Font("Segoe UI", 9.2f),
                AutoEllipsis = true,
            };
            header.Controls.Add(subtitle);

            var hint = new SmoothLabel
            {
                Text = "Frequent commands stay on the ribbon. Everything else lives here.",
                Location = new Point(16, 78),
                Size = new Size(360, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(185, 196, 233),
                Font = new Font("Segoe UI", 8.3f),
                AutoEllipsis = true,
            };
            header.Controls.Add(hint);

            Action arrange = () =>
            {
                int width = header.ClientSize.Width;
                bool showPill = width >= 300;
                hubPill.Visible = showPill;
                hubPill.Left = Math.Max(0, width - hubPill.Width - 13);
                title.Width = Math.Max(90, width - title.Left - (showPill ? hubPill.Width + 25 : 12));
                subtitle.Width = Math.Max(90, width - subtitle.Left - 12);
                hint.Width = Math.Max(120, width - hint.Left - 14);
            };
            header.Resize += (_, __) => arrange();
            arrange();
            return header;
        }

        private Control BuildSearch()
        {
            var shell = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = new Padding(12, 9, 12, 7),
                BackColor = Background,
            };

            var box = new BorderedPanel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderColor = Border,
                Radius = 8,
                Padding = new Padding(38, 8, 34, 6),
            };
            shell.Controls.Add(box);

            var searchMark = new SmoothLabel
            {
                Text = "⌕",
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(9, 5),
                Size = new Size(24, 26),
                ForeColor = Color.FromArgb(93, 103, 127),
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI Symbol", 15f),
            };
            box.Controls.Add(searchMark);

            _search.BorderStyle = BorderStyle.None;
            _search.Dock = DockStyle.Fill;
            _search.Font = new Font("Segoe UI", 10f);
            _search.ForeColor = Ink;
            _search.BackColor = Color.White;
            _search.AccessibleName = "Search Power Tools";
            _search.TextChanged += (_, __) =>
            {
                _clearSearch.Visible = _search.TextLength > 0;
                if (_refreshingCatalog) return;
                if (_search.TextLength > 0 && _selectedCategory != "All")
                    SetCategoryVisual("All");
                RefreshCatalog();
            };
            box.Controls.Add(_search);
            SetCueBanner(_search, SearchPlaceholder);

            _clearSearch.Text = "×";
            _clearSearch.TextAlign = ContentAlignment.MiddleCenter;
            _clearSearch.FlatStyle = FlatStyle.Flat;
            _clearSearch.FlatAppearance.BorderSize = 0;
            _clearSearch.BackColor = Color.White;
            _clearSearch.ForeColor = Muted;
            _clearSearch.Font = new Font("Segoe UI", 12f);
            _clearSearch.Cursor = Cursors.Hand;
            _clearSearch.TabStop = false;
            _clearSearch.AccessibleName = "Clear search";
            _clearSearch.Visible = false;
            _clearSearch.Size = new Size(28, 28);
            _clearSearch.Location = new Point(0, 4);
            _clearSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            box.Resize += (_, __) => _clearSearch.Left = box.ClientSize.Width - _clearSearch.Width - 5;
            _clearSearch.Click += (_, __) =>
            {
                _search.Clear();
                _search.Focus();
            };
            box.Controls.Add(_clearSearch);
            _clearSearch.BringToFront();

            return shell;
        }

        private Control BuildCategoryBar()
        {
            _categories.Dock = DockStyle.Fill;
            _categories.Margin = Padding.Empty;
            _categories.Padding = new Padding(10, 3, 8, 4);
            _categories.BackColor = Background;
            _categories.WrapContents = true;
            _categories.AutoScroll = false;
            _categories.FlowDirection = FlowDirection.LeftToRight;

            AddCategory("All", 42);
            AddCategory("Compare", 66);
            AddCategory("Files", 48);
            AddCategory("Transform", 70);
            AddCategory("Data", 48);
            AddCategory("Utilities", 62);
            return _categories;
        }

        private Control BuildCatalog()
        {
            _catalog.Dock = DockStyle.Fill;
            _catalog.Margin = Padding.Empty;
            _catalog.Padding = new Padding(8, 4, 8, 10);
            _catalog.BackColor = Background;
            _catalog.AutoScroll = true;
            _catalog.WrapContents = false;
            _catalog.FlowDirection = FlowDirection.TopDown;
            _catalog.Resize += (_, __) => ResizeCatalogItems();
            return _catalog;
        }

        private Control BuildFooter()
        {
            var footer = new TopBorderPanel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                BackColor = Color.White,
                BorderColor = Border,
                Padding = new Padding(12, 0, 10, 0),
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.White,
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 76));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            footer.Controls.Add(layout);

            _context.Dock = DockStyle.Fill;
            _context.Margin = Padding.Empty;
            _context.TextAlign = ContentAlignment.MiddleLeft;
            _context.ForeColor = Muted;
            _context.Font = new Font("Segoe UI", 8.1f);
            _context.AutoEllipsis = true;
            layout.Controls.Add(_context, 0, 0);

            _resultCount.Dock = DockStyle.Fill;
            _resultCount.Margin = Padding.Empty;
            _resultCount.TextAlign = ContentAlignment.MiddleRight;
            _resultCount.ForeColor = Color.FromArgb(113, 121, 143);
            _resultCount.Font = new Font("Segoe UI Semibold", 8.1f, FontStyle.Bold);
            layout.Controls.Add(_resultCount, 1, 0);
            return footer;
        }

        private void AddCategory(string name, int width)
        {
            var chip = new CategoryChip
            {
                Text = name,
                Size = new Size(width, 31),
                Margin = new Padding(2, 1, 2, 1),
                Cursor = Cursors.Hand,
            };
            chip.Click += (_, __) => SelectCategory(name);
            _chips[name] = chip;
            _categories.Controls.Add(chip);
        }

        private void SelectCategory(string category)
        {
            SetCategoryVisual(category);
            RefreshCatalog();
        }

        private void SetCategoryVisual(string category)
        {
            _selectedCategory = category;
            foreach (KeyValuePair<string, CategoryChip> pair in _chips)
                pair.Value.Selected = string.Equals(pair.Key, category, StringComparison.OrdinalIgnoreCase);
        }

        private void RefreshCatalog()
        {
            if (_catalog == null || _catalog.IsDisposed) return;

            _refreshingCatalog = true;
            _catalog.SuspendLayout();
            try
            {
                string query = (_search.Text ?? "").Trim();
                int visible = 0;

                foreach (string section in SectionOrder)
                {
                    int sectionCount = 0;
                    foreach (ToolDefinition tool in _tools)
                    {
                        if (tool.Section != section) continue;
                        bool matches = (_selectedCategory == "All" || _selectedCategory == tool.Category) &&
                                       tool.Matches(query);
                        _cards[tool].Visible = matches;
                        if (matches) sectionCount++;
                    }
                    Label header = _sectionHeaders[section];
                    header.Text = SectionHeaderText(section, sectionCount);
                    header.Visible = sectionCount > 0;
                    visible += sectionCount;
                }

                _emptyState.Visible = visible == 0;
                if (visible == 0)
                    _emptyTitle.Text = string.IsNullOrEmpty(query)
                        ? "No tools in this category"
                        : "No tools match “" + query + "”";
                _resultCount.Text = visible == 1 ? "1 tool" : visible + " tools";
                ResizeCatalogItems();
            }
            finally
            {
                _catalog.ResumeLayout(true);
                _refreshingCatalog = false;
            }
        }

        private void InitializeCatalog()
        {
            _catalog.SuspendLayout();
            foreach (string section in SectionOrder)
            {
                int count = 0;
                foreach (ToolDefinition tool in _tools)
                    if (tool.Section == section) count++;

                Label header = MakeSectionHeader(section, count);
                _sectionHeaders[section] = header;
                _catalog.Controls.Add(header);

                foreach (ToolDefinition tool in _tools)
                {
                    if (tool.Section != section) continue;
                    var card = new ToolCard(tool, RunTool);
                    _cards[tool] = card;
                    _catalog.Controls.Add(card);
                }
            }
            _emptyState = MakeEmptyState();
            _emptyState.Visible = false;
            _catalog.Controls.Add(_emptyState);
            _catalog.ResumeLayout(true);
        }

        private static Label MakeSectionHeader(string section, int count)
        {
            return new SmoothLabel
            {
                Text = SectionHeaderText(section, count),
                Height = 28,
                Margin = new Padding(5, 7, 5, 0),
                Padding = new Padding(3, 6, 0, 0),
                ForeColor = Color.FromArgb(112, 121, 145),
                BackColor = Background,
                Font = new Font("Segoe UI Semibold", 7.5f, FontStyle.Bold),
                UseMnemonic = false,
            };
        }

        private static string SectionHeaderText(string section, int count)
        {
            string title;
            switch (section)
            {
                case "Comparison": title = "COMPARE & REVIEW"; break;
                case "Files": title = "DELIMITED FILES & WORKSPACES"; break;
                case "Transform": title = "TRANSFORM SELECTION"; break;
                case "Data": title = "CLEAN & ANALYZE"; break;
                case "History": title = "HISTORY"; break;
                case "Workbook": title = "WORKBOOK"; break;
                case "IoPipeline": title = "IO PIPELINE"; break;
                case "Settings": title = "BEHAVIOR SETTINGS"; break;
                default: title = "DEVELOPER & DIAGNOSTICS"; break;
            }
            return title + "   ·   " + count;
        }

        private Panel MakeEmptyState()
        {
            var panel = new Panel
            {
                Height = 150,
                Margin = new Padding(6, 18, 6, 6),
                BackColor = Background,
            };
            var icon = new SmoothLabel
            {
                Text = "⌕",
                Location = new Point(0, 12),
                Size = new Size(100, 42),
                Anchor = AnchorStyles.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(157, 165, 184),
                Font = new Font("Segoe UI Symbol", 25f),
            };
            panel.Controls.Add(icon);

            _emptyTitle = new SmoothLabel
            {
                Text = "No tools in this category",
                Location = new Point(10, 63),
                Size = new Size(180, 25),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Ink,
                Font = new Font("Segoe UI Semibold", 10f, FontStyle.Bold),
                AutoEllipsis = true,
            };
            panel.Controls.Add(_emptyTitle);

            var hint = new SmoothLabel
            {
                Text = "Try another keyword or choose All.",
                Location = new Point(10, 91),
                Size = new Size(180, 22),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Muted,
                Font = new Font("Segoe UI", 8.5f),
            };
            panel.Controls.Add(hint);
            panel.Resize += (_, __) =>
            {
                icon.Left = Math.Max(0, (panel.ClientSize.Width - icon.Width) / 2);
                _emptyTitle.Left = 10;
                _emptyTitle.Width = Math.Max(80, panel.ClientSize.Width - 20);
                hint.Left = 10;
                hint.Width = Math.Max(80, panel.ClientSize.Width - 20);
            };
            return panel;
        }

        private void ResizeCatalogItems()
        {
            if (_catalog.IsDisposed) return;
            int width = Math.Max(210,
                _catalog.ClientSize.Width - _catalog.Padding.Horizontal - 8);
            foreach (Control control in _catalog.Controls) control.Width = width;
        }

        private void UpdateResponsiveLayout()
        {
            if (_root.RowStyles.Count < 3) return;
            _categories.WrapContents = true;

            // Control dimensions have already been DPI-scaled by WinForms at this point. Deriving the row
            // count from those live dimensions avoids resetting an autoscaled task pane to raw 45/78px values
            // and keeps all chips visible even when Excel docks the pane very narrowly.
            int available = Math.Max(1,
                (_root.ClientSize.Width > 0 ? _root.ClientSize.Width : ClientSize.Width) -
                _categories.Padding.Horizontal);
            int rows = 1;
            int used = 0;
            int rowHeight = 0;
            foreach (Control chip in _categories.Controls)
            {
                if (!chip.Visible) continue;
                int width = chip.Width + chip.Margin.Horizontal;
                int height = chip.Height + chip.Margin.Vertical;
                if (used > 0 && used + width > available)
                {
                    rows++;
                    used = 0;
                }
                used += width;
                rowHeight = Math.Max(rowHeight, height);
            }
            _root.RowStyles[2].Height = _categories.Padding.Vertical + rows * Math.Max(1, rowHeight);
        }

        private void RunTool(ToolDefinition tool)
        {
            if (tool.PaneAction != null)
            {
                try { tool.PaneAction(); }
                catch (Exception ex) { Dialogs.Info(tool.Title + " could not be opened:\n" + ex.Message); }
                return;
            }
            if (tool.Action == null) return;
            ToolExecutionContext context = null;
            if (tool.RequiresStableContext)
            {
                context = ToolExecutionContext.Capture(_ownerWindowHandle);
                if (context == null)
                {
                    Dialogs.Info("Excel is busy or this Hub is not the active window.\n\n" +
                                 "Return to the intended workbook and try " + tool.Title + " again.");
                    return;
                }
            }

            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    if (context != null && !context.MatchesCurrent())
                    {
                        Dialogs.Info("The workbook, sheet, or selection changed before " + tool.Title +
                                     " could run. Nothing was changed.\n\nSelect the intended cells and try again.");
                        return;
                    }
                    try { tool.Action(); }
                    catch (Exception ex) { Dialogs.Info(tool.Title + " failed:\n" + ex.Message); }
                });
            }
            catch (Exception ex) { Dialogs.Info(tool.Title + " could not be queued:\n" + ex.Message); }
        }

        private void QueuePaneCommand(string title, Action action)
        {
            if (action == null || _detached || IsDisposed) return;
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    if (_detached || IsDisposed) return;
                    try { action(); }
                    catch (Exception ex) { Dialogs.Info(title + " failed:\n" + ex.Message); }
                });
            }
            catch (Exception ex) { Dialogs.Info(title + " could not be queued:\n" + ex.Message); }
        }

        private void ShowComparisonWorkspace()
        {
            _root.Visible = false;
            _comparisonWorkspace.Visible = true;
            _comparisonWorkspace.BringToFront();
            _comparisonWorkspace.FocusPageStart();
        }

        internal void OpenComparisonReviewPicker()
        {
            if (_detached || IsDisposed) return;
            ShowComparisonWorkspace();
            QueuePaneCommand("Validate Review workbook", SelectReviewForWorkspace);
        }

        private void SelectReviewForWorkspace()
        {
            if (_detached || IsDisposed || _comparisonWorkspace.IsDisposed) return;
            string path = Dialogs.OpenFile(ComparisonCommands.ReviewWorkbookFilter);
            if (path == null) return;
            _comparisonWorkspace.ClearLoadedReview();
            try
            {
                ReviewWorkbookHydrationSession session = ComparisonBuildProgressDialog.Run(
                    ExcelWindow.Current,
                    new ComparisonBuildProgressDialogOptions
                    {
                        Title = "Validate Review workbook",
                        Heading = "Checking Review.xlsx integrity",
                        InitialPhase = "Opening the package safely",
                        InitialDetail = "Verifying the manifest, state chunks, and visible projection before Excel opens.",
                        InitialCanCancel = false,
                        NonCancelableClosingPhase = "Finishing integrity validation...",
                        NonCancelableClosingDetail =
                            "This window can close after the selected Review workbook has been checked safely.",
                    },
                    async (progress, cancellationToken) =>
                    {
                        progress.Report(new ComparisonBuildProgressUpdate(
                            "Validating Review.xlsx",
                            "Checking the package envelope, embedded state digest, and visible projection digest.",
                            null,
                            false));
                        ReviewWorkbookHydrationSession hydrated = await Task.Run(
                            () => ReviewWorkbookHydrationService.Hydrate(path)).ConfigureAwait(false);
                        progress.Report(new ComparisonBuildProgressUpdate(
                            "Integrity verified",
                            "Preparing the in-pane validation summary.",
                            100,
                            false));
                        return hydrated;
                    });
                if (_detached || IsDisposed || _comparisonWorkspace.IsDisposed) return;
                _comparisonWorkspace.ShowHydratedReview(session);
            }
            catch (Exception exception)
            {
                if (!_comparisonWorkspace.IsDisposed) _comparisonWorkspace.ClearLoadedReview();
                Dialogs.Info(
                    "This file could not be validated as a Power Tools comparison Review. Nothing was opened.\n\n" +
                    ComparisonCommands.DescribeException(exception));
            }
        }

        private void OpenHydratedReviewForWorkspace(ReviewWorkbookHydrationSession selected)
        {
            if (selected == null) throw new ArgumentNullException(nameof(selected));
            if (_detached || IsDisposed || _comparisonWorkspace.IsDisposed) return;
            if (!ReferenceEquals(_comparisonWorkspace.LoadedSession, selected))
                throw new InvalidOperationException(
                    "The loaded comparison project changed before the protected-open request ran. Select it again.");

            try
            {
                ReviewWorkbookHydrationSession current =
                    ComparisonCommands.OpenValidatedReviewInExcel(selected);
                if (_detached || IsDisposed || _comparisonWorkspace.IsDisposed) return;
                if (ReferenceEquals(_comparisonWorkspace.LoadedSession, selected))
                    _comparisonWorkspace.ShowHydratedReview(current);
            }
            catch
            {
                if (!_comparisonWorkspace.IsDisposed)
                    _comparisonWorkspace.ClearLoadedReview(selected);
                throw;
            }
        }

        private void CommitDecisionRevisionForWorkspace(
            ComparisonDecisionCommitRequestedEventArgs request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            if (_detached || IsDisposed || _comparisonWorkspace.IsDisposed) return;
            ReviewWorkbookHydrationSession selected = request.Session;
            if (!ReferenceEquals(_comparisonWorkspace.LoadedSession, selected))
                throw new InvalidOperationException(
                    "The loaded comparison project changed before the pending decisions could be saved. " +
                    "Nothing was written.");

            string sourceDirectory = Path.GetDirectoryName(selected.FullPath) ?? "";
            string sourceName = Path.GetFileNameWithoutExtension(selected.FullPath);
            string suggestedName = sourceName + ".review-" +
                                   DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) +
                                   ".xlsx";
            string suggestedPath = sourceDirectory.Length == 0
                ? suggestedName
                : Path.Combine(sourceDirectory, suggestedName);
            string destination = Dialogs.SaveFile(
                ComparisonCommands.ReviewWorkbookFilter,
                suggestedPath);
            if (destination == null) return;
            destination = Path.GetFullPath(destination);
            if (File.Exists(destination) || Directory.Exists(destination))
                throw new IOException(
                    "The selected output already exists. Choose a new Review revision name; existing files are never overwritten.");

            ReviewWorkbookHydrationSession revised = ComparisonBuildProgressDialog.Run(
                ExcelWindow.Current,
                new ComparisonBuildProgressDialogOptions
                {
                    Title = "Save review decisions",
                    Heading = "Creating a new immutable Review revision",
                    InitialPhase = "Validating pending decisions",
                    InitialDetail =
                        "The loaded source remains unchanged while the decision ledger is replayed and verified.",
                    InitialCanCancel = false,
                    NonCancelableClosingPhase = "Finishing Review revision...",
                    NonCancelableClosingDetail =
                        "The window can close after the new package and decision history are fully validated.",
                },
                async (progress, cancellationToken) =>
                {
                    progress.Report(new ComparisonBuildProgressUpdate(
                        "Replaying decision history",
                        request.Requests.Count.ToString("N0", CultureInfo.CurrentCulture) +
                        " pending decision event(s) will be written into one new Review workbook.",
                        25,
                        false));
                    ReviewWorkbookHydrationSession hydrated = await Task.Run(() =>
                    {
                        ReviewWorkbookFile.CreateDecisionRevisionBatch(
                            selected,
                            destination,
                            request.Requests);
                        return ReviewWorkbookHydrationService.Hydrate(destination);
                    }).ConfigureAwait(false);
                    progress.Report(new ComparisonBuildProgressUpdate(
                        "Revision verified",
                        "The new Review workbook reproduces the complete immutable decision generation.",
                        100,
                        false));
                    return hydrated;
                });

            if (!string.Equals(
                    revised.ReadResult.DecisionGeneration?.GenerationId,
                    request.ExpectedFinalGenerationId,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The saved Review revision did not reproduce the expected pending decision generation.");
            if (_detached || IsDisposed || _comparisonWorkspace.IsDisposed) return;
            if (!ReferenceEquals(_comparisonWorkspace.LoadedSession, selected))
            {
                Dialogs.Info(
                    "The new Review revision was created and verified, but another project is now loaded.\n\n" +
                    destination);
                return;
            }
            _comparisonWorkspace.ShowHydratedReview(revised);
        }

        private void ShowToolsCatalog()
        {
            _comparisonWorkspace.Visible = false;
            _root.Visible = true;
            _root.BringToFront();
            if (_search.CanFocus) _search.Focus();
        }

        private void SetSearchPlaceholder()
        {
            if (_search.TextLength > 0) _search.Clear();
            _search.ForeColor = Ink;
            _clearSearch.Visible = false;
            SetCueBanner(_search, SearchPlaceholder);
        }

        private static void SetCueBanner(TextBox textBox, string text)
        {
            try { SendMessage(textBox.Handle, EmSetCueBanner, new IntPtr(1), text); }
            catch { }
        }

        private const int EmSetCueBanner = 0x1501;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        private void OnExcelContextChanged()
        {
            if (_detached || IsDisposed || !IsHandleCreated) return;
            if (!OwnsActiveWindow()) return;
            try { BeginInvoke((Action)RefreshContext); } catch { }
        }

        private void RefreshContext()
        {
            if (_detached || IsDisposed) return;
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic activeWindow = app?.ActiveWindow;
                if (_ownerWindowHandle != 0 &&
                    (activeWindow == null || Convert.ToInt32(activeWindow.Hwnd) != _ownerWindowHandle)) return;
                dynamic workbook = app?.ActiveWorkbook;
                dynamic sheet = app?.ActiveSheet;
                string workbookName = workbook == null ? "No workbook" : (string)workbook.Name;
                string sheetName = sheet == null ? "" : (string)sheet.Name;
                _context.Text = string.IsNullOrEmpty(sheetName)
                    ? workbookName
                    : workbookName + "  ›  " + sheetName;
            }
            catch { _context.Text = "Excel context unavailable"; }
        }

        private bool OwnsActiveWindow()
        {
            if (_ownerWindowHandle == 0) return true;
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic window = app?.ActiveWindow;
                return window != null && Convert.ToInt32(window.Hwnd) == _ownerWindowHandle;
            }
            catch { return false; }
        }

        private List<ToolDefinition> BuildToolCatalog()
        {
            Color comparison = Color.FromArgb(104, 70, 198);
            Color files = Color.FromArgb(0, 137, 165);
            Color transform = Color.FromArgb(232, 100, 48);
            Color data = Color.FromArgb(36, 151, 104);
            Color history = Color.FromArgb(193, 63, 120);
            Color utility = Color.FromArgb(94, 82, 190);
            Color developer = Color.FromArgb(70, 82, 105);

            return new List<ToolDefinition>
            {
                ToolDefinition.Navigation("Compare", "Compare revisions",
                    "Build a multi-revision project with editable lineage, aligned logical changes and a Review workbook.",
                    "TREE", comparison, ShowComparisonWorkspace,
                    "spreadsheet workbook revisions differences history branch conflict formula value review xlsx xlsm",
                    "ALPHA", section: "Comparison"),
                ToolDefinition.Command("Compare", "Exact Compare (Preview)",
                    "Compare exactly two files positionally and generate an integrity-checked Review workbook.",
                    "A↔B", comparison, ComparisonCommands.StartExactPreview,
                    "spreadsheet workbook revisions differences formula value review xlsx xlsm csv tsv",
                    "PREVIEW", requiresStableContext: false, section: "Comparison"),
                ToolDefinition.Navigation("Compare", "Open Review workbook",
                    "Validate a Review.xlsx in the Comparison workspace, then continue its decision ledger or open a protected copy.",
                    "OPEN", comparison, OpenComparisonReviewPicker,
                    "resume comparison project report decisions", section: "Comparison"),

                ToolDefinition.Command("Files", "Open files", "Open CSV, TSV or separated-value TXT files in one workbook.",
                    "CSV", files, CsvCommands.OpenCsv, "import multiple delimited", requiresStableContext: false),
                ToolDefinition.Command("Files", "Open whole folder", "Load every supported delimited file from a folder.",
                    "DIR", files, CsvCommands.OpenCsvFolder, "bulk directory tsv txt", requiresStableContext: false),
                ToolDefinition.Command("Files", "Save active source", "Write this source sheet back safely with conflict protection.",
                    "SAVE", files, CsvCommands.SaveCsv, "csv current backup"),
                ToolDefinition.Command("Files", "Save all sources", "Save every edited source sheet in the current workspace.",
                    "ALL", files, CsvCommands.SaveAllCsv, "workspace bulk"),
                ToolDefinition.Command("Files", "Workspace index", "Refresh file links, source paths, timestamps, formats and status.",
                    "IDX", files, CsvCommands.ShowIndex, "metadata links modified path"),
                ToolDefinition.Command("Files", "Export sheet", "Export the active worksheet as a delimited file.",
                    "OUT", files, CsvCommands.ExportSheetAsCsv, "csv separator save as"),
                ToolDefinition.Command("Files", "Re-style source view", "Re-detect the header and reapply workspace styling.",
                    "VIEW", files, CsvCommands.RestyleActiveSheet, "format header freeze"),
                ToolDefinition.Command("Files", "Backup source file", "Create a timestamped copy of the active source file.",
                    "BAK", files, CsvCommands.BackupActiveFile, "copy safety"),

                ToolDefinition.Command("Transform", "Keep values only", "Replace formulas with current values in visible selected cells.",
                    "123", transform, SelectionCommands.KeepValuesOnly, "paste formulas ctrl q"),
                ToolDefinition.Command("Transform", "Smart fill down", "Increment trailing numbers or IP octets; two seeds set the step.",
                    "+1", transform, FillCommands.SmartFillDown, "series seed ctrl shift d"),
                ToolDefinition.Command("Transform", "Change delimited field", "Edit one field inside every selected cell with saved parameters.",
                    "Aa", transform, FieldEditorCommands.ChangeField, "separator index color"),
                ToolDefinition.Command("Transform", "Regex replace", "Preview regex or plain-text replacements before applying them.",
                    ".*", transform, RegexReplaceCommands.ShowRegexReplace, "find replace capture groups"),

                ToolDefinition.Command("Data", "Column statistics", "Profile counts, distinct values, numeric range and top values live.",
                    "Σ", data, DataCommands.ShowColumnStats, "profile frequency min max"),
                ToolDefinition.Command("Data", "Natural sort ascending", "Stable A–Z sort that places DI2 before DI10.",
                    "A↑", data, DataCommands.SortAscending, "order low high"),
                ToolDefinition.Command("Data", "Natural sort descending", "Stable Z–A natural sort of the active data block.",
                    "Z↓", data, DataCommands.SortDescending, "order high low"),
                ToolDefinition.Command("Data", "Deduplicate rows", "Remove duplicate rows using the selected columns as keys.",
                    "1×", data, DataCommands.DedupRows, "unique duplicates key"),

                ToolDefinition.Command("Utilities", "Timeline", "Inspect tracked manual edits and Power Tools operations; jump to any step.",
                    "HIST", history, UndoCommands.ShowTimeline, "history manual edits",
                    requiresStableContext: false, section: "History"),
                ToolDefinition.Command("Utilities", "Undo Power Tools", "Undo the newest step in the Power Tools timeline.",
                    "↶", history, UndoCommands.UndoLast, "revert ctrl alt z",
                    requiresStableContext: false, section: "History"),
                ToolDefinition.Command("Utilities", "Redo Power Tools", "Redo the newest undone timeline step.",
                    "↷", history, UndoCommands.RedoLast, "restore ctrl alt y",
                    requiresStableContext: false, section: "History"),
                ToolDefinition.Command("Utilities", "Split workbook into files", "Export the workbook in groups of a chosen number of sheets.",
                    "SPLIT", utility, WorkbookCommands.ExportSheetsAsFiles, "workbook sheets export", section: "Workbook"),

                ToolDefinition.Command("Utilities", "Build IO //Address", "Add the helper address column to an Inputs/Outputs source.",
                    "ADDR", utility, IoPipeline.IoListCommands.BuildAddressColumn, "node offset bit", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "Pull IO-list data", "Match addresses and fill mnemonic, ID and description fields.",
                    "IO", utility, IoPipeline.IoListCommands.PullFromIoList, "plant card board number", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "Pull mnemonics", "Fill the //Mnemonic helper on an R.S.C card sheet from plant source files.",
                    "MN", utility, IoPipeline.IoListCommands.PullMnemonicsIntoIoList, "io list bridge card sheet", section: "IoPipeline"),

                ToolDefinition.Command("Utilities", "IO Bridge: open workspace", "Join an IO list with its plant CSVs; edits queue and apply in batch.",
                    "BR", utility, IoPipeline.IoBridge.IoBridgeCommands.OpenWorkspace, "io bridge workspace queue bidirectional", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "IO Bridge: row actions", "Queue linked-field updates, integrate, disable/enable or delete for the selected row.",
                    "BR▸", utility, IoPipeline.IoBridge.IoBridgeCommands.RowActions, "io bridge integrate copy across disable delete", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "IO Bridge: apply batch", "Write the included pending actions to the CSVs and the IO-list workbook.",
                    "BR✔", utility, IoPipeline.IoBridge.IoBridgeCommands.ApplyBatch, "io bridge apply write batch", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "IO Bridge: refresh", "Rebuild the bridge sheets from the current documents; the queue re-anchors.",
                    "BR↻", utility, IoPipeline.IoBridge.IoBridgeCommands.RefreshWorkspace, "io bridge rebuild reanchor", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "IO Bridge: column config", "Save the _Columns sheet to the registry and regroup the workspace.",
                    "BR⚙", utility, IoPipeline.IoBridge.IoBridgeCommands.ApplyColumnConfig, "io bridge columns groups hidden", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "IO Bridge: clear history", "Drop applied actions from the queue, keeping pending and conflicting ones.",
                    "BR🗑", utility, IoPipeline.IoBridge.IoBridgeCommands.ClearHistory, "io bridge purge history", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "IO Bridge: expanded view", "Unfold a multi-mnemonic point into one temporary row per constituent.",
                    "BR⤢", utility, IoPipeline.IoBridge.IoBridgeCommands.ExpandView, "io bridge expand multi mnemonic", section: "IoPipeline"),

                ToolDefinition.Command("Utilities", "SWP Bridge: open workspace", "Diff the IO list's diagnostic mnemonics against the SWP workbook.",
                    "SW", utility, IoPipeline.SwpBridge.SwpBridgeCommands.OpenWorkspace, "swp bridge diagnostic alarms warnings diff", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "SWP Bridge: apply batch", "Patch the SWP headlessly — in place with .bak, or into a new dated version folder.",
                    "SW✔", utility, IoPipeline.SwpBridge.SwpBridgeCommands.ApplyBatch, "swp bridge apply write patch", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "SWP Bridge: refresh", "Re-read the IO list and SWP, recompute the diff, re-render.",
                    "SW↻", utility, IoPipeline.SwpBridge.SwpBridgeCommands.RefreshWorkspace, "swp bridge rebuild recompute", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "SWP Bridge: preview block", "Show the selected row's SWP section with the incoming values highlighted.",
                    "SW⤢", utility, IoPipeline.SwpBridge.SwpBridgeCommands.PreviewSelection, "swp bridge preview block section", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "SWP Bridge: dashboard", "Progress, live tallies, per-sheet breakdown and the findings tail.",
                    "SW▦", utility, IoPipeline.SwpBridge.SwpBridgeCommands.ShowDashboard, "swp bridge dashboard tallies progress", section: "IoPipeline"),
                ToolDefinition.Command("Utilities", "SWP Bridge: clear history", "Drop applied actions from the queue, keeping pending and conflicting ones.",
                    "SW🗑", utility, IoPipeline.SwpBridge.SwpBridgeCommands.ClearHistory, "swp bridge purge history", section: "IoPipeline"),

                ToolDefinition.Setting("Utilities", "Track manual edits", "Include typing, delete and paste-over-selection in the timeline.",
                    "REC", history, () => ManualEditTracker.Enabled, ManualEditTracker.SetEnabled, "undo setting", "Settings"),
                ToolDefinition.Setting("Utilities", "Intercept native delimited opens", "Route files Excel opens natively through the safe workspace pipeline.",
                    "SAFE", files, CsvCommands.InterceptOpenEnabled, CsvCommands.SetInterceptOpen, "csv tsv txt setting", "Settings"),

                ToolDefinition.Command("Utilities", "Open solution in Visual Studio", "Open the add-in solution in the configured Visual Studio.",
                    "VS", developer, DevCommands.OpenSolutionInVisualStudio, "code framework", "DEVELOPER",
                    false, "Developer"),
                ToolDefinition.Command("Utilities", "Open repository folder", "Show the add-in repository in File Explorer.",
                    "<>", developer, DevCommands.OpenRepoFolder, "source code", "DEVELOPER",
                    false, "Developer"),
                ToolDefinition.Command("Utilities", "Open build output", "Show the folder containing the currently loaded add-in build.",
                    "BIN", developer, DevCommands.OpenBuildOutputFolder, "xll release debug", "DEVELOPER",
                    false, "Developer"),
                ToolDefinition.Command("Utilities", "About Power Tools", "Show version, loaded XLL, solution and repository details.",
                    "i", developer, DevCommands.About, "information path version",
                    requiresStableContext: false, section: "Developer"),
            };
        }

        private sealed class ToolExecutionContext
        {
            private int WindowHandle;
            private string WorkbookName;
            private string SheetIdentity;
            private string SelectionAddress;

            public static ToolExecutionContext Capture(int expectedWindowHandle)
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic window = app?.ActiveWindow;
                    if (window == null) return null;
                    int handle = Convert.ToInt32(window.Hwnd);
                    if (expectedWindowHandle != 0 && handle != expectedWindowHandle) return null;

                    dynamic workbook = app.ActiveWorkbook;
                    dynamic sheet = app.ActiveSheet;
                    dynamic selection = app.Selection;
                    var context = new ToolExecutionContext { WindowHandle = handle };
                    try { context.WorkbookName = workbook == null ? null : (string)workbook.Name; } catch { }
                    try { context.SheetIdentity = sheet == null ? null : (string)sheet.CodeName; }
                    catch { try { context.SheetIdentity = sheet == null ? null : (string)sheet.Name; } catch { } }
                    try { context.SelectionAddress = selection == null ? null : (string)selection.Address(false, false); }
                    catch { context.SelectionAddress = null; }
                    return context;
                }
                catch { return null; }
            }

            public bool MatchesCurrent()
            {
                ToolExecutionContext current = Capture(WindowHandle);
                return current != null &&
                       string.Equals(WorkbookName, current.WorkbookName, StringComparison.Ordinal) &&
                       string.Equals(SheetIdentity, current.SheetIdentity, StringComparison.Ordinal) &&
                       string.Equals(SelectionAddress, current.SelectionAddress, StringComparison.Ordinal);
            }
        }

        private sealed class ToolDefinition
        {
            public string Category;
            public string Section;
            public string Title;
            public string Description;
            public string IconText;
            public Color AccentColor;
            public Action Action;
            public Action PaneAction;
            public string Keywords;
            public string Badge;
            public bool RequiresStableContext;
            public Func<bool> SettingGetter;
            public Action<bool> SettingSetter;

            public bool IsSetting => SettingGetter != null && SettingSetter != null;

            public static ToolDefinition Command(string category, string title, string description,
                string iconText, Color accent, Action action, string keywords, string badge = null,
                bool requiresStableContext = true, string section = null)
            {
                return new ToolDefinition
                {
                    Category = category,
                    Section = section ?? category,
                    Title = title,
                    Description = description,
                    IconText = iconText,
                    AccentColor = accent,
                    Action = action,
                    Keywords = keywords,
                    Badge = badge,
                    RequiresStableContext = requiresStableContext,
                };
            }

            public static ToolDefinition Navigation(string category, string title, string description,
                string iconText, Color accent, Action paneAction, string keywords, string badge = null,
                string section = null)
            {
                if (paneAction == null) throw new ArgumentNullException(nameof(paneAction));
                return new ToolDefinition
                {
                    Category = category,
                    Section = section ?? category,
                    Title = title,
                    Description = description,
                    IconText = iconText,
                    AccentColor = accent,
                    PaneAction = paneAction,
                    Keywords = keywords,
                    Badge = badge,
                    RequiresStableContext = false,
                };
            }

            public static ToolDefinition Setting(string category, string title, string description,
                string iconText, Color accent, Func<bool> getter, Action<bool> setter, string keywords,
                string section = null)
            {
                return new ToolDefinition
                {
                    Category = category,
                    Section = section ?? category,
                    Title = title,
                    Description = description,
                    IconText = iconText,
                    AccentColor = accent,
                    SettingGetter = getter,
                    SettingSetter = setter,
                    Keywords = keywords,
                    Badge = "SETTING",
                };
            }

            public bool Matches(string query)
            {
                if (string.IsNullOrWhiteSpace(query)) return true;
                string haystack = Title + " " + Description + " " + Keywords + " " + Badge;
                string[] terms = query.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string term in terms)
                    if (haystack.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0) return false;
                return true;
            }
        }

        private sealed class ToolCard : Panel
        {
            private readonly ToolDefinition _tool;
            private readonly Action<ToolDefinition> _run;
            private readonly Label _title;
            private readonly Label _description;
            private readonly Label _badge;
            private readonly IconTile _icon;
            private readonly ToggleSwitch _toggle;
            private bool _hovered;
            private bool _syncingToggle;

            public ToolCard(ToolDefinition tool, Action<ToolDefinition> run)
            {
                _tool = tool;
                _run = run;
                Height = 78;
                Margin = new Padding(4, 4, 4, 4);
                BackColor = Color.Transparent;
                Cursor = Cursors.Hand;
                DoubleBuffered = true;
                SetStyle(ControlStyles.Selectable, true);
                TabStop = true;
                AccessibleName = tool.Title;
                AccessibleDescription = tool.Description;
                AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
                if (tool.IsSetting) AccessibleDefaultActionDescription = "Toggle setting";

                _icon = new IconTile(tool.IconText, tool.AccentColor)
                {
                    Location = new Point(12, 15),
                    Size = new Size(46, 46),
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_icon);

                _title = new SmoothLabel
                {
                    Text = tool.Title,
                    Location = new Point(70, 10),
                    Height = 23,
                    ForeColor = Ink,
                    BackColor = Color.Transparent,
                    Font = new Font("Segoe UI Semibold", 10f, FontStyle.Bold),
                    AutoEllipsis = true,
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_title);

                _description = new SmoothLabel
                {
                    Text = tool.Description,
                    Location = new Point(70, 34),
                    Height = 34,
                    ForeColor = Muted,
                    BackColor = Color.Transparent,
                    Font = new Font("Segoe UI", 8.2f),
                    AutoEllipsis = true,
                    Cursor = Cursors.Hand,
                };
                Controls.Add(_description);

                if (tool.IsSetting)
                {
                    _toggle = new ToggleSwitch
                    {
                        Size = new Size(42, 23),
                        Location = new Point(0, 12),
                        Anchor = AnchorStyles.Top | AnchorStyles.Right,
                        Cursor = Cursors.Hand,
                        TabStop = false,
                        AccessibleName = tool.Title,
                    };
                    _toggle.CheckedChanged += (_, __) =>
                    {
                        if (_syncingToggle) return;
                        try { _tool.SettingSetter(_toggle.Checked); }
                        catch (Exception ex) { Dialogs.Info("Couldn't change " + _tool.Title + ":\n" + ex.Message); }
                        RefreshSetting();
                    };
                    Controls.Add(_toggle);
                    RefreshSetting();
                }
                else if (!string.IsNullOrEmpty(tool.Badge))
                {
                    _badge = new SmoothLabel
                    {
                        Text = tool.Badge,
                        Size = new Size(66, 19),
                        Location = new Point(0, 12),
                        Anchor = AnchorStyles.Top | AnchorStyles.Right,
                        TextAlign = ContentAlignment.MiddleCenter,
                        ForeColor = tool.AccentColor,
                        BackColor = Color.FromArgb(239, 239, 250),
                        Font = new Font("Segoe UI Semibold", 6.6f, FontStyle.Bold),
                        Cursor = Cursors.Hand,
                    };
                    Controls.Add(_badge);
                }

                BindInteractive(this);
                KeyDown += (_, e) =>
                {
                    if (e.KeyCode != Keys.Enter && e.KeyCode != Keys.Space) return;
                    ActivateTool();
                    e.Handled = e.SuppressKeyPress = true;
                };
                GotFocus += (_, __) => Invalidate();
                LostFocus += (_, __) => Invalidate();
                Resize += (_, __) => Arrange();
                Arrange();
            }

            public void RefreshSetting()
            {
                if (_toggle == null) return;
                bool value = false;
                try { value = _tool.SettingGetter(); } catch { }
                _syncingToggle = true;
                _toggle.Checked = value;
                _syncingToggle = false;
                _toggle.Invalidate();
                AccessibleDescription = (value ? "On. " : "Off. ") + _tool.Description;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = Rounded(bounds, 9))
                using (var fill = new SolidBrush(_hovered ? Color.FromArgb(250, 251, 255) : Color.White))
                using (var border = new Pen(_hovered ? Color.FromArgb(176, 181, 245) : Border, _hovered ? 1.35f : 1f))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }
                if (Focused)
                {
                    Rectangle focus = Rectangle.Inflate(bounds, -4, -4);
                    ControlPaint.DrawFocusRectangle(e.Graphics, focus, Ink, Color.White);
                }
            }

            protected override void OnResize(EventArgs eventargs)
            {
                base.OnResize(eventargs);
                Invalidate();
            }

            private void Arrange()
            {
                int reserve = 14;
                if (_toggle != null)
                {
                    _toggle.Left = Width - _toggle.Width - 13;
                    reserve = _toggle.Width + 20;
                }
                else if (_badge != null)
                {
                    _badge.Left = Width - _badge.Width - 12;
                    reserve = _badge.Width + 18;
                }
                int textWidth = Math.Max(70, Width - 70 - reserve);
                _title.Width = textWidth;
                _description.Width = Math.Max(80, Width - 82);
            }

            private void BindInteractive(Control control)
            {
                if (control != _toggle)
                {
                    control.MouseEnter += (_, __) => SetHovered(true);
                    control.MouseLeave += (_, __) => SetHovered(ClientRectangle.Contains(PointToClient(Cursor.Position)));
                    control.MouseDown += (_, __) => Focus();
                    control.Click += (_, __) => ActivateTool();
                }
                foreach (Control child in control.Controls)
                    if (child != _toggle) BindInteractive(child);
            }

            private void SetHovered(bool hovered)
            {
                if (_hovered == hovered) return;
                _hovered = hovered;
                Invalidate();
            }

            private void ActivateTool()
            {
                if (_tool.IsSetting)
                {
                    _toggle.Checked = !_toggle.Checked;
                    return;
                }
                _run(_tool);
            }

            internal static GraphicsPath Rounded(Rectangle bounds, int radius)
            {
                int d = radius * 2;
                var path = new GraphicsPath();
                path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }

        private sealed class IconTile : Control
        {
            private readonly string _text;
            private readonly Color _color;

            public IconTile(string text, Color color)
            {
                _text = text;
                _color = color;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
                BackColor = Color.Transparent;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle bounds = new Rectangle(0, 0, Width - 1, Height - 1);
                using (GraphicsPath path = ToolCard.Rounded(bounds, 9))
                using (var fill = new SolidBrush(Color.FromArgb(24, _color)))
                using (var border = new Pen(Color.FromArgb(54, _color)))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }

                float size = _text.Length >= 5 ? 7.2f : _text.Length >= 3 ? 8.5f : 11f;
                using (var font = new Font("Segoe UI Semibold", size, FontStyle.Bold))
                using (var brush = new SolidBrush(_color))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    e.Graphics.DrawString(_text, font, brush, ClientRectangle, format);
            }
        }

        private sealed class ToggleSwitch : CheckBox
        {
            public ToggleSwitch()
            {
                Appearance = Appearance.Button;
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                Text = "";
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Color track = Checked ? Accent : Color.FromArgb(190, 197, 211);
                RectangleF bounds = new RectangleF(1, 2, Width - 2, Height - 4);
                using (GraphicsPath path = Pill(bounds))
                using (var brush = new SolidBrush(track))
                    e.Graphics.FillPath(brush, path);

                float diameter = Height - 8;
                float x = Checked ? Width - diameter - 4 : 4;
                using (var knob = new SolidBrush(Color.White))
                    e.Graphics.FillEllipse(knob, x, 4, diameter, diameter);
            }

            private static GraphicsPath Pill(RectangleF bounds)
            {
                float d = bounds.Height;
                var path = new GraphicsPath();
                path.AddArc(bounds.Left, bounds.Top, d, d, 90, 180);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 180);
                path.CloseFigure();
                return path;
            }
        }

        private sealed class CategoryChip : Button
        {
            private bool _selected;

            public bool Selected
            {
                get => _selected;
                set { _selected = value; Invalidate(); }
            }

            public CategoryChip()
            {
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                BackColor = Color.Transparent;
                ForeColor = Muted;
                Font = new Font("Segoe UI Semibold", 8.2f, FontStyle.Bold);
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = ToolCard.Rounded(bounds, 8))
                using (var fill = new SolidBrush(_selected ? Accent : Color.White))
                using (var border = new Pen(_selected ? Accent : Border))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }
                using (var brush = new SolidBrush(_selected ? Color.White : Muted))
                using (var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap,
                })
                    e.Graphics.DrawString(Text, Font, brush, bounds, format);
                if (Focused)
                {
                    Rectangle focus = Rectangle.Inflate(bounds, -4, -4);
                    ControlPaint.DrawFocusRectangle(e.Graphics, focus,
                        _selected ? Color.White : Ink, _selected ? Accent : Color.White);
                }
            }
        }

        private sealed class GradientPanel : Panel
        {
            public Color StartColor { get; set; }
            public Color EndColor { get; set; }

            public GradientPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            }

            protected override void OnPaintBackground(PaintEventArgs e)
            {
                using (var brush = new LinearGradientBrush(ClientRectangle, StartColor, EndColor, 12f))
                    e.Graphics.FillRectangle(brush, ClientRectangle);
            }
        }

        private sealed class SmoothLabel : Label
        {
            public SmoothLabel()
            {
                // GDI labels can fall back to visibly jagged glyphs over transparent/custom-painted hosts.
                // Keep this local to the Hub; changing WinForms' process-wide rendering mode would also affect
                // Excel and every other add-in hosted in the process.
                UseCompatibleTextRendering = true;
                UseMnemonic = false;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                base.OnPaint(e);
            }
        }

        private sealed class BorderedPanel : Panel
        {
            public Color BorderColor { get; set; }
            public int Radius { get; set; }

            public BorderedPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = ClientRectangle;
                bounds.Width--;
                bounds.Height--;
                using (GraphicsPath path = ToolCard.Rounded(bounds, Radius))
                using (var pen = new Pen(BorderColor))
                    e.Graphics.DrawPath(pen, path);
            }
        }

        private sealed class TopBorderPanel : Panel
        {
            public Color BorderColor { get; set; }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(BorderColor)) e.Graphics.DrawLine(pen, 0, 0, Width, 0);
            }
        }
    }
}
