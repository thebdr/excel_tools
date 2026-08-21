using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelPowerTools.Interop
{
    /// <summary>Message boxes owned by the Excel window (the add-in replacement for VBA's MsgBox).</summary>
    internal static class Dialogs
    {
        public const string Title = "Power Tools";

        /// <summary>
        /// Asks the user to pick one of <paramref name="options"/>. Returns the chosen index, or -1 if
        /// cancelled. With a single option the caller should not be asking — this still prompts.
        /// The window sizes itself to the longest option and is user-resizable.
        /// </summary>
        public static int SelectOne(string prompt, string[] options, string caption = null)
        {
            if (options == null || options.Length == 0) return -1;

            // Two points larger than the system dialog font — these lists carry dense signal names.
            using (var font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 2f))
            {
                // Wide enough for the longest option (within reason) so no label is cut off.
                int longest = 0;
                foreach (string option in options)
                    longest = System.Math.Max(longest, TextRenderer.MeasureText(option ?? "", font).Width);
                int width = System.Math.Max(520, System.Math.Min(longest + 80, 1120));

                return SelectOneCore(prompt, options, caption, font, width);
            }
        }

        private static int SelectOneCore(string prompt, string[] options, string caption, Font font, int width)
        {
            using (var form = new Form
            {
                Text = caption ?? Title,
                FormBorderStyle = FormBorderStyle.Sizable,
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                ClientSize = new Size(width, 360),
                MinimumSize = new Size(420, 260),
                Font = font,
            })
            using (var label = new Label
            {
                Text = prompt,
                Dock = DockStyle.Top,
                Height = 56,
                Padding = new Padding(12, 12, 12, 4),
            })
            using (var list = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                Margin = new Padding(12),
            })
            using (var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 44,
                Padding = new Padding(8),
            })
            using (var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 88 })
            using (var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 88 })
            {
                list.Items.AddRange(options);
                list.SelectedIndex = 0;
                list.DoubleClick += (s, e) => { if (list.SelectedIndex >= 0) form.DialogResult = DialogResult.OK; };
                // The list must hold focus for its native type-ahead: pressing an option's leading
                // letter selects it, Enter confirms — callers prefix options with accelerator letters.
                form.Shown += (s, e) => list.Focus();

                buttons.Controls.Add(cancel);
                buttons.Controls.Add(ok);
                form.Controls.Add(list);
                form.Controls.Add(buttons);
                form.Controls.Add(label);
                form.AcceptButton = ok;
                form.CancelButton = cancel;

                return form.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? list.SelectedIndex : -1;
            }
        }

        // The three message boxes are the add-in's OWN modal form, not Win32 MessageBox: MessageBox
        // always plays a system sound (the icon's, or the default beep without one) and nothing in
        // its API silences it. Same ownership (the invoking Excel window), same keys (Enter = the
        // affirmative button, Esc = the negative one), no sound, no icon.

        public static void Info(string message) =>
            Message(message, MessageBoxButtons.OK);

        public static void Info(string message, string caption, IWin32Window owner = null) =>
            Message(message, MessageBoxButtons.OK, caption, owner);

        public static bool Confirm(string message) =>
            Message(message, MessageBoxButtons.YesNo) == DialogResult.Yes;

        public static DialogResult YesNoCancel(string message) =>
            Message(message, MessageBoxButtons.YesNoCancel);

        public static bool OkCancel(string message, string caption, IWin32Window owner = null) =>
            Message(message, MessageBoxButtons.OKCancel, caption, owner) == DialogResult.OK;

        public static bool RetryCancel(string message, string caption, IWin32Window owner = null) =>
            Message(message, MessageBoxButtons.RetryCancel, caption, owner) == DialogResult.Retry;

        /// <summary>The silent modal box behind every variant; owner defaults to the Excel window.</summary>
        public static DialogResult Message(string message, MessageBoxButtons buttons, string caption = null,
            IWin32Window owner = null, bool defaultToNegative = false)
        {
            Font font = SystemFonts.MessageBoxFont;
            const int maxTextWidth = 640;
            Size text = TextRenderer.MeasureText(message ?? "", font, new Size(maxTextWidth, 0),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
            int width = Math.Max(360, Math.Min(text.Width, maxTextWidth) + 48);
            int textHeight = Math.Max(text.Height, font.Height) + 8;

            using (var form = new Form
            {
                Text = caption ?? Title,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                Font = font,
                ClientSize = new Size(width, textHeight + 24 + 52),
            })
            using (var label = new Label
            {
                Text = message ?? "",
                AutoSize = false,
                UseMnemonic = false,
                Location = new Point(16, 14),
                Size = new Size(width - 32, textHeight),
                TextAlign = ContentAlignment.TopLeft,
            })
            using (var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 52,
                Padding = new Padding(10, 10, 10, 10),
            })
            {
                Button affirmative = null, negative = null;
                void Add(string caption2, DialogResult result, bool isAffirmative, bool isNegative)
                {
                    var button = new Button { Text = caption2, DialogResult = result, Width = 88, Height = 28 };
                    panel.Controls.Add(button);   // right-to-left flow: the first added sits rightmost
                    if (isAffirmative) affirmative = button;
                    if (isNegative) negative = button;
                }

                switch (buttons)
                {
                    case MessageBoxButtons.YesNo:
                        Add("No", DialogResult.No, false, true);
                        Add("Yes", DialogResult.Yes, true, false);
                        break;
                    case MessageBoxButtons.YesNoCancel:
                        Add("Cancel", DialogResult.Cancel, false, true);
                        Add("No", DialogResult.No, false, false);
                        Add("Yes", DialogResult.Yes, true, false);
                        break;
                    case MessageBoxButtons.OKCancel:
                        Add("Cancel", DialogResult.Cancel, false, true);
                        Add("OK", DialogResult.OK, true, false);
                        break;
                    case MessageBoxButtons.RetryCancel:
                        Add("Cancel", DialogResult.Cancel, false, true);
                        Add("Retry", DialogResult.Retry, true, false);
                        break;
                    default:
                        Add("OK", DialogResult.OK, true, true);
                        break;
                }

                form.Controls.Add(label);
                form.Controls.Add(panel);
                form.AcceptButton = affirmative;
                form.CancelButton = negative;
                UI.GdiText.Apply(form);
                form.Shown += (s, e) => (defaultToNegative ? negative : affirmative)?.Focus();

                DialogResult result = form.ShowDialog(owner ?? ExcelWindow.Current);
                // Closing with the title-bar X yields Cancel; a Yes/No box has no Cancel — read it as No.
                if (buttons == MessageBoxButtons.YesNo && result == DialogResult.Cancel) return DialogResult.No;
                if (buttons == MessageBoxButtons.OK) return DialogResult.OK;
                return result;
            }
        }

        public static string OpenFile(string filter)
        {
            using (var d = new OpenFileDialog { Filter = filter, CheckFileExists = true })
                return d.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? d.FileName : null;
        }

        public static string[] OpenFiles(string filter)
        {
            using (var d = new OpenFileDialog
            {
                Filter = filter,
                CheckFileExists = true,
                Multiselect = true,
            })
                return d.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? d.FileNames : null;
        }

        public static string SelectFolder(string description)
        {
            using (var d = new FolderBrowserDialog
            {
                Description = description,
                ShowNewFolderButton = false,
            })
                return d.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? d.SelectedPath : null;
        }

        public static string SaveFile(string filter, string defaultName)
        {
            using (var d = new SaveFileDialog { Filter = filter, FileName = defaultName, OverwritePrompt = true })
                return d.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? d.FileName : null;
        }
    }
}
