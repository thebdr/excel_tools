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

        public static void Info(string message) =>
            MessageBox.Show(ExcelWindow.Current, message, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        public static bool Confirm(string message) =>
            MessageBox.Show(ExcelWindow.Current, message, Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        public static DialogResult YesNoCancel(string message) =>
            MessageBox.Show(ExcelWindow.Current, message, Title, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

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
